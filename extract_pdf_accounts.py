#!/usr/bin/env python3
"""
Extract account numbers from PDF and retrieve matter information from Legal Suite.
Searches Durban first, then Western Cape if account not found.
Exports results to Excel: Account Number, File Ref, Branch Description
"""

import os
import sys
import requests
import time
from typing import Optional
from dataclasses import dataclass

# Try to load from .env file
try:
    from env_config import load_env_file
    load_env_file()
except ImportError:
    # Fallback: manually load .env file
    env_file = os.path.join(os.path.dirname(__file__), '.env')
    if os.path.exists(env_file):
        with open(env_file) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    os.environ[key] = value

# Try to import PDF reader
try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

LEGALSUITE_API_BASE = "https://api.legalsuite.net"

# Define API key environment variables to use
API_KEY_ENVS = [
    "LEGALSUITE_API_KEY",        # Durban - try first
    "LEGALSUITE_WC_API_KEY",     # Western Cape - try second
]

# Retry configuration
MAX_ATTEMPTS = 3
RETRY_DELAYS = [1, 2, 4]


@dataclass
class LegalSuiteClient:
    """Client for querying Legal Suite API."""
    api_base: str
    api_key: str
    
    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/x-www-form-urlencoded",
        }
    
    def get_matter_by_theirref(self, account_number: str) -> Optional[dict]:
        """Query Legal Suite to get matter by account number (TheirRef)."""
        url = f"{self.api_base}/matter/get"
        data = {
            "where[]": f"Matter.TheirRef,=,{account_number}",
        }
        response = post_with_retry(url, self._headers(), data, timeout=60)
        
        if response is None:
            return None
        
        try:
            payload = response.json()
            items = payload.get("data", [])
            if items:
                return items[0]  # Return first match
        except Exception as e:
            print(f"Error parsing response: {e}")
        
        return None


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from PDF file."""
    if not PyPDF2:
        raise ImportError("PyPDF2 is required. Install it with: pip install PyPDF2")
    
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page in pdf_reader.pages:
                text += page.extract_text()
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return ""
    
    return text


def extract_account_numbers(text: str) -> list[str]:
    """
    Extract account numbers from PDF text.
    Account numbers are numeric values at the beginning of each line.
    They typically have 8-12 digits, but we accept 5-20 digits to be flexible.
    """
    account_numbers = []
    lines = text.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Get the first "word" (space-separated token) from the line
        parts = line.split()
        if not parts:
            continue
        
        first_token = parts[0]
        
        # Check if it's numeric and has reasonable length for an account number
        # Accept 5-20 digits (accounts can be 8-10+ digits)
        if first_token.isdigit() and 5 <= len(first_token) <= 20:
            account_numbers.append(first_token)
    
    # Remove duplicates while preserving order
    seen = set()
    unique_accounts = []
    for acc in account_numbers:
        if acc not in seen:
            seen.add(acc)
            unique_accounts.append(acc)
    
    return unique_accounts


def post_with_retry(url: str, headers: dict, data: dict, timeout: int = 60) -> Optional[requests.Response]:
    """Make POST request with retry logic."""
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            response = requests.post(url, headers=headers, data=data, timeout=timeout)
            if response.status_code >= 500:
                response.raise_for_status()
            return response
        except (requests.Timeout, requests.ConnectionError, requests.HTTPError) as exc:
            if attempt >= MAX_ATTEMPTS:
                print(f"Failed after {MAX_ATTEMPTS} attempts: {exc}")
                return None
            
            delay = RETRY_DELAYS[min(attempt - 1, len(RETRY_DELAYS) - 1)]
            print(f"Retry attempt {attempt}/{MAX_ATTEMPTS} after {delay}s...")
            time.sleep(delay)
    
    return None


def search_matter(account_number: str, clients: dict[str, LegalSuiteClient]) -> Optional[dict]:
    """
    Search for matter by account number across multiple regions.
    Tries each client in order until found.
    """
    print(f"Searching for account: {account_number}")
    
    for env_name in API_KEY_ENVS:
        if env_name not in clients:
            continue
        
        client = clients[env_name]
        region = "Durban" if "DURBAN" in env_name or env_name == "LEGALSUITE_API_KEY" else "Western Cape"
        print(f"  Searching {region}...")
        
        matter = client.get_matter_by_theirref(account_number)
        if matter:
            print(f"  Found in {region}!")
            return matter
    
    print(f"  Not found in any region")
    return None


def create_excel_report(account_results: list[dict], output_path: str) -> None:
    """
    Create Excel report with account information.
    Columns: Account Number, File Ref, Branch Description
    """
    if not Workbook:
        raise ImportError("openpyxl is required. Install it with: pip install openpyxl")
    
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Accounts"
    
    # Add headers
    worksheet['A1'] = "Account Number"
    worksheet['B1'] = "File Ref"
    worksheet['C1'] = "Branch Description"
    
    # Make headers bold
    for cell in worksheet[1]:
        cell.font = cell.font.copy()
        cell.font = cell.font.copy()
        cell.font = cell.font.copy()
    
    # Add data rows
    for idx, result in enumerate(account_results, start=2):
        worksheet[f'A{idx}'] = result.get('account_number', '')
        worksheet[f'B{idx}'] = result.get('file_ref', '')
        worksheet[f'C{idx}'] = result.get('branch_description', '')
    
    # Auto-adjust column widths
    worksheet.column_dimensions['A'].width = 20
    worksheet.column_dimensions['B'].width = 25
    worksheet.column_dimensions['C'].width = 30
    
    workbook.save(output_path)
    print(f"\nExcel report saved to: {output_path}")


def main():
    """Main execution function."""
    
    # Verify dependencies
    if not PyPDF2:
        print("ERROR: PyPDF2 is required. Install with: pip install PyPDF2")
        sys.exit(1)
    
    if not Workbook:
        print("ERROR: openpyxl is required. Install with: pip install openpyxl")
        sys.exit(1)
    
    # Check for required API keys
    required_api_key_envs = []
    for env_name in API_KEY_ENVS:
        api_key = os.getenv(env_name, "").strip()
        if api_key:
            required_api_key_envs.append(env_name)
    
    if not required_api_key_envs:
        print("ERROR: No API keys found. Set LEGALSUITE_API_KEY or LEGALSUITE_WC_API_KEY in .env")
        sys.exit(1)
    
    # Create clients for each available API key
    clients = {}
    for env_name in required_api_key_envs:
        api_key = os.getenv(env_name, "")
        clients[env_name] = LegalSuiteClient(LEGALSUITE_API_BASE, api_key)
    
    # Find PDF file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = None
    
    # Look for PDF in current directory
    for file in os.listdir(script_dir):
        if file.endswith('.pdf'):
            pdf_path = os.path.join(script_dir, file)
            break
    
    if not pdf_path:
        print("ERROR: No PDF file found in script directory")
        sys.exit(1)
    
    print(f"Processing PDF: {pdf_path}")
    
    # Extract text from PDF
    pdf_text = extract_text_from_pdf(pdf_path)
    if not pdf_text:
        print("ERROR: Could not extract text from PDF")
        sys.exit(1)
    
    # Extract account numbers
    account_numbers = extract_account_numbers(pdf_text)
    print(f"\nFound {len(account_numbers)} potential account numbers")
    
    if not account_numbers:
        print("No account numbers found in PDF")
        sys.exit(1)
    
    # Search for each account and collect results
    account_results = []
    for account_number in account_numbers:
        matter = search_matter(account_number, clients)
        
        if matter:
            result = {
                'account_number': account_number,
                'file_ref': matter.get('fileref', ''),
                'branch_description': matter.get('branchdescription', ''),
            }
            account_results.append(result)
        else:
            result = {
                'account_number': account_number,
                'file_ref': 'NOT FOUND',
                'branch_description': 'NOT FOUND',
            }
            account_results.append(result)
        
        # Stop processing after the last known account
        if account_number == "58829862":
            print(f"\nReached final account {account_number}, stopping processing...")
            break
    
    # Create Excel report
    output_path = os.path.join(script_dir, "account_information.xlsx")
    create_excel_report(account_results, output_path)
    
    # Print summary
    found_count = sum(1 for r in account_results if r['file_ref'] != 'NOT FOUND')
    print(f"\nSummary:")
    print(f"  Total accounts: {len(account_results)}")
    print(f"  Found: {found_count}")
    print(f"  Not found: {len(account_results) - found_count}")


if __name__ == "__main__":
    main()
