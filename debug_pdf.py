#!/usr/bin/env python3
"""
Debug script to examine PDF content and identify account number patterns
"""

import os
import sys

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

try:
    import PyPDF2
except ImportError:
    print("ERROR: PyPDF2 is required. Install with: pip install PyPDF2")
    sys.exit(1)

def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from PDF file."""
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            print(f"Total pages: {len(pdf_reader.pages)}\n")
            for page_num, page in enumerate(pdf_reader.pages[:3]):  # First 3 pages
                text += page.extract_text()
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return ""
    
    return text

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_path = None
    
    for file in os.listdir(script_dir):
        if file.endswith('.pdf'):
            pdf_path = os.path.join(script_dir, file)
            break
    
    if not pdf_path:
        print("ERROR: No PDF file found")
        sys.exit(1)
    
    print(f"Reading PDF: {pdf_path}\n")
    text = extract_text_from_pdf(pdf_path)
    
    print("=" * 80)
    print("FIRST 2000 CHARACTERS OF PDF TEXT:")
    print("=" * 80)
    print(text[:2000])
    print("=" * 80)
    
    # Show sample lines
    print("\nSAMPLE LINES FROM PDF:")
    print("=" * 80)
    lines = text.split('\n')
    for i, line in enumerate(lines[:50]):
        if line.strip():
            print(f"{i}: {line[:100]}")

if __name__ == "__main__":
    main()
