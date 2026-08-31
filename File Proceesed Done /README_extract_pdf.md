# PDF Account Extractor

This script extracts account numbers from a PDF file and retrieves their information from Legal Suite.

## Features
- Reads account numbers from PDF files
- Searches Legal Suite API with dual regional support:
  - **Durban** (Primary) - uses `LEGALSUITE_API_KEY`
  - **Western Cape** (Fallback) - uses `LEGALSUITE_WC_API_KEY`
- Extracts: Account Number, File Ref, Branch Description
- Exports results to Excel file

## Requirements

Install required Python packages:
```bash
pip install PyPDF2 openpyxl requests
```

## Setup

1. Ensure your `.env` file has both API keys configured:
```
LEGALSUITE_API_KEY=<durban_key>
LEGALSUITE_WC_API_KEY=<western_cape_key>
```

2. Place your PDF file in the same directory as the script

## Usage

Run the script:
```bash
python extract_pdf_accounts.py
```

The script will:
1. Read the PDF file in the current directory
2. Extract potential account numbers from the PDF text
3. Search Legal Suite for each account number (Durban first, then Western Cape)
4. Generate an Excel file named `account_information.xlsx` with the results

## Output

The generated Excel file contains:
- **Column A**: Account Number
- **Column B**: File Ref
- **Column C**: Branch Description

Accounts not found in either region will show "NOT FOUND" in the File Ref and Branch Description columns.

## Notes

- The account number is matched against the "TheirRef" field in Legal Suite
- If an account is not found in Durban, the script automatically searches Western Cape
- The script includes retry logic with exponential backoff for API requests
- Account extraction from PDF is generic - may need customization based on your specific PDF format
