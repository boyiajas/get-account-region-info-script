# Codex Context

Last updated: 2026-08-31

## Project Purpose

This folder contains local Python utilities for working with Legal Suite data and related Excel or PDF inputs.

The main use case is enriching account-based workbooks by looking up Legal Suite matters using the account number stored in `Matter.TheirRef`.

## Main Scripts

### `process_excel_accounts.py`

Primary workbook enrichment script.

What it does:
- Loads a source workbook from the current directory or an explicit path.
- Detects the account column per worksheet by header name where possible.
- Looks up each unique numeric account number in Legal Suite using `matter/get`.
- Tries `LEGALSUITE_API_KEY` first, then `LEGALSUITE_WC_API_KEY`.
- Selects the best matter match using preferred matter type, active/archive status, and record id.
- Fetches defendant information via `matparty/get`, `parlang/get`, and `party/get`.
- Fetches file notes via `filenote/get` using `FileNote.MatterID`.
- Returns only the latest 3 file notes, sorted locally by `date`, `time`, and `recordid` descending.
- Writes a new workbook with appended Legal Suite fields.

Current appended output fields:
- `file_reference`
- `branch`
- `defendant_first_name`
- `defendant_surname`
- `matter_description`
- `formatteddateinstructed`
- `casenumber`
- `laststageid`
- `last_file_note_1`
- `last_file_note_2`
- `last_file_note_3`

Current behavior notes:
- File notes are fetched by `Matter.RecordID`, not by file reference.
- The script currently returns `laststageid` only. If stage description is needed later, use `stage/get` with `Stage.RecordID = laststageid`.
- Output path defaults to `<source stem> with matter details.xlsx` for custom input files.

Typical usage:

```bash
python3 process_excel_accounts.py "Workbook.xlsx"
python3 process_excel_accounts.py "Workbook.xlsx" -o "Workbook output.xlsx"
```

### `extract_pdf_accounts.py`

Reads account numbers from a PDF, looks them up in Legal Suite, and exports basic matter results to Excel.

Main fields returned:
- account number
- file ref
- branch description

### `ftp_download_today.py`

Large operational script for Standard Bank FTP download and Legal Suite update workflows.

Key responsibilities:
- Downloads target files from FTP using `FTP_HOST`, `FTP_USER`, and `FTP_PASS`.
- Creates or updates matters.
- Updates matter extra screens via `matdocsc/get` and `matdocsc/update`.
- Uses `matter/get`, `matter/update`, `matter/store`, `party/store`, and related endpoints.
- Sets `dateinstructed` during matter creation.

Relevant note:
- This script contains extra-screen fields like `lastquickcomment`, but that is not the same as Legal Suite file notes from `filenote/get`.

### `absa_home_loan_extrascreen_update.py`

Processes ABSA Home Loan comment and PTP Excel files and updates Legal Suite extra screens for DBN/JHB and WC regions.

### `update_legalsuite_licenses.py`

Calls `/licensed/get` and updates specific worksheets in an Excel workbook with matched license information.

## Legal Suite API Conventions In This Repo

Preferred request style for the workbook scripts:
- `POST`
- `Authorization: Bearer <api key>`
- `Content-Type: application/x-www-form-urlencoded`
- `data={"where[]": "Table.Field,=,Value"}` for simple lookups
- `data=[("where[]", "..."), ("where[]", "...")]` for multiple filters

Examples already used in this repo:
- `matter/get` by `Matter.TheirRef`
- `filenote/get` by `FileNote.MatterID`
- `matparty/get` by `MatParty.MatterID`
- `parlang/get` by `ParLang.PartyID`
- `party/get` by `Party.RecordID`

When adding new Legal Suite queries to `process_excel_accounts.py`, keep the same client pattern used by `LegalSuiteClient._post`.

## Environment

Common variables:
- `LEGALSUITE_API_KEY`
- `LEGALSUITE_WC_API_KEY`

FTP-only variables:
- `FTP_HOST`
- `FTP_USER`
- `FTP_PASS`

The scripts load `.env` from the project directory if `env_config` is unavailable.

## Workbook Assumptions

For `process_excel_accounts.py`:
- Workbooks may contain multiple sheets.
- Header row is expected in row 1.
- Data starts from row 2.
- Account numbers must be numeric-only after normalization.
- Recognized account headers include:
  - `account_number`
  - `account number`
  - `accountnumber`
  - `account_no`
  - `account no`

## Operational Notes

- Generated Excel files in this folder are output artifacts, not source code.
- Do not commit `.env`, `.venv/`, `__pycache__/`, or large output workbooks unless explicitly intended.
- Network access is required for live Legal Suite lookups.
- For one-off verification, testing a single known matter is usually sufficient before running a full workbook.
