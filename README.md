# Personal Fin Tracker

Local parsers for:
1. Credit card statements (PDF)
2. Savings account statements (PDF)
3. UPI statements:
   - Paytm (Excel)
   - PhonePe (PDF)
   - MobiKwik (PDF)

Outputs are generated as Excel workbooks under `Output/`.

## Recent Updates (May 2026)
1. Credit card mapping moved to the consolidated workbook:
- Mapping file: `Reference Documents/Merchant category mapping.xlsx`
- Category rules sheet: `CC Merchant category mapping`
- Card/statement label sheet: `Label Mapping`
- Old CSV files removed:
  - `Reference Documents/Category_Mapping_v1.csv`
  - `Reference Documents/Outstanding_Label_Mapping.csv`

2. Credit card master parser was upgraded:
- detects card type from PDF content when filenames are generic
- resolves bank/card variants from `Label Mapping`
- extracts statement due date and reconciliation labels from workbook-driven mappings
- handles ICICI PDF text with doubled letters in due-date labels
- keeps cards with no available statement at the bottom of expense and reconciliation sheets
- writes `No STMT avaliable` for missing-statement rows and carries the dominant transaction period to those rows
- applies orange formatting only to actual headers in `Credit card summary Per card`
- adds HDFC Tata Neu parser support:
  - `Pdf_Parser_Code/CC_Parser/hdfc_tata_neu_cc_pdf_parser.py`
- improved IDFC and Uni Gold parsing for newer PDF formats

3. Savings account mapping moved to a dedicated workbook:
- Mapping file: `Reference Documents/SB Mapping.xlsx`
- Main rules sheet: `SB Mapping`
- Account-name override sheet: `Bank Name map`
- Transaction-date extraction sheet: `Trans Date`

4. SB master parser mapping engine was upgraded:
- bank-aware rule selection
- stronger keyword matching (normalized + token matching)
- derived money-in/money-out expense types
- merchant-category reversal for self-transfer directions
- archive folders named `Archive` or `Archived` are skipped during normal folder scans
- added/expanded support for HDFC, SBI, HSBC, and IndusInd formats in the master parser

5. Added bank-specific savings parsers:
- `Pdf_Parser_Code/SB_Parser_Code/axis_sb_parser.py`
- `Pdf_Parser_Code/SB_Parser_Code/icici_sb_parser.py`
- `Pdf_Parser_Code/SB_Parser_Code/idfc_sb_parser.py`

6. Axis and IDFC savings parsers can write using template workbooks:
- Template: `Reference Documents/template file/axis_sb_template.xlsx`
- Template: `Reference Documents/template file/idfc_sb_template.xlsx`
- Output: `Output/axis_summary.xlsx`
- Output: `Output/idfc_summary.xlsx`
- Template pivots are preserved during refresh.
- Important: second pivot should be maintained in template itself (auto-creation was removed after Excel desktop crash risk).

7. Added new UPI PDF parsers:
- `Pdf_Parser_Code/UPI_Parser_Code/PhonePe_Parser.py`
- `Pdf_Parser_Code/UPI_Parser_Code/MobiKwik_Parser.py`
- both write to `Output/` and log runs in `Logs/File_Parser_log.txt`

## Folder Structure
```text
Monthly_Fin_Tracker/
├── Bank_Statements/
│   ├── CC_Statements/
│   ├── SB_Statements/
│   └── UPI Statements/
├── Pdf_Parser_Code/
│   ├── CC_Parser/
│   ├── SB_Parser_Code/
│   └── UPI_Parser_Code/
├── Reference Documents/
├── Output/
├── Logs/
├── Error/
├── Archive/
└── Old_Code/
```

## Parsers
1. Credit cards:
- Script: `Pdf_Parser_Code/CC_Parser/Credit_Card_Master_Parser.py`
- Input: `Bank_Statements/CC_Statements/`
- Output: `Output/CC_Monthly_Master_Tracker.xlsx`
- Mapping: `Reference Documents/Merchant category mapping.xlsx`
- Main sheets:
  - `Credit card expenses`
  - `Credit card Reconciliation`
  - `Credit card summary`
  - `Credit card summary Per card`
- Missing-statement behavior:
  - cards configured in `Label Mapping` but not found in input PDFs are emitted with description `No STMT avaliable`
  - these rows use the dominant period from available transaction rows
  - these rows are sorted to the bottom of `Credit card expenses`
  - matching reconciliation rows are sorted to the bottom and show `No STMT avaliable` in `Payment Due Date`
- Supported parser modules include:
  - ICICI Amazon Pay
  - IDFC FIRST
  - Uni Gold
  - Uni Gold UPI
  - Axis Select / Indian Oil / Rewards
  - HDFC Tata Neu

2. Savings accounts:
- Script: `Pdf_Parser_Code/SB_Parser_Code/SB_Master_Parser.py`
- Input: `Bank_Statements/SB_Statements/`
- Output: `Output/SB_Monthly_Master_Tracker.xlsx`
- Mapping: `Reference Documents/SB Mapping.xlsx`
- Main sheet:
  - `SB AC expenses`
- Summary sheet:
  - `SB Categorized Summary`
- Normal folder scans skip `Archive/` and `Archived/` folders.
- Supported bank formats include:
  - Axis
  - HDFC
  - ICICI
  - IDFC
  - Yes
  - SBI
  - HSBC
  - IndusInd
- Axis logic includes:
  - parsing transaction block under `Statement for Account No...`
  - opening balance synthetic row:
    - previous month period
    - previous month end date
    - `Amount` blank
    - `Balance` from opening balance line

3. Bank-specific savings parsers:
- Axis:
  - Script: `Pdf_Parser_Code/SB_Parser_Code/axis_sb_parser.py`
  - Input: Axis PDFs in `Bank_Statements/SB_Statements/`
  - Output: `Output/axis_summary.xlsx`
  - Uses template: `Reference Documents/template file/axis_sb_template.xlsx`
  - Sheets refreshed: `Axis Transactions`, `Axis Categorized Summary`
- ICICI:
  - Script: `Pdf_Parser_Code/SB_Parser_Code/icici_sb_parser.py`
  - Input: ICICI PDFs in `Bank_Statements/SB_Statements/`
  - Output: `Output/icici_summary.xlsx`
  - Sheets: `ICICI Transactions`, `ICICI Categorized Summary`
- IDFC:
  - Script: `Pdf_Parser_Code/SB_Parser_Code/idfc_sb_parser.py`
  - Input: IDFC PDFs in `Bank_Statements/SB_Statements/`
  - Output: `Output/idfc_summary.xlsx`
  - Uses template: `Reference Documents/template file/idfc_sb_template.xlsx`
  - Sheets: `IDFC Transactions`, `IDFC Categorized Summary`

4. Paytm UPI:
- Script: `Pdf_Parser_Code/UPI_Parser_Code/Paytm_Parser.py`
- Input: `Bank_Statements/UPI Statements/*.xlsx` (sheet `Passbook Payment History`)
- Mapping: `Reference Documents/Merchant category mapping.xlsx`
- Output: `Output/Paytm_<Mon'YY>.xlsx` (example: `Paytm_Jan'26.xlsx`)
- Logging: appends parser run details to `Logs/File_Parser_log.txt`
- Archive move: currently disabled

5. PhonePe UPI:
- Script: `Pdf_Parser_Code/UPI_Parser_Code/PhonePe_Parser.py`
- Input: `Bank_Statements/UPI Statements/PhonePe*.pdf`
- Mapping: `Reference Documents/Merchant category mapping.xlsx` (sheet `UPIs`)
- Output: `Output/PhonePe_<Mon'YY>.xlsx`
- Output sheets:
  - `PhonePe Transactions`
  - `Categorized Txn Summary`

6. MobiKwik UPI:
- Script: `Pdf_Parser_Code/UPI_Parser_Code/MobiKwik_Parser.py`
- Input: `Bank_Statements/UPI Statements/MobiKwik*.pdf`
- Mapping: `Reference Documents/Merchant category mapping.xlsx` (sheet `UPIs`)
- Output: `Output/MobiKwik_<Mon'YY>.xlsx`
- Output sheets:
  - `MobiKwik Transactions`
  - `Categorized Txn Summary`

## Paytm Mapping Rules
Uses sheet `PayTm_1` from mapping file:
- Classification lookup columns (`A:F`):
  - `A = Tags` (source `Tags`)
  - `B = Description` (source `Transaction Details`)
  - `C = Other Transaction Details` (source `Other Transaction Details (UPI ID or A/c No)`)
  - `D = Your Account` (source `Your Account`)
  - `E = Expense Type`
  - `F = Merchant Category`
- Account mapping columns (`G:H`):
  - `Your Account` -> `Value`

Matching behavior:
1. Lookup scan is strict top-to-bottom.
2. For each row, only populated lookup fields among `A/B/C/D` are compared.
3. Comparison uses partial matching.
4. First matching row wins (no further checks after a match).
5. If no rule matches, defaults are:
  - `Expense Type = Miscellaneous`
  - `Merchant Category = cleaned Tags`

## Paytm Output Sheets
1. `Paytm Transactions`:
- Full source rows copied as-is from `Passbook Payment History`
- Numeric formatting fixes applied (no scientific notation for ID columns in Excel display)
- Header format:
  - bold
  - Orange Accent style fill
  - filter
  - freeze top row
- Borders applied only to filled cells

2. `Categorized Txn Summary`:
- Columns:
  - `Period`
  - `Account`
  - `Expense Type`
  - `Merchant Category`
  - `Amount`
- Pivot block starts at `H2`:
  - grouped by `Account`, `Expense Type`, `Merchant Category`
  - sums `Amount`
  - excludes `Account = Gold Coins`
  - includes total row at end
- Additional pivot block starts at `M2`:
  - grouped by `Expense Type`, `Merchant Category`
  - sums `Amount`
  - excludes `Account = Gold Coins`
  - includes total row at end

## Mapping Workbooks
1. Credit card and UPI mappings:
- File: `Reference Documents/Merchant category mapping.xlsx`
- Important sheets:
  - `CC Merchant category mapping`: credit-card keyword to expense/category/store mapping
  - `Label Mapping`: card variants, PDF card numbers, outstanding labels, due-date labels, and reconciliation labels
  - `UPIs`: PhonePe/MobiKwik mapping
  - `PayTm_1`: Paytm mapping
  - `UPI Name`: UPI file-name/account mapping

2. Savings account mappings:
- File: `Reference Documents/SB Mapping.xlsx`
- Important sheets:
  - `SB Mapping`: bank, keyword, mode, expense type, merchant category, store name, direction, and priority rules
  - `Bank Name map`: maps PDF text/customer names to output account names
  - `Trans Date`: bank-specific transaction-date/description field hints

3. Savings templates:
- `Reference Documents/template file/axis_sb_template.xlsx`
- `Reference Documents/template file/idfc_sb_template.xlsx`

## Run Commands
From project root:

```bash
cd /Users/abhishekjain/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker
```

Credit card parser:
```bash
python3 Pdf_Parser_Code/CC_Parser/Credit_Card_Master_Parser.py
```

Savings parser (all files):
```bash
python3 Pdf_Parser_Code/SB_Parser_Code/SB_Master_Parser.py
```

Savings parser (single file):
```bash
python3 Pdf_Parser_Code/SB_Parser_Code/SB_Master_Parser.py "Bank_Statements/SB_Statements/Axis.pdf"
```

Paytm parser:
```bash
python3 Pdf_Parser_Code/UPI_Parser_Code/Paytm_Parser.py "Bank_Statements/UPI Statements/Paytm_Statement_December_2025.xlsx"
```

Axis savings parser:
```bash
python3 Pdf_Parser_Code/SB_Parser_Code/axis_sb_parser.py
```

ICICI savings parser:
```bash
python3 Pdf_Parser_Code/SB_Parser_Code/icici_sb_parser.py
```

IDFC savings parser:
```bash
python3 Pdf_Parser_Code/SB_Parser_Code/idfc_sb_parser.py
```

PhonePe parser:
```bash
python3 Pdf_Parser_Code/UPI_Parser_Code/PhonePe_Parser.py
```

MobiKwik parser:
```bash
python3 Pdf_Parser_Code/UPI_Parser_Code/MobiKwik_Parser.py
```

## Git Notes
- Keep statement files out of git.
- Keep accidental personal-folder copies out of git.
- Current `.gitignore` already excludes:
  - `Bank_Statements/CC_Statements/`
  - `Bank_Statements/SB_Statements/`
  - `Bank_Statements/UPI Statements/`
  - `Output/`, `Logs/`, `Error/`, `__pycache__/`
  - `Reference Documents/OneDrive-Personal/`
