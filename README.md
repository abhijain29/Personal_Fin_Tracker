# Personal Fin Tracker

Local parsers for:
1. Credit card statements (PDF)
2. Savings account statements (PDF)
3. Paytm UPI statements (Excel)

Outputs are generated as Excel workbooks under `Output/`.

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
- Main sheets:
  - `Credit card expenses`
  - `Credit card Reconciliation`
  - `Credit card summary`

2. Savings accounts:
- Script: `Pdf_Parser_Code/SB_Parser_Code/SB_Master_Parser.py`
- Input: `Bank_Statements/SB_Statements/`
- Output: `Output/SB_Monthly_Master_Tracker.xlsx`
- Main sheet:
  - `SB AC expenses`
- Axis logic includes:
  - parsing transaction block under `Statement for Account No...`
  - opening balance synthetic row:
    - previous month period
    - previous month end date
    - `Amount` blank
    - `Balance` from opening balance line

3. Paytm UPI:
- Script: `Pdf_Parser_Code/UPI_Parser_Code/Paytm_Parser.py`
- Input: `Bank_Statements/UPI Statements/Paytm_Statement_*.xlsx` (sheet `Passbook Payment History`)
- Mapping: `Reference Documents/Merchant category mapping.xlsx`
- Output: `Output/Paytm_transactions.xlsx`

## Paytm Mapping Rules
Uses sheet `PayTm_1` from mapping file:
- Classification columns:
  - `Tags`
  - `Description`
  - `Expense Type`
  - `Merchant Category`
- Account mapping columns:
  - `Your Account` -> `Value`

Matching behavior:
1. If reference `Description` is blank: match by `Tags` only.
2. If reference `Tags` is blank: match by `Description` only.
3. If both are present: both must match (partial match).
4. Fallback when no mapping match:
  - if source text starts with `Paid to` or `Money sent to`:
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

## Git Notes
- Keep statement files out of git.
- Current `.gitignore` already excludes:
  - `Bank_Statements/CC_Statements/`
  - `Bank_Statements/SB_Statements/`
  - `Bank_Statements/UPI Statements/`
  - `Output/`, `Logs/`, `Error/`, `__pycache__/`
