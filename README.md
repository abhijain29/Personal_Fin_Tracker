# Personal Finance Tracker

A local-first parser that consolidates credit card statements from multiple banks into a single Excel workbook with reconciliation and category summaries.

## What It Does
- Parses PDFs for supported credit cards
- Normalizes DR/CR signs
- Extracts statement period and outstanding amounts
- Categorizes transactions using your mapping file
- Outputs a multi-sheet Excel file with reconciliation and summary views

## Folder Layout
```
Monthly_Fin_Tracker/
├── CC_Parser_Code/                 # Python parsers + master script
├── CC statements/                  # Input PDFs (bank/card subfolders)
├── Reference Documents/            # Mapping CSVs used during parsing
├── Output/                          # Generated Excel output
├── Logs/                            # Runtime logs (ignored by git)
├── Error/                           # Failed PDFs (ignored by git)
├── Archive/                         # Processed PDFs (ignored by git)
└── Old_Code/                        # Archived code
```

## Key Files
- `CC_Parser_Code/Credit_Card_Master_Parser.py` — main entry point
- `Reference Documents/Merchant category mapping.csv` — categorization map
- `Reference Documents/Outstanding_Label_Mapping.csv` — outstanding label map

## Run
```bash
cd /Users/abhishekjain/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker
python3 CC_Parser_Code/Credit_Card_Master_Parser.py
```

## Output
- Excel file written to:
  - `Output/CC_Monthly_Master_Tracker.xlsx`
- Sheets include:
  - `Credit card expenses`
  - `Credit card Reconciliation`
  - `Credit card summary`

## Notes
- `CC statements/`, `Output/`, `Logs/`, `Error/`, `Archive/` are ignored by git.
- Update the mapping CSVs to refine categorization without changing code.

## Supported Cards
- Axis (Rewards, Select, Indian Oil)
- ICICI (Amazon Pay)
- IDFC (First Select)
- Uni (Gold, Gold UPI)

---
If you add a new bank/card, place PDFs under `CC statements/<Bank>/<Card>` and add a parser in `CC_Parser_Code/`.
