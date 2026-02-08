import pdfplumber
from pathlib import Path

PDF_PATH = Path(
    "/Users/abhishekjain/Library/CloudStorage/OneDrive-Personal/"
    "Personal/Finance/projects/Monthly_Fin_Tracker/CC statements/Axis Indian Oil/Dec.pdf"
)

print("PDF exists:", PDF_PATH.exists())
print("PDF path:", PDF_PATH)
print("=" * 60)

with pdfplumber.open(PDF_PATH) as pdf:
    page = pdf.pages[0]

    print("\n=== extract_text() ===")
    text = page.extract_text()
    print(text)

    print("\n=== extract_table() ===")
    table = page.extract_table()
    print(table)

    print("\n=== extract_tables() ===")
    tables = page.extract_tables()
    print(tables)
