from pathlib import Path
from axis_bank_indian_oil_cc_pdf_parser import (
    parse_axis_bank_indian_oil_cc_pdf
)

print("=" * 70)
print("Testing Axis Bank Indian Oil Credit Card PDF Parser")
print("=" * 70)

PDF_PATH = Path(
    "/Users/abhishekjain/Library/CloudStorage/OneDrive-Personal/"
    "Personal/Finance/projects/Monthly_Fin_Tracker/"
    "CC statements/Axis Indian Oil/Dec.pdf"
)

print("PDF Path:", PDF_PATH)
print("PDF Exists:", PDF_PATH.exists())
print("=" * 70)

df = parse_axis_bank_indian_oil_cc_pdf(PDF_PATH)

if df is None or df.empty:
    print("❌ Parsing failed or no transactions found")
else:
    print(f"✅ Parsed {len(df)} transactions\n")
    print(df)
    df.to_csv("../Output/axis_indian_oil_cc_output.csv", index=False)
    print("\n📄 Output written to Output/axis_indian_oil_cc_output.csv")
