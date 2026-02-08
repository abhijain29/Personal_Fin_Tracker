import sys
from axis_unified_pdf_parser import parse_axis_pdf

print("=" * 70)
print("Testing Axis Select Credit Card Parser")
print("=" * 70)

if len(sys.argv) < 2:
    print("Usage: python test_axis_select_parser.py <pdf_path>")
    sys.exit(1)

pdf_path = sys.argv[1]

print(f"PDF Path: {pdf_path}")
print("=" * 70)

transactions = parse_axis_pdf(pdf_path)

if not transactions:
    print("⚠️ No transactions extracted from Axis Select PDF")
    sys.exit(1)

print(f"\n✅ Extracted {len(transactions)} transactions from Axis Select\n")

for tx in transactions:
    print(tx)
