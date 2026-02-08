import sys
from axis_unified_pdf_parser import parse_axis_pdf

if len(sys.argv) < 2:
    print("Usage: python test_axis_unified_parser.py <pdf_path>")
    sys.exit(1)

pdf_path = sys.argv[1]

print("=" * 70)
print("Testing Unified Axis Bank Parser")
print("=" * 70)

print("PDF Path:", pdf_path)

transactions = parse_axis_pdf(pdf_path)

if transactions:
    print(f"\nFound {len(transactions)} transactions\n")

    for tx in transactions:
        print(tx)
else:
    print("⚠️ No transactions found")
