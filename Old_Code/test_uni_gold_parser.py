#test_uni_gold_parser.py
import sys
from pathlib import Path
from uni_gold_cc_pdf_parser import parse_uni_gold_cc_pdf

print("=" * 70)
print("Testing Uni Gold Credit Card PDF Parser")
print("=" * 70)

if len(sys.argv) < 2:
    print("❌ Please provide PDF path")
    sys.exit(1)

pdf_path = Path(sys.argv[1]).expanduser().resolve()

print(f"PDF Path: {pdf_path}")
print(f"PDF Exists: {pdf_path.exists()}")
print("=" * 70)

df = parse_uni_gold_cc_pdf(pdf_path)

if df.empty:
    print("⚠️ No transactions found")
    print("❌ Parsing failed or no transactions extracted")
    sys.exit(1)

print(f"✅ Parsed {len(df)} transactions\n")
print(df)

output_path = Path("../Output/uni_gold_cc_output.csv").resolve()
output_path.parent.mkdir(parents=True, exist_ok=True)

df.to_csv(output_path, index=False)

print("\n📄 Output saved to:")
print(output_path)
