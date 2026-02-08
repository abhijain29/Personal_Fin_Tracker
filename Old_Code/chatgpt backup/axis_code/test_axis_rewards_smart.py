import sys
import os
# Add parent folder to Python path
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from axis_rewards_smart_parser import parse_axis_rewards_smart


print("=" * 70)
print("SMART AXIS REWARDS PARSER TEST")
print("=" * 70)


if len(sys.argv) < 2:
    print("Usage: python test_axis_rewards_smart.py <pdf1> <pdf2> ...")
    sys.exit(1)


all_transactions = []


for pdf in sys.argv[1:]:

    print("\nProcessing:", pdf)

    tx = parse_axis_rewards_smart(pdf)

    if not tx:
        print("⚠ No transactions found in:", pdf)
    else:
        print(f"✅ Extracted {len(tx)} transactions from {pdf}")
        all_transactions.extend(tx)


print("\n" + "=" * 70)
print("FINAL RESULT SUMMARY")
print("=" * 70)

print(f"Total transactions extracted from all PDFs: {len(all_transactions)}\n")

for t in all_transactions:
    print(t)
