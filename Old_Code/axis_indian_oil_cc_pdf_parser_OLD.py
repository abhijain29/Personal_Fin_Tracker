import pdfplumber
import pandas as pd
import re
from pathlib import Path


DATE_REGEX = re.compile(r"\d{2}/\d{2}/\d{4}")
AMOUNT_REGEX = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{2}")


def parse_axis_indian_oil_cc_pdf(file_path):
    """
    Parse Axis Bank Indian Oil Credit Card PDF statement
    Returns: DataFrame with columns [period, date, description, amount, source, account]
    """
    rows = []
    in_transaction_section = False

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for raw_line in text.split("\n"):
                line = raw_line.strip()
                if not line:
                    continue

                # Detect start of transaction section
                if "DATE TRANSACTION DETAILS" in line or "Account Summary" in line:
                    in_transaction_section = True
                    continue
                
                # Detect end of transaction section
                if "End of Statement" in line:
                    in_transaction_section = False
                    continue
                
                # Only process lines in transaction section
                if not in_transaction_section:
                    continue

                # Must contain a date
                date_match = DATE_REGEX.search(line)
                if not date_match:
                    continue

                date = date_match.group(0)

                # Must end with Dr or Cr
                if not (line.endswith(" Dr") or line.endswith(" Cr")):
                    continue

                is_credit = line.endswith(" Cr")

                # Extract monetary amounts
                amounts = AMOUNT_REGEX.findall(line)
                if not amounts:
                    continue

                # Last amount before Dr/Cr is the transaction amount
                amount = float(amounts[-1].replace(",", ""))

                # For credit cards: Dr = spending (positive), Cr = payment/refund (negative)
                if is_credit:
                    amount = -abs(amount)

                # Build description - remove date, amount, and Dr/Cr
                desc = line
                desc = desc.replace(date, "")
                desc = re.sub(r"\s+Dr$", "", desc)
                desc = re.sub(r"\s+Cr$", "", desc)
                desc = re.sub(AMOUNT_REGEX, "", desc)  # Remove amount
                desc = re.sub(r"\s{2,}", " ", desc)  # Collapse multiple spaces

                description = desc.strip()
                if not description:
                    continue

                rows.append(
                    {
                        "date": date,
                        "description": description,
                        "amount": amount,
                    }
                )

    if not rows:
        print(f"⚠️ No transactions found in {file_path.name}")
        return None

    df = pd.DataFrame(rows)
    df["source"] = "CC"
    df["account"] = "Axis Indian Oil"
    df["date_parsed"] = pd.to_datetime(df["date"], dayfirst=True)
    df["period"] = df["date_parsed"].dt.strftime("%b-%Y")
    df = df.drop(columns=["date_parsed"])
    
    # Reorder columns to match ICICI format
    cols = ["period"] + [c for c in df.columns if c != "period"]
    df = df[cols]

    return df


# Test function
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python3 axis_indian_oil_cc_pdf_parser.py <path_to_pdf>")
        sys.exit(1)
    
    test_file = Path(sys.argv[1])
    
    if not test_file.exists():
        print(f"❌ File not found: {test_file}")
        sys.exit(1)
    
    print(f"Parsing {test_file.name}...")
    print("="*60)
    
    df = parse_axis_indian_oil_cc_pdf(test_file)
    
    if df is not None:
        print(f"\n✅ Successfully parsed {len(df)} transactions\n")
        print("📊 Preview:")
        print(df.to_string(index=False))
        print(f"\n💰 Total Amount: ₹{df['amount'].sum():,.2f}")
        
        # Save to CSV
        output_file = "parsed_axis_transactions.csv"
        df.to_csv(output_file, index=False)
        print(f"\n💾 Saved to {output_file}")
    else:
        print("\n❌ Parsing failed")
        sys.exit(1)
