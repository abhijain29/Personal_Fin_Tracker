import os
import csv
import pandas as pd
from pathlib import Path

# Import individual parsers (YOUR exact names)
from icici_cc_pdf_parser import parse_icici_cc_pdf
from axis_unified_pdf_parser import parse_axis_pdf
from idfc_cc_pdf_parser import parse_idfc_cc_pdf
from uni_gold_cc_pdf_parser import parse_uni_gold_cc_pdf
from uni_gold_upi_cc_pdf_parser import parse_uni_gold_upi_cc_pdf


BASE_DIR = Path(__file__).resolve().parent.parent
STATEMENTS_DIR = BASE_DIR / "CC statements"
OUTPUT_DIR = BASE_DIR / "Output"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "credit_card_transactions.csv"


def identify_parser(pdf_path):

    path = str(pdf_path).lower()

    if "icici" in path:
        return parse_icici_cc_pdf, "ICICI"

    if "axis" in path:
        return parse_axis_pdf, "Axis"

    if "idfc" in path:
        return parse_idfc_cc_pdf, "IDFC"

    if "uni gold upi" in path:
        return parse_uni_gold_upi_cc_pdf, "UNI Gold UPI"

    if "uni gold" in path:
        return parse_uni_gold_cc_pdf, "UNI Gold"

    return None, None


def extract_period_from_path(pdf_path):
    """
    Extract period only if it matches Month-Year pattern like:
    Dec-25, Jan-26, Aug-2025 etc.
    """

    parts = pdf_path.parts

    import re
    pattern = r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[- ]?\d{2,4}"

    for part in parts:
        if re.search(pattern, part, re.IGNORECASE):
            return part

    return "Unknown"



def normalize_record(tx):

    return {
        "Period": tx.get("Period") or tx.get("period") or "",
        "Bank": tx.get("Bank") or tx.get("bank") or "",
        "Account": tx.get("Account") or tx.get("account") or "",
        "Date": tx.get("Date") or tx.get("date") or "",
        "Description": tx.get("Description") or tx.get("description") or "",
        "Amount": tx.get("Amount") or tx.get("amount") or "",
        "Type": tx.get("Type") or tx.get("type") or "",
        "Source_File": tx.get("Source_File") or tx.get("source_file") or ""
    }


def aggregate_transactions():

    all_transactions = []

    print("=" * 70)
    print("Starting Credit Card Aggregation")
    print("=" * 70)

    for root, dirs, files in os.walk(STATEMENTS_DIR):
        for file in files:
            if not file.lower().endswith(".pdf"):
                continue

            pdf_path = Path(root) / file

            parser_func, bank = identify_parser(pdf_path)

            if not parser_func:
                print(f"⚠️ No parser mapped for: {pdf_path}")
                continue

            print(f"\nProcessing: {pdf_path}")

            try:
                transactions = parser_func(str(pdf_path))

                if transactions is None:
                    print("⚠️ No transactions extracted")
                    continue

                if isinstance(transactions, pd.DataFrame):
                    if transactions.empty:
                        print("⚠️ No transactions extracted")
                        continue
                    else:
                        transactions = transactions.to_dict(orient="records")

                if isinstance(transactions, list) and len(transactions) == 0:
                    print("⚠️ No transactions extracted")
                    continue

                period = extract_period_from_path(pdf_path)

                for tx in transactions:
                    tx["Bank"] = bank
                    tx["Period"] = period
                    tx["Source_File"] = file
                    all_transactions.append(tx)

                print(f"✔ Extracted {len(transactions)} transactions")

            except Exception as e:
                print(f"❌ Error processing {file}: {e}")

    if not all_transactions:
        print("\nNo transactions extracted from any PDFs")
        return

    write_output_csv(all_transactions)


def write_output_csv(data):

    print("\nWriting final output CSV...")

    fieldnames = [
        "Period",
        "Bank",
        "Account",
        "Date",
        "Description",
        "Amount",
        "Type",
        "Source_File"
    ]

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:

        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for row in data:
            clean_row = normalize_record(row)
            writer.writerow(clean_row)

    print("\n==============================================")
    print("✅ Aggregation complete!")
    print(f"Output saved at: {OUTPUT_FILE}")
    print("==============================================")


if __name__ == "__main__":
    aggregate_transactions()
