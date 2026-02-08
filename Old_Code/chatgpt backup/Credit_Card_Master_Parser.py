import os
import csv

# Correct imports – matching your actual parser file names
from icici_cc_pdf_parser import extract_icici_transactions
from idfc_cc_pdf_parser import extract_idfc_transactions
from uni_gold_cc_pdf_parser import parse_uni_gold_cc_pdf
from uni_gold_upi_cc_pdf_parser import parse_uni_gold_upi_cc_pdf
from axis_unified_pdf_parser import parse_axis_pdf

# OCR smart parser only for Axis Rewards
from axis_code.axis_rewards_smart_parser import parse_axis_rewards_smart


BASE_DIR = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker/CC statements"
)

OUTPUT_FILE = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker/Output/All_Credit_Card_Transactions.csv"
)


def get_parser(file_path):

    path = file_path.lower()

    if "icici" in path:
        return extract_icici_transactions

    if "idfc" in path:
        return extract_idfc_transactions

    if "uni gold upi" in path:
        return parse_uni_gold_upi_cc_pdf

    if "uni gold" in path:
        return parse_uni_gold_cc_pdf

    if "axis rewards" in path:
        return parse_axis_rewards_smart

    if "axis" in path:
        return parse_axis_pdf

    return None


def normalize(records):

    valid = []

    for r in records:

        if not isinstance(r, dict):
            continue

        if not r.get("Date"):
            continue

        # Ensure all required keys exist
        if "Period" not in r:
            r["Period"] = ""

        if "Type" not in r:
            r["Type"] = "Dr"

        valid.append(r)

    return valid


def aggregate():

    print("\nMASTER CREDIT CARD AGGREGATOR\n")

    all_records = []

    for root, dirs, files in os.walk(BASE_DIR):

        for file in files:

            if not file.lower().endswith(".pdf"):
                continue

            file_path = os.path.join(root, file)

            print(f"\nProcessing: {file_path}")

            parser = get_parser(file_path)

            if not parser:
                print("⚠ No parser found")
                continue

            try:
                records = parser(file_path)

                records = normalize(records)

                print(f"✔ Extracted {len(records)} transactions")

                all_records.extend(records)

            except Exception as e:
                print(f"❌ Error processing {file}: {e}")

    print("\nWriting final CSV output...\n")

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:

        writer = csv.DictWriter(
            f,
            fieldnames=["Period", "Account", "Date", "Description", "Amount", "Type"]
        )

        writer.writeheader()

        for r in all_records:
            writer.writerow(r)

    print("AGGREGATION COMPLETE")
    print(f"Total Transactions: {len(all_records)}")
    print(f"Output File: {OUTPUT_FILE}\n")


if __name__ == "__main__":
    aggregate()
