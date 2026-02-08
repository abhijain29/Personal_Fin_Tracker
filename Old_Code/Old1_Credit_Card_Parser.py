import os
import csv
import pandas as pd
from pathlib import Path

# Import individual parsers
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

# Payment keywords to identify payment transactions
PAYMENT_KEYWORDS = [
    "payment received",
    "payment recieved",  # Common typo
    "si payment",
    "auto debit",
    "neft",
    "rtgs",
    "upi payment",
    "imps",
    "thank you"
]


def identify_parser(pdf_path):
    """
    Identify which parser to use based on folder path
    """
    path_str = str(pdf_path).lower()

    if "icici" in path_str:
        return parse_icici_cc_pdf, "ICICI Amazon"

    if "axis indian oil" in path_str:
        return parse_axis_indian_oil_cc_pdf, "Axis Indian Oil"
    
    if "axis select" in path_str or "axis rewards" in path_str:
        # You'll need to create this parser or use the Indian Oil one if format is same
        return parse_axis_indian_oil_cc_pdf, "Axis Select"

    if "idfc" in path_str:
        return parse_idfc_cc_pdf, "IDFC FIRST"

    if "uni gold upi" in path:
        return parse_uni_gold_upi_cc_pdf, "UNI Gold UPI"

    if "uni gold" in path_str:
        return parse_uni_gold_cc_pdf, "Uni Gold Card"

    return None, None


def extract_period_from_path(pdf_path):
    """
    Extract period from path or filename
    Supports formats: Dec-25, Jan-26, Aug.pdf, Nov.pdf, December-2025, etc.
    """
    import re
    
    # Try filename first (most reliable)
    filename = pdf_path.stem  # Gets filename without extension
    
    # Pattern 1: Month-YY or Month-YYYY (e.g., Dec-25, Jan-2026)
    match = re.search(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[- ]?(\d{2,4})", filename, re.IGNORECASE)
    if match:
        month = match.group(1).capitalize()
        year = match.group(2)
        # Convert 2-digit year to 4-digit
        if len(year) == 2:
            year = "20" + year
        return f"{month}-{year}"
    
    # Pattern 2: Just month name (e.g., Aug.pdf, Nov.pdf)
    match = re.search(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$", filename, re.IGNORECASE)
    if match:
        # Try to extract year from parent folder or use current year
        month = match.group(1).capitalize()
        # Look for year in parent folder names
        for part in pdf_path.parts:
            year_match = re.search(r"20\d{2}", part)
            if year_match:
                return f"{month}-{year_match.group(0)}"
        # Default to current year if not found
        from datetime import datetime
        return f"{month}-{datetime.now().year}"
    
    # Pattern 3: Full month name (e.g., December, November)
    full_months = {
        "january": "Jan", "february": "Feb", "march": "Mar", "april": "Apr",
        "may": "May", "june": "Jun", "july": "Jul", "august": "Aug",
        "september": "Sep", "october": "Oct", "november": "Nov", "december": "Dec"
    }
    
    for full, short in full_months.items():
        if full in filename.lower():
            # Look for year
            year_match = re.search(r"20\d{2}", filename)
            if year_match:
                return f"{short}-{year_match.group(0)}"
            return f"{short}-{datetime.now().year}"
    
    # If nothing found, try to extract from parent folders
    for part in pdf_path.parts:
        match = re.search(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[- ]?(\d{2,4})", part, re.IGNORECASE)
        if match:
            month = match.group(1).capitalize()
            year = match.group(2)
            if len(year) == 2:
                year = "20" + year
            return f"{month}-{year}"
    
    return "Unknown"


def is_payment_transaction(description):
    """
    Check if a transaction is a payment (not an expense)
    """
    desc_lower = description.lower()
    return any(keyword in desc_lower for keyword in PAYMENT_KEYWORDS)


def classify_transaction(amount, description):
    """
    Classify transaction type:
    - Payment: Money paid to bank (shown as credit in statement)
    - Expense: Money spent (shown as debit in statement)
    - Refund: Money refunded back (shown as credit in statement)
    """
    # Payments are typically negative amounts (credits) with payment keywords
    if amount < 0 and is_payment_transaction(description):
        return "Payment"
    
    # Other negative amounts are refunds
    elif amount < 0:
        return "Refund"
    
    # Positive amounts are expenses
    else:
        return "Expense"


def normalize_record(tx):
    """
    Normalize transaction record and add transaction type
    """
    description = tx.get("Description") or tx.get("description") or ""
    amount = tx.get("Amount") or tx.get("amount") or 0
    
    # Convert amount to float if it's a string
    if isinstance(amount, str):
        amount = float(amount.replace(",", ""))
    
    tx_type = classify_transaction(amount, description)
    
    return {
        "Period": tx.get("Period") or tx.get("period") or "",
        "Bank": tx.get("Bank") or tx.get("bank") or "",
        "Account": tx.get("Account") or tx.get("account") or "",
        "Date": tx.get("Date") or tx.get("date") or "",
        "Description": description,
        "Amount": amount,
        "Type": tx_type,  # Payment, Expense, or Refund
        "Source_File": tx.get("Source_File") or tx.get("source_file") or ""
    }


def aggregate_transactions():
    """
    Main function to aggregate all credit card transactions
    """
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

            print(f"\nProcessing: {pdf_path.name} ({bank})")

            try:
                transactions = parser_func(str(pdf_path))

                if transactions is None:
                    print("  ⚠️ No transactions extracted")
                    continue

                # Convert DataFrame to list of dicts
                if isinstance(transactions, pd.DataFrame):
                    if transactions.empty:
                        print("  ⚠️ No transactions extracted")
                        continue
                    else:
                        transactions = transactions.to_dict(orient="records")

                if isinstance(transactions, list) and len(transactions) == 0:
                    print("  ⚠️ No transactions extracted")
                    continue

                # Extract period from path/filename
                period = extract_period_from_path(pdf_path)

                # Add metadata to each transaction
                for tx in transactions:
                    tx["Bank"] = bank
                    tx["Period"] = period
                    tx["Source_File"] = file
                    all_transactions.append(tx)

                print(f"  ✅ Extracted {len(transactions)} transactions (Period: {period})")

            except Exception as e:
                print(f"  ❌ Error processing {file}: {e}")
                import traceback
                traceback.print_exc()

    if not all_transactions:
        print("\n❌ No transactions extracted from any PDFs")
        return

    write_output_csv(all_transactions)


def write_output_csv(data):
    """
    Write aggregated transactions to CSV with proper classification
    """
    print("\n" + "=" * 70)
    print("Writing final output CSV...")
    
    fieldnames = [
        "Period",
        "Bank",
        "Account",
        "Date",
        "Description",
        "Amount",
        "Type",  # Payment, Expense, or Refund
        "Source_File"
    ]

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for row in data:
            clean_row = normalize_record(row)
            writer.writerow(clean_row)

    # Print summary statistics
    df = pd.read_csv(OUTPUT_FILE)
    
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"Total transactions:     {len(df)}")
    print(f"\nBy Type:")
    print(df['Type'].value_counts().to_string())
    print(f"\nBy Bank:")
    print(df['Bank'].value_counts().to_string())
    
    # Calculate net outstanding per bank (Expenses - Payments - Refunds)
    print(f"\n" + "=" * 70)
    print("NET OUTSTANDING BY BANK")
    print("=" * 70)
    for bank in df['Bank'].unique():
        bank_df = df[df['Bank'] == bank]
        expenses = bank_df[bank_df['Type'] == 'Expense']['Amount'].sum()
        payments = abs(bank_df[bank_df['Type'] == 'Payment']['Amount'].sum())
        refunds = abs(bank_df[bank_df['Type'] == 'Refund']['Amount'].sum())
        net = expenses - payments - refunds
        print(f"{bank:20} | Expenses: ₹{expenses:>10,.2f} | Payments: ₹{payments:>10,.2f} | Refunds: ₹{refunds:>10,.2f} | Net: ₹{net:>10,.2f}")
    
    print("\n" + "=" * 70)
    print("✅ Aggregation complete!")
    print(f"Output saved at: {OUTPUT_FILE}")
    print("=" * 70)


if __name__ == "__main__":
    aggregate_transactions()
