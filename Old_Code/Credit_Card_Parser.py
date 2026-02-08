import os
import csv
import pandas as pd
from pathlib import Path
from datetime import datetime

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
    "thank you",
    "payment-thank you"
]


def identify_parser(pdf_path):
    """
    Identify which parser to use based on folder path
    
    Handles 7 credit cards:
    1. ICICI Amazon
    2-4. Axis Indian Oil, Axis Select, Axis Rewards (unified parser)
    5. IDFC FIRST
    6. Uni Gold Card
    7. Uni Gold Card UPI
    """
    path_str = str(pdf_path).lower()

    # ICICI Amazon
    if "icici" in path_str and "amazon" in path_str:
        return parse_icici_cc_pdf, "ICICI Amazon"

    # All Axis cards use unified parser
    if "axis" in path_str:
        if "indian oil" in path_str:
            return parse_axis_pdf, "Axis Indian Oil"
        elif "select" in path_str:
            return parse_axis_pdf, "Axis Select"
        elif "rewards" in path_str:
            return parse_axis_pdf, "Axis Rewards"
        else:
            return parse_axis_pdf, "Axis Bank"

    # IDFC FIRST
    if "idfc" in path_str:
        return parse_idfc_cc_pdf, "IDFC FIRST"

    # Uni Gold Card UPI (check this before regular Uni Gold)
    if "uni gold" in path_str and "upi" in path_str:
        return parse_uni_gold_upi_cc_pdf, "Uni Gold UPI"

    # Uni Gold Card
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
        month = match.group(1).capitalize()
        # Look for year in parent folder names
        for part in pdf_path.parts:
            year_match = re.search(r"20\d{2}", part)
            if year_match:
                return f"{month}-{year_match.group(0)}"
        # Default to current year if not found
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
    Handles both dict and DataFrame row formats
    """
    # Handle both dict access and DataFrame column access
    description = tx.get("Description") or tx.get("description") or ""
    amount = tx.get("Amount") or tx.get("amount") or 0
    period = tx.get("Period") or tx.get("period") or ""
    bank = tx.get("Bank") or tx.get("bank") or ""
    account = tx.get("Account") or tx.get("account") or ""
    date = tx.get("Date") or tx.get("date") or ""
    source_file = tx.get("Source_File") or tx.get("source_file") or ""
    
    # Convert amount to float if it's a string
    if isinstance(amount, str):
        amount = float(amount.replace(",", ""))
    
    # Convert date to string if it's a date object
    if hasattr(date, 'strftime'):
        date = date.strftime("%d/%m/%Y")
    
    tx_type = classify_transaction(amount, str(description))
    
    return {
        "Period": period,
        "Bank": bank,
        "Account": account,
        "Date": str(date),
        "Description": str(description),
        "Amount": amount,
        "Type": tx_type,  # Payment, Expense, or Refund
        "Source_File": source_file
    }


def aggregate_transactions():
    """
    Main function to aggregate all credit card transactions
    
    Processes 7 credit cards:
    - ICICI Amazon
    - Axis Indian Oil
    - Axis Select
    - Axis Rewards
    - IDFC FIRST
    - Uni Gold Card
    - Uni Gold Card UPI
    """
    all_transactions = []
    stats = {
        "total_pdfs": 0,
        "successful": 0,
        "failed": 0,
        "total_transactions": 0
    }

    print("=" * 70)
    print("Starting Credit Card Aggregation")
    print("=" * 70)
    print(f"Scanning: {STATEMENTS_DIR}")
    print()

    for root, dirs, files in os.walk(STATEMENTS_DIR):
        for file in files:
            if not file.lower().endswith(".pdf"):
                continue

            stats["total_pdfs"] += 1
            pdf_path = Path(root) / file

            parser_func, bank = identify_parser(pdf_path)

            if not parser_func:
                print(f"⚠️ No parser mapped for: {pdf_path}")
                stats["failed"] += 1
                continue

            print(f"📄 Processing: {file} ({bank})")

            try:
                transactions = parser_func(str(pdf_path))

                if transactions is None:
                    print(f"   ⚠️ No transactions extracted")
                    stats["failed"] += 1
                    continue

                # Convert DataFrame to list of dicts
                if isinstance(transactions, pd.DataFrame):
                    if transactions.empty:
                        print(f"   ⚠️ No transactions extracted")
                        stats["failed"] += 1
                        continue
                    else:
                        transactions = transactions.to_dict(orient="records")

                if isinstance(transactions, list) and len(transactions) == 0:
                    print(f"   ⚠️ No transactions extracted")
                    stats["failed"] += 1
                    continue

                # Extract period from path/filename (if not already in transaction)
                extracted_period = extract_period_from_path(pdf_path)

                # Add metadata to each transaction
                for tx in transactions:
                    # Use period from parser if available, otherwise use extracted
                    if not tx.get("Period") and not tx.get("period"):
                        tx["Period"] = extracted_period
                    
                    # Add bank if not present
                    if not tx.get("Bank") and not tx.get("bank"):
                        tx["Bank"] = bank
                    
                    # Add source file
                    tx["Source_File"] = file
                    
                    all_transactions.append(tx)

                period_display = transactions[0].get("Period") or transactions[0].get("period") or extracted_period
                print(f"   ✅ Extracted {len(transactions)} transactions (Period: {period_display})")
                stats["successful"] += 1
                stats["total_transactions"] += len(transactions)

            except Exception as e:
                print(f"   ❌ Error: {e}")
                stats["failed"] += 1
                import traceback
                traceback.print_exc()

    print("\n" + "=" * 70)
    print("PROCESSING SUMMARY")
    print("=" * 70)
    print(f"PDFs found:             {stats['total_pdfs']}")
    print(f"Successfully parsed:    {stats['successful']}")
    print(f"Failed:                 {stats['failed']}")
    print(f"Total transactions:     {stats['total_transactions']}")

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
    print("SUMMARY BY TYPE")
    print("=" * 70)
    type_summary = df['Type'].value_counts()
    for tx_type, count in type_summary.items():
        total = df[df['Type'] == tx_type]['Amount'].sum()
        print(f"{tx_type:15} | Count: {count:>5} | Total: ₹{abs(total):>12,.2f}")
    
    print("\n" + "=" * 70)
    print("SUMMARY BY BANK")
    print("=" * 70)
    for bank in sorted(df['Bank'].unique()):
        count = len(df[df['Bank'] == bank])
        print(f"{bank:<25} | Transactions: {count:>4}")
    
    # Calculate net outstanding per bank (Expenses - Payments - Refunds)
    print("\n" + "=" * 70)
    print("NET OUTSTANDING BY BANK")
    print("=" * 70)
    print(f"{'Bank':<25} | {'Expenses':>12} | {'Payments':>12} | {'Refunds':>12} | {'Net':>12}")
    print("-" * 70)
    
    total_expenses = 0
    total_payments = 0
    total_refunds = 0
    
    for bank in sorted(df['Bank'].unique()):
        bank_df = df[df['Bank'] == bank]
        expenses = bank_df[bank_df['Type'] == 'Expense']['Amount'].sum()
        payments = abs(bank_df[bank_df['Type'] == 'Payment']['Amount'].sum())
        refunds = abs(bank_df[bank_df['Type'] == 'Refund']['Amount'].sum())
        net = expenses - payments - refunds
        
        total_expenses += expenses
        total_payments += payments
        total_refunds += refunds
        
        print(f"{bank:<25} | ₹{expenses:>11,.2f} | ₹{payments:>11,.2f} | ₹{refunds:>11,.2f} | ₹{net:>11,.2f}")
    
    print("-" * 70)
    total_net = total_expenses - total_payments - total_refunds
    print(f"{'TOTAL':<25} | ₹{total_expenses:>11,.2f} | ₹{total_payments:>11,.2f} | ₹{total_refunds:>11,.2f} | ₹{total_net:>11,.2f}")
    
    print("\n" + "=" * 70)
    print("✅ Aggregation complete!")
    print(f"📊 Output saved at: {OUTPUT_FILE}")
    print("=" * 70)


if __name__ == "__main__":
    aggregate_transactions()
