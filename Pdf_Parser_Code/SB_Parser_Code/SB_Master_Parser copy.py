import os
import re
from datetime import datetime
from pathlib import Path

import pdfplumber
import pandas as pd

PROJECT_DIR = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker"
)

BASE_DIR = os.path.join(PROJECT_DIR, "Bank_Statements", "SB_Statements")
OUTPUT_FILE = os.path.join(PROJECT_DIR, "Output", "SB_Monthly_Master_Tracker.xlsx")


def parse_amount(value):
    if value is None:
        return None
    s = str(value).strip().replace(",", "")
    if not s:
        return None
    s = re.sub(r"\bCR\b|\bDR\b", "", s, flags=re.I).strip()
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(value):
    if not value:
        return None
    s = str(value).strip()
    for fmt in (
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d-%m-%y",
        "%d/%m/%y",
        "%d %b %y",
        "%d %b %Y",
        "%d-%b-%Y",
        "%d-%b-%y",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def format_period_from_date(d):
    if not d:
        return "Unknown"
    return d.strftime("%b-%y")


def extract_period(text, bank_name):
    text = text or ""
    if bank_name == "Axis":
        m = re.search(r"period.*From\\s*:?\\s*(\\d{2}-\\d{2}-\\d{4}).*To\\s*:?\\s*(\\d{2}-\\d{2}-\\d{4})", text, re.I)
        if m:
            end = parse_date(m.group(2))
            return format_period_from_date(end)
    if bank_name == "HDFC":
        m = re.search(r"Statement as on\\s*:\\s*(\\d{2}/\\d{2}/\\d{4})", text, re.I)
        if m:
            return format_period_from_date(parse_date(m.group(1)))
    if bank_name == "ICICI":
        m = re.search(r"period\\s+([A-Za-z]+\\s+\\d{2},\\s+\\d{4})\\s*-\\s*([A-Za-z]+\\s+\\d{2},\\s+\\d{4})", text, re.I)
        if m:
            try:
                end = datetime.strptime(m.group(2), "%B %d, %Y").date()
                return format_period_from_date(end)
            except ValueError:
                pass
    if bank_name == "IDFC":
        m = re.search(r"STATEMENT PERIOD\\s*:\\s*\\d{2}-[A-Z]{3}-\\d{4}\\s*to\\s*(\\d{2}-[A-Z]{3}-\\d{4})", text, re.I)
        if m:
            try:
                end = datetime.strptime(m.group(1), "%d-%b-%Y").date()
                return format_period_from_date(end)
            except ValueError:
                pass
    if bank_name == "YES":
        m = re.search(r"Period Of\\s*(\\d{2}-[A-Za-z]{3}-\\d{4})\\s*to\\s*(\\d{2}-[A-Za-z]{3}-\\d{4})", text, re.I)
        if m:
            try:
                end = datetime.strptime(m.group(2), "%d-%b-%Y").date()
                return format_period_from_date(end)
            except ValueError:
                pass
    return "Unknown"


def extract_tables(page):
    return page.extract_tables() or []


def parse_axis(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "Axis")
        for page in pdf.pages:
            for table in extract_tables(page):
                header = table[0] if table else []
                if not header or "Tran Date" not in " ".join([str(c) for c in header]):
                    continue
                for row in table[1:]:
                    if not row or len(row) < 6:
                        continue
                    date = parse_date(row[0])
                    if not date:
                        continue
                    desc = (row[2] or "").replace("\n", " ").strip()
                    debit = parse_amount(row[3])
                    credit = parse_amount(row[4])
                    balance = parse_amount(row[5])
                    amount = None
                    if debit is not None and debit != 0:
                        amount = debit
                    elif credit is not None and credit != 0:
                        amount = -credit
                    if amount is None:
                        continue
                    records.append(
                        {
                            "Period": period,
                            "Account": "Axis",
                            "Date": date,
                            "Description": desc,
                            "Amount": amount,
                            "Balance": balance,
                        }
                    )
    return records


def parse_hdfc(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "HDFC")
        for page in pdf.pages:
            for table in extract_tables(page):
                header = table[0] if table else []
                if not header or "Txn Date" not in " ".join([str(c) for c in header]):
                    continue
                for row in table[1:]:
                    if not row or len(row) < 5:
                        continue
                    date = parse_date(row[0])
                    if not date:
                        continue
                    desc = (row[1] or "").replace("\n", " ").strip()
                    debit = parse_amount(row[2])
                    credit = parse_amount(row[3])
                    balance = parse_amount(row[4])
                    amount = None
                    if debit is not None and debit != 0:
                        amount = debit
                    elif credit is not None and credit != 0:
                        amount = -credit
                    if amount is None:
                        continue
                    records.append(
                        {
                            "Period": period,
                            "Account": "HDFC",
                            "Date": date,
                            "Description": desc,
                            "Amount": amount,
                            "Balance": balance,
                        }
                    )
    return records


def parse_icici(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "ICICI")
        for page in pdf.pages:
            for table in extract_tables(page):
                header = table[0] if table else []
                if not header or "DATE" not in " ".join([str(c) for c in header]):
                    continue
                if "PARTICULARS" not in " ".join([str(c) for c in header]):
                    continue
                for row in table[1:]:
                    if not row or len(row) < 6:
                        continue
                    date = parse_date(row[0])
                    if not date:
                        continue
                    desc = (row[2] or "").replace("\n", " ").strip()
                    credit = parse_amount(row[3])
                    debit = parse_amount(row[4])
                    balance = parse_amount(row[5])
                    amount = None
                    if debit is not None and debit != 0:
                        amount = debit
                    elif credit is not None and credit != 0:
                        amount = -credit
                    if amount is None:
                        continue
                    records.append(
                        {
                            "Period": period,
                            "Account": "ICICI",
                            "Date": date,
                            "Description": desc,
                            "Amount": amount,
                            "Balance": balance,
                        }
                    )
    return records


def parse_idfc(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "IDFC")
        for page in pdf.pages:
            for table in extract_tables(page):
                header = table[0] if table else []
                if not header:
                    continue
                header_text = " ".join([str(c) for c in header])
                if "Transaction Details" not in header_text:
                    continue
                for row in table[1:]:
                    if not row or len(row) < 6:
                        continue
                    date = parse_date(row[1]) or parse_date(row[0])
                    if not date:
                        continue
                    desc = (row[2] or "").replace("\n", " ").strip()
                    debit = parse_amount(row[4])
                    credit = parse_amount(row[5])
                    balance = parse_amount(row[6]) if len(row) > 6 else parse_amount(row[5])
                    amount = None
                    if debit is not None and debit != 0:
                        amount = debit
                    elif credit is not None and credit != 0:
                        amount = -credit
                    if amount is None:
                        continue
                    records.append(
                        {
                            "Period": period,
                            "Account": "IDFC",
                            "Date": date,
                            "Description": desc,
                            "Amount": amount,
                            "Balance": balance,
                        }
                    )
    return records


def parse_yes(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "YES")
        for page in pdf.pages:
            for table in extract_tables(page):
                header = table[0] if table else []
                if not header:
                    continue
                header_text = " ".join([str(c) for c in header])
                if "Transaction" not in header_text or "Withdrawals" not in header_text:
                    continue
                for row in table[1:]:
                    if not row or len(row) < 6:
                        continue
                    date = parse_date(row[0])
                    if not date:
                        continue
                    desc = (row[2] or "").replace("\n", " ").strip()
                    debit = parse_amount(row[4])
                    credit = parse_amount(row[5])
                    balance = parse_amount(row[6]) if len(row) > 6 else parse_amount(row[5])
                    amount = None
                    if debit is not None and debit != 0:
                        amount = debit
                    elif credit is not None and credit != 0:
                        amount = -credit
                    if amount is None:
                        continue
                    records.append(
                        {
                            "Period": period,
                            "Account": "YES",
                            "Date": date,
                            "Description": desc,
                            "Amount": amount,
                            "Balance": balance,
                        }
                    )
    return records


def get_parser(file_path):
    path = file_path.lower()
    if "axis" in path:
        return parse_axis
    if "hdfc" in path:
        return parse_hdfc
    if "icici" in path:
        return parse_icici
    if "idfc" in path:
        return parse_idfc
    if "yes" in path:
        return parse_yes
    return None


def main():
    print("=" * 70)
    print("SAVINGS ACCOUNT MASTER PARSER")
    print("=" * 70)
    print(f"Scanning: {BASE_DIR}")

    all_records = []
    pdf_paths = []
    for root, _, files in os.walk(BASE_DIR):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdf_paths.append(os.path.join(root, f))

    for pdf_path in pdf_paths:
        print(f"\\n📄 Processing: {os.path.basename(pdf_path)}")
        parser = get_parser(pdf_path)
        if not parser:
            print("   ⚠️ No parser found for this file")
            continue
        try:
            records = parser(pdf_path)
            print(f"   ✅ Extracted {len(records)} transactions")
            all_records.extend(records)
        except Exception as e:
            print(f"   ❌ Failed: {e}")

    if not all_records:
        print("\\nNo transactions found.")
        return

    df = pd.DataFrame(all_records)
    df = df.sort_values(by=["Account", "Period", "Date"], ascending=[True, True, True])
    df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%d/%m/%Y")

    # SB AC expenses columns A-F (no Card Variant)
    df = df[["Period", "Account", "Date", "Description", "Amount", "Balance"]]

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="SB AC expenses", index=False)

    print("\\n======================================================================")
    print("✅ SB AGGREGATION COMPLETE")
    print("======================================================================")
    print(f"Total PDFs:            {len(pdf_paths)}")
    print(f"Total Transactions:    {len(df)}")
    print(f"Output File:           {OUTPUT_FILE}")
    print("======================================================================")


if __name__ == "__main__":
    main()
