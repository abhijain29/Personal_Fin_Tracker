import os
import re
from datetime import datetime, timedelta
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


def extract_amounts_with_decimals(text):
    if not text:
        return []
    return re.findall(r"\d{1,3}(?:,\d{2,3})*\.\d{2}", text)


def format_period_from_date(d):
    if not d:
        return "Unknown"
    return d.strftime("%b-%y")


def extract_period(text, bank_name):
    text = text or ""
    if bank_name == "Axis":
        m = re.search(r"period.*From\s*:?\s*(\d{2}-\d{2}-\d{4}).*To\s*:?\s*(\d{2}-\d{2}-\d{4})", text, re.I)
        if m:
            end = parse_date(m.group(2))
            return format_period_from_date(end)
    if bank_name == "HDFC":
        m = re.search(r"Statement as on\s*:\s*(\d{2}/\d{2}/\d{4})", text, re.I)
        if m:
            return format_period_from_date(parse_date(m.group(1)))
    if bank_name == "ICICI":
        m = re.search(r"period\s+([A-Za-z]+\s+\d{2},\s+\d{4})\s*-\s*([A-Za-z]+\s+\d{2},\s+\d{4})", text, re.I)
        if m:
            try:
                end = datetime.strptime(m.group(2), "%B %d, %Y").date()
                return format_period_from_date(end)
            except ValueError:
                pass
    if bank_name == "IDFC":
        m = re.search(r"STATEMENT PERIOD\s*:\s*\d{2}-[A-Z]{3}-\d{4}\s*to\s*(\d{2}-[A-Z]{3}-\d{4})", text, re.I)
        if m:
            try:
                end = datetime.strptime(m.group(1), "%d-%b-%Y").date()
                return format_period_from_date(end)
            except ValueError:
                pass
    if bank_name == "YES":
        m = re.search(r"Period Of\s*(\d{2}-[A-Za-z]{3}-\d{4})\s*to\s*(\d{2}-[A-Za-z]{3}-\d{4})", text, re.I)
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
    seen = set()
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "Axis")
        statement_start_date = None
        statement_end_date = None
        statement_started = False
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "Statement for Account No." in text:
                statement_started = True
                m = re.search(
                    r"Statement for Account No\..*?from\s+(\d{2}-\d{2}-\d{4})\s+to\s+(\d{2}-\d{2}-\d{4})",
                    text,
                    re.I | re.S,
                )
                if m:
                    statement_start_date = parse_date(m.group(1))
                    statement_end_date = parse_date(m.group(2))
                    if statement_end_date:
                        period = format_period_from_date(statement_end_date)
            for table in extract_tables(page):
                header = table[0] if table else []
                if not header:
                    continue
                header_text = " ".join([str(c) for c in header])
                is_old = "Tran Date" in header_text
                is_new = "Date" in header_text and "Transaction Details" in header_text and "Withdrawal" in header_text
                if not (is_old or is_new):
                    continue
                for row in table[1:]:
                    if not row or len(row) < 5:
                        continue
                    date = parse_date(row[0])
                    if not date:
                        continue
                    if is_old:
                        desc = (row[2] or "").replace("\n", " ").strip()
                        debit = parse_amount(row[3])
                        credit = parse_amount(row[4])
                        balance = parse_amount(row[5])
                    else:
                        desc = (row[1] or "").replace("\n", " ").strip()
                        debit = parse_amount(row[3])
                        credit = parse_amount(row[4])
                        balance = parse_amount(row[5])
                    amount = None
                    if debit is not None and debit != 0:
                        amount = -debit
                    elif credit is not None and credit != 0:
                        amount = credit
                    if amount is None:
                        continue
                    key = (date, desc, amount, balance)
                    if key not in seen:
                        seen.add(key)
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
            # Fallback: parse text lines for newer Axis format (tables without rows)
            if statement_started:
                lines = [l.strip() for l in text.split("\n") if l.strip()]
                current = None
                for line in lines:
                    if line.lower().startswith("opening balance"):
                        nums = extract_amounts_with_decimals(line)
                        if nums:
                            bal = parse_amount(nums[-1])
                            prev_date = None
                            prev_period = None
                            if statement_start_date:
                                prev_date = statement_start_date - timedelta(days=1)
                                prev_period = format_period_from_date(prev_date)
                            key = (prev_date, "Opening Balance", None, bal)
                            if key not in seen:
                                seen.add(key)
                                records.append(
                                    {
                                        "Period": prev_period or period,
                                        "Account": "Axis",
                                        "Date": prev_date,
                                        "Description": "Opening Balance",
                                        "Amount": None,
                                        "Balance": bal,
                                    }
                                )
                        continue
                    if re.match(r"^\d{2}-\d{2}-\d{4}\b", line):
                        nums = extract_amounts_with_decimals(line)
                        if len(nums) < 3:
                            continue
                        date = parse_date(line.split()[0])
                        if not date:
                            continue
                        withdrawal = parse_amount(nums[-3])
                        deposit = parse_amount(nums[-2])
                        balance = parse_amount(nums[-1])
                        desc = re.sub(r"^\d{2}-\d{2}-\d{4}\s+", "", line)
                        desc = re.sub(
                            r"\s+" + re.escape(nums[-3]) + r"\s+" + re.escape(nums[-2]) + r"\s+" + re.escape(nums[-1]) + r"$",
                            "",
                            desc,
                        ).strip()
                        amount = None
                        if withdrawal is not None and withdrawal != 0:
                            amount = -withdrawal
                        elif deposit is not None and deposit != 0:
                            amount = deposit
                        if amount is None:
                            continue
                        current = {
                            "Period": period,
                            "Account": "Axis",
                            "Date": date,
                            "Description": desc,
                            "Amount": amount,
                            "Balance": balance,
                        }
                        key = (date, desc, amount, balance)
                        if key not in seen:
                            seen.add(key)
                            records.append(current)
                    elif current and line.startswith("("):
                        # continuation line (e.g., reference no)
                        current["Description"] = (current["Description"] + " " + line).strip()
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
                        amount = -debit
                    elif credit is not None and credit != 0:
                        amount = credit
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
            text = page.extract_text() or ""
            lines = [l.strip() for l in text.split("\n") if re.match(r"^\d{2}-\d{2}-\d{4}\b", l.strip())]
            prev_balance = None
            for line in lines:
                if "B/F" in line:
                    # Opening balance row
                    nums = extract_amounts_with_decimals(line)
                    if nums:
                        prev_balance = parse_amount(nums[-1])
                    continue
                nums = extract_amounts_with_decimals(line)
                if len(nums) < 2:
                    continue
                date = parse_date(line.split()[0])
                if not date:
                    continue
                balance = parse_amount(nums[-1])
                amt = parse_amount(nums[-2])
                desc = line
                # Remove leading date and trailing amounts from description
                desc = re.sub(r"^\d{2}-\d{2}-\d{4}\s+", "", desc)
                desc = re.sub(r"\s+" + re.escape(nums[-2]) + r"\s+" + re.escape(nums[-1]) + r"$", "", desc).strip()
                if amt is None or balance is None:
                    continue
                if prev_balance is None:
                    prev_balance = balance
                    continue
                # Infer sign from balance movement
                amount = amt if balance > prev_balance else -amt
                prev_balance = balance
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
            text = page.extract_text() or ""
            lines = [l.strip() for l in text.split("\n") if re.match(r"^\d{2}\s+[A-Za-z]{3}\s+\d{2}\b", l.strip())]
            prev_balance = None
            for line in lines:
                dates = re.findall(r"\b\d{2}\s+[A-Za-z]{3}\s+\d{2}\b", line)
                if not dates:
                    continue
                date = parse_date(dates[-1]) or parse_date(dates[0])
                nums = extract_amounts_with_decimals(line)
                if len(nums) < 2:
                    continue
                balance = parse_amount(nums[-1])
                amt = parse_amount(nums[-2])
                desc = line
                # Remove leading date/time and trailing amounts
                desc = re.sub(r"^\d{2}\s+[A-Za-z]{3}\s+\d{2}\s+\d{2}:\d{2}\s+\d{2}\s+[A-Za-z]{3}\s+\d{2}\s+", "", desc)
                desc = re.sub(r"\s+" + re.escape(nums[-2]) + r"\s+" + re.escape(nums[-1]) + r"\s*CR?$", "", desc).strip()
                if amt is None or balance is None:
                    continue
                if prev_balance is None:
                    prev_balance = balance
                    continue
                amount = amt if balance > prev_balance else -amt
                prev_balance = balance
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
            text = page.extract_text() or ""
            lines = [l.strip() for l in text.split("\n") if re.match(r"^\d{2}/\d{2}/\d{4}\b", l.strip())]
            prev_balance = None
            for line in lines:
                if "B/F" in line:
                    nums = extract_amounts_with_decimals(line)
                    if nums:
                        prev_balance = parse_amount(nums[-1])
                    continue
                nums = extract_amounts_with_decimals(line)
                if len(nums) < 3:
                    continue
                date = parse_date(line.split()[0])
                if not date:
                    continue
                debit = parse_amount(nums[-3])
                credit = parse_amount(nums[-2])
                balance = parse_amount(nums[-1])
                desc = line
                desc = re.sub(r"^\d{2}/\d{2}/\d{4}\s+\d{2}/\d{2}/\d{4}\s+", "", desc)
                desc = re.sub(r"\s+" + re.escape(nums[-3]) + r"\s+" + re.escape(nums[-2]) + r"\s+" + re.escape(nums[-1]) + r"$", "", desc).strip()
                amount = None
                if debit is not None and debit != 0:
                    amount = -debit
                elif credit is not None and credit != 0:
                    amount = credit
                elif prev_balance is not None and balance is not None:
                    amount = (balance - prev_balance) if balance > prev_balance else -(prev_balance - balance)
                if amount is None:
                    continue
                prev_balance = balance if balance is not None else prev_balance
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
    import sys
    print("=" * 70)
    print("SAVINGS ACCOUNT MASTER PARSER")
    print("=" * 70)
    print(f"Scanning: {BASE_DIR}")

    all_records = []
    pdf_paths = []
    if len(sys.argv) > 1:
        pdf_paths = [sys.argv[1]]
    else:
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
    df["_sort_date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.sort_values(by=["Account", "_sort_date"], ascending=[True, True]).drop(columns=["_sort_date"])
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%d/%m/%Y")

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
