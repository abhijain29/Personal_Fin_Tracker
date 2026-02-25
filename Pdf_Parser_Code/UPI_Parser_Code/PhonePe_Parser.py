import os
import re
from datetime import datetime
from pathlib import Path
from difflib import SequenceMatcher

import pandas as pd
import pdfplumber

PROJECT_DIR = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker"
)
INPUT_DIR = os.path.join(PROJECT_DIR, "Bank_Statements", "UPI Statements")
OUTPUT_DIR = os.path.join(PROJECT_DIR, "Output")
MAPPING_FILE = os.path.join(PROJECT_DIR, "Reference Documents", "Merchant category mapping.xlsx")
LOG_FILE = os.path.join(PROJECT_DIR, "Logs", "File_Parser_log.txt")


def clean_text(value):
    if value is None:
        return ""
    s = str(value).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def append_log(file_name, output_file_name, error_text=""):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    line = (
        f"Type: UPI, File Name: {file_name}, Date: {now}, "
        f"Output file name: {output_file_name}, Errors: {error_text or 'None'}\n"
    )
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line)


def load_category_mapping():
    df = pd.read_excel(MAPPING_FILE, sheet_name="UPIs")
    rules = []
    for _, row in df.iterrows():
        kw = clean_text(row.get("Description", ""))
        if not kw:
            continue
        rules.append(
            (
                kw.lower(),
                clean_text(row.get("Expense Type", "")) or "Miscellaneous",
                clean_text(row.get("Merchant Category", "")) or "Miscellaneous",
                clean_text(row.get("Store Name", "")) or "Unknown",
            )
        )
    return rules


def classify(description, rules):
    desc = clean_text(description).lower()
    best = None
    best_score = 0.0
    for kw, exp_type, merch_cat, store_name in rules:
        if kw in desc or desc in kw:
            score = 1.0
        else:
            score = SequenceMatcher(None, kw, desc).ratio()
        if score > best_score:
            best_score = score
            best = (exp_type, merch_cat, store_name)
    if best and best_score >= 0.55:
        return best
    return "Miscellaneous", "Miscellaneous", "Unknown"


def parse_amount(amount_str):
    return float(str(amount_str).replace(",", "").strip())


def extract_transactions(pdf_path):
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            rows.extend(txt.splitlines())

    header_re = re.compile(
        r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}\s+(.+?)\s+(DEBIT|CREDIT)\s+₹\s*([0-9,]+(?:\.\d+)?)$"
    )

    records = []
    current = None
    for raw in rows:
        line = clean_text(raw)
        if not line:
            continue
        m = header_re.match(line)
        if m:
            if current:
                records.append(current)
            date_part = line.split(" ", 3)[0:3]
            date_str = " ".join(date_part).replace(",", ",")
            txn_date = datetime.strptime(date_str, "%b %d, %Y")
            desc = m.group(2)
            txn_type = m.group(3)
            amt = parse_amount(m.group(4))
            if txn_type.upper() == "DEBIT":
                amt = -amt
            current = {
                "Period": txn_date.strftime("%b-%Y"),
                "Date": txn_date.strftime("%d/%m/%Y"),
                "Description": desc,
                "Type": txn_type,
                "Amount": amt,
                "Account": "PhonePe",
                "Paid By": "",
                "Transaction ID": "",
                "UTR No": "",
            }
            continue

        if current is None:
            continue
        if line.startswith("Transaction ID "):
            current["Transaction ID"] = line.replace("Transaction ID ", "", 1).strip()
        elif line.startswith("UTR No. "):
            current["UTR No"] = line.replace("UTR No. ", "", 1).strip()
        elif line.startswith("Paid by "):
            current["Paid By"] = line.replace("Paid by ", "", 1).strip()
        elif line.startswith("Page ") or line.startswith("This is an automatically generated"):
            continue

    if current:
        records.append(current)
    return pd.DataFrame(records)


def format_output(writer, txns_df, summary_df, txns_sheet_name):
    workbook = writer.book
    ws_txn = writer.sheets[txns_sheet_name]
    ws_sum = writer.sheets["Categorized Txn Summary"]

    header_fmt = workbook.add_format({"bold": True, "bg_color": "#F4B183", "border": 1})
    cell_fmt = workbook.add_format({"border": 1})
    num_fmt = workbook.add_format({"num_format": "#,##0.00", "border": 1})

    def style(worksheet, dataframe):
        nrows, ncols = dataframe.shape
        for c, col in enumerate(dataframe.columns):
            worksheet.write(0, c, col, header_fmt)
            width = min(
                max(len(str(col)), dataframe[col].astype(str).map(len).max() if nrows else 0) + 2,
                60,
            )
            worksheet.set_column(c, c, width)
        for r in range(1, nrows + 1):
            for c, col in enumerate(dataframe.columns):
                val = dataframe.iloc[r - 1, c]
                if pd.isna(val) or (isinstance(val, str) and not val.strip()):
                    continue
                if col == "Amount":
                    worksheet.write_number(r, c, float(val), num_fmt)
                else:
                    worksheet.write(r, c, val, cell_fmt)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, nrows, ncols - 1)

    style(ws_txn, txns_df)
    style(ws_sum, summary_df)


def run(input_pdf):
    rules = load_category_mapping()
    txns_df = extract_transactions(input_pdf)
    if txns_df.empty:
        raise ValueError("No transactions parsed from PhonePe statement")

    txns_df[["Expense Type", "Merchant Category", "Store Name"]] = txns_df["Description"].apply(
        lambda x: pd.Series(classify(x, rules))
    )
    summary_df = txns_df[
        ["Period", "Account", "Expense Type", "Merchant Category", "Store Name", "Amount"]
    ].copy()
    period_label = pd.to_datetime(txns_df["Date"], dayfirst=True).max().strftime("%b'%y")
    output_name = f"PhonePe_{period_label}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_name)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        txns_df.to_excel(writer, sheet_name="PhonePe Transactions", index=False)
        summary_df.to_excel(writer, sheet_name="Categorized Txn Summary", index=False)
        format_output(writer, txns_df, summary_df, "PhonePe Transactions")
    return output_name, output_path, len(txns_df)


def main():
    files = sorted(Path(INPUT_DIR).glob("PhonePe*.pdf"))
    if not files:
        raise FileNotFoundError(f"No PhonePe PDF found in {INPUT_DIR}")
    input_pdf = str(files[0])
    in_name = Path(input_pdf).name
    try:
        out_name, out_path, rows = run(input_pdf)
        append_log(in_name, out_name, "")
        print(f"Input: {input_pdf}")
        print(f"Output: {out_path}")
        print(f"Rows: {rows}")
    except Exception as e:
        append_log(in_name, "", str(e))
        raise


if __name__ == "__main__":
    main()
