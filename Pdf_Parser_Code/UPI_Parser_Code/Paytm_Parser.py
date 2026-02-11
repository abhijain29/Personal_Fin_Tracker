import os
import re
import sys
from pathlib import Path

import pandas as pd

PROJECT_DIR = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker"
)

DEFAULT_INPUT_DIR = os.path.join(PROJECT_DIR, "Bank_Statements", "UPI Statements")
DEFAULT_MAPPING_FILE = os.path.join(PROJECT_DIR, "Reference Documents", "Merchant category mapping.xlsx")
OUTPUT_FILE = os.path.join(PROJECT_DIR, "Output", "Paytm_transactions.xlsx")


def clean_text(value):
    if value is None:
        return ""
    s = str(value).strip()
    s = re.sub(r"[\W_]+", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def norm(value):
    return clean_text(value).lower()


def parse_amount(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s:
        return None
    s = s.replace(",", "")
    s = s.replace("\u2212", "-")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        return float(s)
    except ValueError:
        return None


def load_source(source_path):
    df = pd.read_excel(source_path, sheet_name="Passbook Payment History")
    required = ["Date", "Transaction Details", "Tags"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Missing required column in input: {col}")
    return df


def load_paytm_mapping(mapping_path):
    # Preferred format: Excel with sheet name "PayTm", columns A-F
    patterns = []
    account_map = []

    if mapping_path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        mdf = pd.read_excel(mapping_path, sheet_name="PayTm")
        for _, row in mdf.iterrows():
            key = str(row.iloc[0]).strip() if len(row) > 0 and pd.notna(row.iloc[0]) else ""
            exp = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
            merch = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
            acc_in = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ""
            acc_out = str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else ""
            if key:
                patterns.append((key, exp, merch))
            if acc_in and acc_out:
                account_map.append((acc_in, acc_out))
        return patterns, account_map

    # Fallback format: CSV with columns Keyword Pattern / Expense Type / Merchant Category
    cdf = pd.read_csv(mapping_path, encoding="utf-8-sig")
    key_col = "Keyword Pattern" if "Keyword Pattern" in cdf.columns else cdf.columns[0]
    exp_col = "Expense Type" if "Expense Type" in cdf.columns else (cdf.columns[1] if len(cdf.columns) > 1 else None)
    mer_col = "Merchant Category" if "Merchant Category" in cdf.columns else (cdf.columns[2] if len(cdf.columns) > 2 else None)

    for _, row in cdf.iterrows():
        key = str(row.get(key_col, "")).strip()
        if not key:
            continue
        exp = str(row.get(exp_col, "")).strip() if exp_col else ""
        merch = str(row.get(mer_col, "")).strip() if mer_col else ""
        patterns.append((key, exp, merch))

    return patterns, account_map


def match_pattern(txn_text, patterns):
    t = norm(txn_text)
    best = None
    best_len = -1
    for key, exp, merch in patterns:
        k = norm(key)
        if not k:
            continue
        if k in t or t in k:
            if len(k) > best_len:
                best = (key, exp, merch)
                best_len = len(k)
    return best


def derive_account_by_source(source_value, account_map, fallback):
    s = norm(source_value)
    best = None
    best_len = -1
    for src, dst in account_map:
        nsrc = norm(src)
        if not nsrc:
            continue
        if nsrc in s or s in nsrc:
            if len(nsrc) > best_len:
                best = dst
                best_len = len(nsrc)
    return best if best else fallback


def parse_paytm(source_path, mapping_path=DEFAULT_MAPPING_FILE):
    sdf = load_source(source_path)
    patterns, account_map = load_paytm_mapping(mapping_path)
    file_fallback_account = Path(source_path).stem

    out_rows = []
    for _, row in sdf.iterrows():
        date_val = pd.to_datetime(row["Date"], dayfirst=True, errors="coerce")
        if pd.isna(date_val):
            continue
        period = date_val.strftime("%b-%Y")

        txn = str(row.get("Transaction Details", "")).strip()
        amount = parse_amount(row.get("Amount", None))
        tags = clean_text(row.get("Tags", ""))

        source_account = str(row.get("Your Account", "")).strip()
        account_value = derive_account_by_source(source_account, account_map, file_fallback_account)

        m = match_pattern(txn, patterns)
        if m:
            desc, exp_type, merch_cat = m
        else:
            txn_low = txn.lower()
            if txn_low.startswith("paid to") or txn_low.startswith("money sent to"):
                desc = "Paytm Payment"
                exp_type = "Miscellaneous"
                merch_cat = tags
            else:
                desc = clean_text(txn)
                exp_type = "Miscellaneous"
                merch_cat = tags

        out_rows.append(
            {
                "Period": period,
                "Description": desc,
                "Amount": amount,
                "Expense Type": exp_type,
                "Merchant Category": merch_cat,
                "Account": account_value,
            }
        )

    odf = pd.DataFrame(out_rows)
    # Ensure column order with Amount at position 3
    odf = odf[["Period", "Description", "Amount", "Expense Type", "Merchant Category", "Account"]]

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        odf.to_excel(writer, sheet_name="Paytm Transactions", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Paytm Transactions"]

        header_fmt = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#F4B183",  # Orange Accent 6 style
                "border": 1,
            }
        )
        cell_fmt = workbook.add_format({"border": 1})
        num_fmt = workbook.add_format({"num_format": "#,##0.00"})

        # Re-write header with format
        for col_idx, col_name in enumerate(odf.columns):
            worksheet.write(0, col_idx, col_name, header_fmt)

        # Apply borders and number format
        nrows, ncols = odf.shape
        worksheet.conditional_format(
            0,
            0,
            nrows,
            ncols - 1,
            {"type": "no_blanks", "format": cell_fmt},
        )
        amt_col = odf.columns.get_loc("Amount")
        worksheet.set_column(amt_col, amt_col, 14, num_fmt)

        # Auto-adjust widths
        for col_idx, col_name in enumerate(odf.columns):
            max_len = len(str(col_name))
            series = odf[col_name]
            col_max = (
                series.map(lambda x: len(str(x)) if pd.notna(x) else 0).max()
                if not series.empty
                else 0
            )
            width = min(max(max_len, col_max) + 2, 60)
            if col_idx != amt_col:
                worksheet.set_column(col_idx, col_idx, width)

        # Freeze top row and add filter
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, nrows, ncols - 1)

        # In-sheet pivot summary at I2: Amount by Expense Type and Merchant Category
        pivot_df = odf[odf["Description"].str.lower() != "gold coin redemption"].copy()
        pivot_df = (
            pivot_df.groupby(["Expense Type", "Merchant Category"], as_index=False)["Amount"]
            .sum()
            .sort_values(by=["Expense Type", "Merchant Category"])
        )

        pivot_start_row = 1  # I2
        pivot_start_col = 8  # Col I
        pivot_headers = ["Expense Type", "Merchant Category", "Amount"]
        for idx, h in enumerate(pivot_headers):
            worksheet.write(pivot_start_row, pivot_start_col + idx, h, header_fmt)

        for r_idx, row in enumerate(pivot_df.itertuples(index=False), start=pivot_start_row + 1):
            worksheet.write(r_idx, pivot_start_col + 0, row[0], cell_fmt)
            worksheet.write(r_idx, pivot_start_col + 1, row[1], cell_fmt)
            worksheet.write_number(r_idx, pivot_start_col + 2, float(row[2]), workbook.add_format({"border": 1, "num_format": "#,##0.00"}))

        # Autofilter for pivot block
        pivot_end_row = pivot_start_row + len(pivot_df)
        worksheet.autofilter(pivot_start_row, pivot_start_col, pivot_end_row, pivot_start_col + 2)

        # Widths for pivot columns
        worksheet.set_column(pivot_start_col + 0, pivot_start_col + 0, 20)
        worksheet.set_column(pivot_start_col + 1, pivot_start_col + 1, 22)
        worksheet.set_column(pivot_start_col + 2, pivot_start_col + 2, 14)

    return odf


def main():
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        candidates = sorted(Path(DEFAULT_INPUT_DIR).glob("*.xlsx"))
        if not candidates:
            raise FileNotFoundError(f"No .xlsx files found in {DEFAULT_INPUT_DIR}")
        input_file = str(candidates[0])

    mapping_file = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_MAPPING_FILE
    result = parse_paytm(input_file, mapping_file)
    print(f"Input: {input_file}")
    print(f"Mapping: {mapping_file}")
    print(f"Output: {OUTPUT_FILE}")
    print(f"Rows: {len(result)}")


if __name__ == "__main__":
    main()
