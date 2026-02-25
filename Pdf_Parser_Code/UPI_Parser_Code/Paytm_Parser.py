import os
import re
import sys
import shutil
from pathlib import Path
from datetime import datetime

import pandas as pd

PROJECT_DIR = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker"
)

DEFAULT_INPUT_DIR = os.path.join(PROJECT_DIR, "Bank_Statements", "UPI Statements")
DEFAULT_MAPPING_FILE = os.path.join(PROJECT_DIR, "Reference Documents", "Merchant category mapping.xlsx")
OUTPUT_DIR = os.path.join(PROJECT_DIR, "Output")
ARCHIVE_DIR = os.path.join(PROJECT_DIR, "Archive", "UPI", "PayTm")
LOG_FILE = os.path.join(PROJECT_DIR, "Logs", "File_Parser_log.txt")
ARCHIVE_ENABLED = False


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


def parse_int_like(value):
    num = parse_amount(value)
    if num is None:
        return None
    try:
        return int(round(num))
    except Exception:
        return None


def derive_period_label_from_dates(df):
    date_series = pd.to_datetime(df.get("Date"), dayfirst=True, errors="coerce").dropna()
    if date_series.empty:
        return datetime.now().strftime("%b'%y")
    dt = date_series.max()
    return dt.strftime("%b'%y")


def build_output_filename(period_label):
    return f"Paytm_{period_label}.xlsx"


def append_parser_log(file_type, file_name, output_file_name, error_text=""):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    line = (
        f"Type: {file_type}, "
        f"File Name: {file_name}, "
        f"Date: {now}, "
        f"Output file name: {output_file_name}, "
        f"Errors: {error_text or 'None'}\n"
    )
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line)


def archive_processed_input(source_path, archive_file_name):
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    target_path = os.path.join(ARCHIVE_DIR, archive_file_name)
    if os.path.exists(target_path):
        stem = Path(archive_file_name).stem
        suffix = Path(archive_file_name).suffix
        target_path = os.path.join(
            ARCHIVE_DIR, f"{stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{suffix}"
        )
    shutil.move(source_path, target_path)
    return target_path


def load_source(source_path):
    df = pd.read_excel(source_path, sheet_name="Passbook Payment History")
    required = ["Date", "Transaction Details", "Tags"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Missing required column in input: {col}")
    return df


def load_paytm_mapping(mapping_path):
    # Excel mapping:
    # - PayTm_1 (A:F): Tags, Description, Other Transaction Details, Your Account,
    #   Expense Type, Merchant Category
    # - PayTm_1 (G:H): account map table (Your Account -> Value)
    rules = []
    account_map = []

    if mapping_path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        xl = pd.ExcelFile(mapping_path)

        class_sheet = "PayTm_1" if "PayTm_1" in xl.sheet_names else "PayTm"
        cdf = pd.read_excel(mapping_path, sheet_name=class_sheet)

        def pick_col(df, preferred):
            lowered = {str(c).strip().lower(): c for c in df.columns}
            for name in preferred:
                if name.lower() in lowered:
                    return lowered[name.lower()]
            return None

        tag_col = pick_col(cdf, ["Tags"])
        desc_col = pick_col(cdf, ["Description"])
        other_col = pick_col(cdf, ["Other Transaction Details"])
        class_acct_col = pick_col(cdf, ["Your Account"])
        exp_col = pick_col(cdf, ["Expense Type"])
        merch_col = pick_col(cdf, ["Merchant Category"])
        # Account map table in G/H: typically "Your Account.1" -> "Value"
        map_acct_col = pick_col(cdf, ["Your Account.1", "Your Account"])
        map_value_col = pick_col(cdf, ["Value"])

        for _, row in cdf.iterrows():
            tag = str(row.get(tag_col, "")).strip() if tag_col and pd.notna(row.get(tag_col, "")) else ""
            desc = str(row.get(desc_col, "")).strip() if desc_col and pd.notna(row.get(desc_col, "")) else ""
            other = str(row.get(other_col, "")).strip() if other_col and pd.notna(row.get(other_col, "")) else ""
            acct = str(row.get(class_acct_col, "")).strip() if class_acct_col and pd.notna(row.get(class_acct_col, "")) else ""
            exp = str(row.get(exp_col, "")).strip() if exp_col and pd.notna(row.get(exp_col, "")) else ""
            merch = str(row.get(merch_col, "")).strip() if merch_col and pd.notna(row.get(merch_col, "")) else ""
            if tag or desc or other or acct:
                rules.append((tag, desc, other, acct, exp, merch))

        for _, row in cdf.iterrows():
            acc_in = str(row.get(map_acct_col, "")).strip() if map_acct_col and pd.notna(row.get(map_acct_col, "")) else ""
            acc_out = str(row.get(map_value_col, "")).strip() if map_value_col and pd.notna(row.get(map_value_col, "")) else ""
            if acc_in and acc_out:
                account_map.append((acc_in, acc_out))
        return rules, account_map

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
        rules.append(("", key, "", "", exp, merch))
    return rules, account_map


def is_partial_match(source_value, rule_value):
    s = norm(source_value)
    r = norm(rule_value)
    if not s or not r:
        return False
    return r in s or s in r


def match_paytm1_rule(source_tags, source_desc, source_other, source_account, rules):
    # Strict top-to-bottom scan of PayTm_1 rows.
    # For each rule row, only populated keys are compared (A/B/C/D).
    # First matching row wins.
    def row_matches(rule):
        rule_tag, rule_desc, rule_other, rule_acct, *_ = rule
        checks = []
        if norm(rule_tag):
            checks.append(is_partial_match(source_tags, rule_tag))
        if norm(rule_desc):
            checks.append(is_partial_match(source_desc, rule_desc))
        if norm(rule_other):
            checks.append(is_partial_match(source_other, rule_other))
        if norm(rule_acct):
            checks.append(is_partial_match(source_account, rule_acct))
        return bool(checks) and all(checks)

    for rule in rules:
        if row_matches(rule):
            return rule
    return None


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
    rules, account_map = load_paytm_mapping(mapping_path)
    file_fallback_account = Path(source_path).stem
    period_label = derive_period_label_from_dates(sdf)
    output_file_name = build_output_filename(period_label)
    output_file_path = os.path.join(OUTPUT_DIR, output_file_name)

    out_rows = []
    for _, row in sdf.iterrows():
        date_val = pd.to_datetime(row["Date"], dayfirst=True, errors="coerce")
        if pd.isna(date_val):
            continue
        period = date_val.strftime("%b-%Y")

        txn = str(row.get("Transaction Details", "")).strip()
        amount = parse_amount(row.get("Amount", None))
        tags = clean_text(row.get("Tags", ""))

        source_other = str(row.get("Other Transaction Details (UPI ID or A/c No)", "")).strip()
        source_account = str(row.get("Your Account", "")).strip()
        account_value = derive_account_by_source(source_account, account_map, file_fallback_account)

        m = match_paytm1_rule(tags, txn, source_other, source_account, rules)
        if m:
            _, _, _, _, exp_type, merch_cat = m
        else:
            txn_low = txn.lower()
            if txn_low.startswith("paid to") or txn_low.startswith("money sent to"):
                exp_type = "Miscellaneous"
                merch_cat = tags
            else:
                exp_type = "Miscellaneous"
                merch_cat = tags

        out_rows.append(
            {
                "Period": period,
                "Account": account_value,
                "Expense Type": exp_type,
                "Merchant Category": merch_cat,
                "Amount": amount,
                "_source_description": clean_text(txn),
            }
        )

    summary_df = pd.DataFrame(out_rows)
    summary_df = summary_df[
        ["Period", "Account", "Expense Type", "Merchant Category", "Amount", "_source_description"]
    ]

    with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
        # Sheet 1: raw source; keep structure but force Amount numeric
        raw_df = sdf.copy()
        if "Amount" in raw_df.columns:
            raw_df["Amount"] = raw_df["Amount"].apply(parse_amount)
        for id_col in ["UPI Ref No.", "Order ID"]:
            if id_col in raw_df.columns:
                raw_df[id_col] = raw_df[id_col].apply(parse_int_like)
        raw_df.to_excel(writer, sheet_name="Paytm Transactions", index=False)
        # Sheet 2: categorized summary
        export_df = summary_df.drop(columns=["_source_description"])
        summary_sheet_name = "Categorized Txn Summary"
        export_df.to_excel(writer, sheet_name=summary_sheet_name, index=False)

        workbook = writer.book
        ws_raw = writer.sheets["Paytm Transactions"]
        ws_summary = writer.sheets[summary_sheet_name]

        header_fmt = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#F4B183",  # Orange Accent 6 style
                "border": 1,
            }
        )
        cell_fmt = workbook.add_format({"border": 1})
        num_fmt = workbook.add_format({"num_format": "#,##0.00"})
        num_border_fmt = workbook.add_format({"border": 1, "num_format": "#,##0.00"})
        int_num_border_fmt = workbook.add_format({"border": 1, "num_format": "0"})
        int_num_fmt = workbook.add_format({"num_format": "0"})

        def style_sheet(worksheet, dataframe):
            nrows, ncols = dataframe.shape
            for col_idx, col_name in enumerate(dataframe.columns):
                worksheet.write(0, col_idx, col_name, header_fmt)

            # Apply borders only to filled cells
            for r in range(1, nrows + 1):
                for c, col_name in enumerate(dataframe.columns):
                    val = dataframe.iloc[r - 1, c]
                    if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
                        continue
                    if col_name == "Amount":
                        try:
                            worksheet.write_number(r, c, float(val), num_border_fmt)
                        except Exception:
                            worksheet.write(r, c, val, cell_fmt)
                    elif col_name in ("UPI Ref No.", "Order ID"):
                        try:
                            worksheet.write_number(r, c, float(val), int_num_border_fmt)
                        except Exception:
                            worksheet.write(r, c, val, cell_fmt)
                    else:
                        worksheet.write(r, c, val, cell_fmt)

            if "Amount" in dataframe.columns:
                amt_col = dataframe.columns.get_loc("Amount")
                worksheet.set_column(amt_col, amt_col, 14, num_fmt)
            for id_col in ("UPI Ref No.", "Order ID"):
                if id_col in dataframe.columns:
                    id_idx = dataframe.columns.get_loc(id_col)
                    worksheet.set_column(id_idx, id_idx, 18, int_num_fmt)
            for col_idx, col_name in enumerate(dataframe.columns):
                max_len = len(str(col_name))
                series = dataframe[col_name]
                col_max = (
                    series.map(lambda x: len(str(x)) if pd.notna(x) else 0).max()
                    if not series.empty
                    else 0
                )
                width = min(max(max_len, col_max) + 2, 60)
                if col_name != "Amount":
                    worksheet.set_column(col_idx, col_idx, width)
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, nrows, ncols - 1)

        style_sheet(ws_raw, raw_df)
        style_sheet(ws_summary, export_df)

        # Pivot block on summary sheet at H2
        pivot_df = (
            export_df[export_df["Account"].astype(str).str.strip().str.lower() != "gold coins"]
            .groupby(["Account", "Expense Type", "Merchant Category"], as_index=False)["Amount"]
            .sum()
            .sort_values(by=["Account", "Expense Type", "Merchant Category"])
        )
        pivot_start_row = 1  # H2
        pivot_start_col = 7  # H
        pivot_headers = ["Account", "Expense Type", "Merchant Category", "Amount"]
        for idx, h in enumerate(pivot_headers):
            ws_summary.write(pivot_start_row, pivot_start_col + idx, h, header_fmt)

        pivot_num_fmt = workbook.add_format({"border": 1, "num_format": "#,##0.00"})
        for r_idx, row in enumerate(pivot_df.itertuples(index=False), start=pivot_start_row + 1):
            ws_summary.write(r_idx, pivot_start_col + 0, row[0], cell_fmt)
            ws_summary.write(r_idx, pivot_start_col + 1, row[1], cell_fmt)
            ws_summary.write(r_idx, pivot_start_col + 2, row[2], cell_fmt)
            ws_summary.write_number(r_idx, pivot_start_col + 3, float(row[3]), pivot_num_fmt)

        total_row = pivot_start_row + len(pivot_df) + 1
        total_amount = float(pivot_df["Amount"].sum()) if not pivot_df.empty else 0.0
        total_label_fmt = workbook.add_format({"bold": True, "border": 1})
        total_num_fmt = workbook.add_format({"bold": True, "border": 1, "num_format": "#,##0.00"})
        ws_summary.write(total_row, pivot_start_col + 2, "Total", total_label_fmt)
        ws_summary.write_number(total_row, pivot_start_col + 3, total_amount, total_num_fmt)

        pivot_end_row = total_row
        ws_summary.autofilter(pivot_start_row, pivot_start_col, pivot_end_row, pivot_start_col + 3)
        ws_summary.set_column(pivot_start_col + 0, pivot_start_col + 0, 16)
        ws_summary.set_column(pivot_start_col + 1, pivot_start_col + 1, 20)
        ws_summary.set_column(pivot_start_col + 2, pivot_start_col + 2, 22)
        ws_summary.set_column(pivot_start_col + 3, pivot_start_col + 3, 14, num_fmt)

        # Second pivot block on summary sheet at M2 (without Account)
        pivot2_df = (
            export_df[export_df["Account"].astype(str).str.strip().str.lower() != "gold coins"]
            .groupby(["Expense Type", "Merchant Category"], as_index=False)["Amount"]
            .sum()
            .sort_values(by=["Expense Type", "Merchant Category"])
        )
        pivot2_start_row = 1  # M2
        pivot2_start_col = 12  # M
        pivot2_headers = ["Expense Type", "Merchant Category", "Amount"]
        for idx, h in enumerate(pivot2_headers):
            ws_summary.write(pivot2_start_row, pivot2_start_col + idx, h, header_fmt)

        for r_idx, row in enumerate(pivot2_df.itertuples(index=False), start=pivot2_start_row + 1):
            ws_summary.write(r_idx, pivot2_start_col + 0, row[0], cell_fmt)
            ws_summary.write(r_idx, pivot2_start_col + 1, row[1], cell_fmt)
            ws_summary.write_number(r_idx, pivot2_start_col + 2, float(row[2]), pivot_num_fmt)

        pivot2_total_row = pivot2_start_row + len(pivot2_df) + 1
        pivot2_total_amount = float(pivot2_df["Amount"].sum()) if not pivot2_df.empty else 0.0
        ws_summary.write(pivot2_total_row, pivot2_start_col + 1, "Total", total_label_fmt)
        ws_summary.write_number(pivot2_total_row, pivot2_start_col + 2, pivot2_total_amount, total_num_fmt)

        ws_summary.autofilter(
            pivot2_start_row, pivot2_start_col, pivot2_total_row, pivot2_start_col + 2
        )
        ws_summary.set_column(pivot2_start_col + 0, pivot2_start_col + 0, 20)
        ws_summary.set_column(pivot2_start_col + 1, pivot2_start_col + 1, 22)
        ws_summary.set_column(pivot2_start_col + 2, pivot2_start_col + 2, 14, num_fmt)

    return export_df, output_file_name, output_file_path


def main():
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        candidates = sorted(Path(DEFAULT_INPUT_DIR).glob("*.xlsx"))
        if not candidates:
            raise FileNotFoundError(f"No .xlsx files found in {DEFAULT_INPUT_DIR}")
        input_file = str(candidates[0])

    mapping_file = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_MAPPING_FILE
    input_name = Path(input_file).name
    try:
        result, output_file_name, output_path = parse_paytm(input_file, mapping_file)
        archive_path = "Skipped (ARCHIVE_ENABLED=False)"
        if ARCHIVE_ENABLED:
            archive_path = archive_processed_input(input_file, output_file_name)
        append_parser_log("UPI", input_name, output_file_name, "")
        print(f"Input: {input_file}")
        print(f"Mapping: {mapping_file}")
        print(f"Output: {output_path}")
        print(f"Rows: {len(result)}")
        print(f"Archived Input: {archive_path}")
    except Exception as e:
        append_parser_log("UPI", input_name, "", str(e))
        raise


if __name__ == "__main__":
    main()
