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
    # Excel mapping:
    # - PayTm_1: Tags, Description, Expense Type, Merchant Category
    # - PayTm_1: account map from columns Your Account -> Value (now F/G as requested)
    rules = []
    account_map = []

    if mapping_path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        xl = pd.ExcelFile(mapping_path)

        class_sheet = "PayTm_1" if "PayTm_1" in xl.sheet_names else "PayTm"
        cdf = pd.read_excel(mapping_path, sheet_name=class_sheet)
        for _, row in cdf.iterrows():
            tag = str(row.get("Tags", "")).strip() if pd.notna(row.get("Tags", "")) else ""
            desc = str(row.get("Description", "")).strip() if pd.notna(row.get("Description", "")) else ""
            exp = str(row.get("Expense Type", "")).strip() if pd.notna(row.get("Expense Type", "")) else ""
            merch = str(row.get("Merchant Category", "")).strip() if pd.notna(row.get("Merchant Category", "")) else ""
            if tag or desc:
                rules.append((tag, desc, exp, merch))

        for _, row in cdf.iterrows():
            acc_in = str(row.get("Your Account", "")).strip() if pd.notna(row.get("Your Account", "")) else ""
            acc_out = str(row.get("Value", "")).strip() if pd.notna(row.get("Value", "")) else ""
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
        rules.append(("", key, exp, merch))
    return rules, account_map


def is_partial_match(source_value, rule_value):
    s = norm(source_value)
    r = norm(rule_value)
    if not s or not r:
        return False
    return r in s or s in r


def match_paytm1_rule(source_tags, source_desc, rules):
    # Matching rules:
    # 1) If rule description is empty -> compare tags only
    # 2) If rule tags is empty -> compare description only
    # 3) Else compare both tags and description
    best = None
    best_score = -1
    for rule_tag, rule_desc, exp_type, merch_cat in rules:
        tag_empty = not norm(rule_tag)
        desc_empty = not norm(rule_desc)

        tag_ok = is_partial_match(source_tags, rule_tag)
        desc_ok = is_partial_match(source_desc, rule_desc)

        if desc_empty and not tag_empty:
            matched = tag_ok
            score = len(norm(rule_tag))
        elif tag_empty and not desc_empty:
            matched = desc_ok
            score = len(norm(rule_desc))
        elif (not tag_empty) and (not desc_empty):
            matched = tag_ok and desc_ok
            score = len(norm(rule_tag)) + len(norm(rule_desc))
        else:
            matched = False
            score = -1

        if matched and score > best_score:
            best = (rule_tag, rule_desc, exp_type, merch_cat)
            best_score = score
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
    rules, account_map = load_paytm_mapping(mapping_path)
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

        m = match_paytm1_rule(tags, txn, rules)
        if m:
            _, _, exp_type, merch_cat = m
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

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        # Sheet 1: raw source as-is
        sdf.to_excel(writer, sheet_name="Paytm Transactions", index=False)
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

        def style_sheet(worksheet, dataframe):
            nrows, ncols = dataframe.shape
            for col_idx, col_name in enumerate(dataframe.columns):
                worksheet.write(0, col_idx, col_name, header_fmt)
            worksheet.conditional_format(
                0,
                0,
                nrows,
                ncols - 1,
                {"type": "no_blanks", "format": cell_fmt},
            )
            if "Amount" in dataframe.columns:
                amt_col = dataframe.columns.get_loc("Amount")
                worksheet.set_column(amt_col, amt_col, 14, num_fmt)
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

        style_sheet(ws_raw, sdf)
        style_sheet(ws_summary, export_df)

    return export_df


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
