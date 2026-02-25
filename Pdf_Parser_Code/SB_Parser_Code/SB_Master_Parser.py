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
MAPPING_FILE = os.path.join(PROJECT_DIR, "Reference Documents", "Merchant category mapping.xlsx")


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


def clean_text(value):
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def load_sb_mapping_rules():
    xl = pd.ExcelFile(MAPPING_FILE)
    sheet_name = None
    for candidate in ("SB Mapping", "SB Merchant category mapping"):
        if candidate in xl.sheet_names:
            sheet_name = candidate
            break
    if not sheet_name:
        raise ValueError("SB mapping sheet not found. Expected 'SB Mapping' or 'SB Merchant category mapping'.")
    df = pd.read_excel(MAPPING_FILE, sheet_name=sheet_name)
    rules_by_bank = {}
    fallback_by_bank = {}
    default_fallback = None
    for _, row in df.iterrows():
        bank = clean_text(row.get("Bank", "")).lower()
        if not bank:
            bank = "default"
        keyword = clean_text(row.get("Keyword Pattern", ""))
        mapped = (
            clean_text(row.get("Mode", "")) or "Uncategorized",
            clean_text(row.get("Expense Type", "")) or "Uncategorized",
            clean_text(row.get("Merchant Category", "")) or "Uncategorized",
            clean_text(row.get("Store Name", "")) or "Unknown",
        )
        if keyword:
            rules_by_bank.setdefault(bank, []).append((keyword.lower(),) + mapped)
        else:
            fallback_by_bank[bank] = mapped
        if bank == "default" and default_fallback is None:
            default_fallback = mapped
    if default_fallback is None:
        default_fallback = ("Uncategorized", "Uncategorized", "Uncategorized", "Unknown")
    return rules_by_bank, fallback_by_bank, default_fallback


def normalize_match_text(value):
    s = clean_text(value).lower()
    return re.sub(r"[^a-z0-9]", "", s)


def ordered_token_match(keyword, description):
    # Match keyword tokens in order, allowing extra tokens in between.
    kw_tokens = [t for t in re.split(r"[^a-z0-9]+", clean_text(keyword).lower()) if t]
    desc_tokens = [t for t in re.split(r"[^a-z0-9]+", clean_text(description).lower()) if t]
    if not kw_tokens:
        return False
    i = 0
    for tok in desc_tokens:
        if tok == kw_tokens[i]:
            i += 1
            if i == len(kw_tokens):
                return True
    return False


def unordered_token_match(keyword, description):
    # Match keyword tokens in any order; allow prefix matches (e.g. SBIN -> SBIN0011739).
    kw_tokens = [t for t in re.split(r"[^a-z0-9]+", clean_text(keyword).lower()) if t]
    desc_tokens = [t for t in re.split(r"[^a-z0-9]+", clean_text(description).lower()) if t]
    if not kw_tokens:
        return False
    for kt in kw_tokens:
        if not any(dt == kt or dt.startswith(kt) for dt in desc_tokens):
            return False
    return True


def resolve_directional_mapping(matches, amount):
    # User rule for duplicate keyword rows:
    # withdrawal => prefer "SBI to IDFC ..."
    # deposit    => prefer "IDFC ... to SBI ..."
    if amount is None or not matches:
        return matches[0]
    is_withdrawal = amount < 0
    for item in matches:
        _, mode, exp_type, merch_cat, store_name = item
        mapped_text = f"{mode} {exp_type} {merch_cat} {store_name}".lower()
        if is_withdrawal and re.search(r"\bsbi\b.*\bto\b.*\bidfc\b", mapped_text):
            return item
        if (not is_withdrawal) and re.search(r"\bidfc\b.*\bto\b.*\bsbi\b", mapped_text):
            return item
    return matches[0]


def classify_sb_description(description, amount, rules, fallback):
    desc = clean_text(description).lower()
    desc_norm = normalize_match_text(desc)
    matches = []
    for keyword, mode, exp_type, merch_cat, store_name in rules:
        key_norm = normalize_match_text(keyword)
        if not key_norm:
            continue
        if (
            keyword in desc
            or key_norm in desc_norm
            or ordered_token_match(keyword, desc)
            or unordered_token_match(keyword, desc)
        ):
            matches.append((keyword, mode, exp_type, merch_cat, store_name))
    if matches:
        chosen = resolve_directional_mapping(matches, amount)
        _, mode, exp_type, merch_cat, store_name = chosen
        return mode, exp_type, merch_cat, store_name
    return fallback


def account_to_bank_key(account):
    a = clean_text(account).lower()
    if "axis" in a:
        return "axis"
    if "hdfc" in a:
        return "hdfc"
    if "icici" in a:
        return "icici"
    if "idfc" in a:
        return "idfc"
    if "yes" in a:
        return "yes"
    return a


def classify_sb_row(description, amount, account, rules_by_bank, fallback_by_bank, default_fallback):
    bank = account_to_bank_key(account)
    bank_rules = rules_by_bank.get(bank, [])
    fallback = fallback_by_bank.get("default", default_fallback)
    if bank_rules:
        return classify_sb_description(description, amount, bank_rules, fallback)
    # If no bank-specific match/rules, use Default bank row values.
    return fallback


def extract_yes_period(text):
    text = text or ""
    # Preferred: explicit "as on dd/mm/yyyy"
    m = re.search(r"Account Relationship Summary as on\s*(\d{2}/\d{2}/\d{4})", text, re.I)
    if m:
        d = parse_date(m.group(1))
        if d:
            return format_period_from_date(d)
    # Fallback: "YOUR CONSOLIDATED STATEMENT FOR JAN' 26"
    m = re.search(r"YOUR\s+CONSOLIDATED\s+STATEMENT\s+FOR\s+([A-Za-z]{3})'?\s*(\d{2})", text, re.I)
    if m:
        mon = m.group(1).title()
        yy = int(m.group(2))
        try:
            d = datetime.strptime(f"01-{mon}-{2000 + yy}", "%d-%b-%Y").date()
            return format_period_from_date(d)
        except ValueError:
            pass
    return "Unknown"


def is_yes_footer_line(line):
    t = (line or "").strip().lower()
    if not t:
        return True
    footer_starts = (
        "opening balance :",
        "total withdrawals :",
        "total deposits:",
        "closing balance :",
        "mandatory disclaimer:",
        "for any assistance required",
        "please refer to important messages",
        "transaction codes in your account statement",
        "have you registered a nominee",
        "*please ignore",
        "please check the entries in the statement",
        'say "hi" on',
        "whatsapp banking",
        "email us at",
        "toll free number",
        "cin:",
        "canada:",
        "uk:",
        "uae:",
        "this is an automatically generated statement",
        "page ",
    )
    if any(t.startswith(x) for x in footer_starts):
        return True
    if "yes rewardz" in t or "account relationship summary" in t:
        return True
    return False


def clean_yes_description(text, strip_leading_ref=True):
    s = clean_text(text)
    # Remove reference-number artifacts that bleed into Description from adjacent column.
    if strip_leading_ref:
        s = re.sub(r"^\d{6,}(?:/\d{6,})+\s*", "", s)
        s = re.sub(r"^\d{9,}\s*", "", s)
    s = re.sub(r"^PCA:[A-Z0-9:/-]+\s*", "", s)
    s = re.sub(r"\bPCA:[A-Z0-9:/-]+\b", "", s)
    s = re.sub(r"\bND\d{2}-[A-Z]{3}-\d{2}\s+\d{2}:\d{2}:\d{2}\b", "", s)
    # remove any trailing amount triplet accidentally attached to description
    s = re.sub(
        r"\s+\d{1,3}(?:,\d{2,3})*\.\d{2}\s+\d{1,3}(?:,\d{2,3})*\.\d{2}\s+\d{1,3}(?:,\d{2,3})*\.\d{2}$",
        "",
        s,
    ).strip()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def append_yes_fragment(desc, frag):
    desc = clean_text(desc)
    frag = clean_text(frag)
    if not frag:
        return desc
    # Generic word-wrap merge rules:
    # 1) "... S" + "ER" => "... SER"
    # 2) "...-P" + "riyanka" => "...-Priyanka"
    # 3) "...-YE" + "SMIDAS..." => "...-YESMIDAS..."
    if re.match(r"^[A-Za-z0-9]", frag):
        if re.search(r"\b[A-Za-z]\s*$", desc):
            return re.sub(r"\s+$", "", desc) + frag
        if re.search(r"[-/][A-Za-z]{1,3}\s*$", desc):
            return re.sub(r"\s+$", "", desc) + frag
    return f"{desc} {frag}".strip()


def append_idfc_fragment(desc, frag):
    desc = clean_text(desc)
    frag = clean_text(frag)
    if not frag:
        return desc
    # IDFC table cells may split a single word at newline (e.g. "Transfe" + "rtofamily...").
    if (
        desc
        and " " not in frag
        and re.search(r"[A-Za-z]$", desc)
        and re.match(r"^[a-z]{3,}[A-Za-z0-9/-]*$", frag)
    ):
        return f"{desc}{frag}"
    return f"{desc} {frag}".strip()


def detect_bank_from_text(text):
    t = (text or "").upper()
    if "AXIS BANK" in t:
        return "Axis"
    if "HDFC BANK" in t:
        return "HDFC"
    if "ICICI BANK" in t:
        return "ICICI"
    if "IDFC FIRST BANK" in t or "IDFC BANK" in t:
        return "IDFC"
    if "YES BANK" in t:
        return "Yes"
    return None


def detect_pdf_context(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages[:2]:
                text += "\n" + (page.extract_text() or "")
    except Exception:
        return {"bank": None, "customer_name": None}
    bank = detect_bank_from_text(text)
    t = text.upper()
    customer_name = None
    if "PRIYANKA JAIN" in t:
        customer_name = "PRIYANKA JAIN"
    elif "ABHISHEK JAIN" in t:
        customer_name = "ABHISHEK JAIN"
    return {"bank": bank, "customer_name": customer_name}


def resolve_account_name(detected_bank, customer_name, fallback_account):
    bank_key = (detected_bank or "").strip().lower()
    if bank_key == "yes":
        if customer_name == "PRIYANKA JAIN":
            return "PJ Yes"
        return "Yes"
    if bank_key == "idfc":
        if customer_name == "ABHISHEK JAIN":
            return "AJ IDFC"
        if customer_name == "PRIYANKA JAIN":
            return "PJ IDFC"
        return "IDFC"
    return detected_bank or fallback_account


def get_parser_by_bank(bank):
    bank_key = (bank or "").lower()
    if bank_key == "axis":
        return parse_axis
    if bank_key == "hdfc":
        return parse_hdfc
    if bank_key == "icici":
        return parse_icici
    if bank_key == "idfc":
        return parse_idfc
    if bank_key == "yes":
        return parse_yes
    return None


def format_sheet(workbook, worksheet, df):
    nrows, ncols = df.shape
    if ncols == 0:
        return

    header_fmt = workbook.add_format({"bold": True, "bg_color": "#F4B183", "border": 1})
    cell_fmt = workbook.add_format({"border": 1})
    amt_fmt = workbook.add_format({"border": 1, "num_format": "#,##0.00"})

    # Header styling
    for c, col in enumerate(df.columns):
        worksheet.write(0, c, col, header_fmt)

    # Borders + numeric formats for all data cells
    numeric_cols = {"Amount", "Balance"}
    for r in range(1, nrows + 1):
        for c, col in enumerate(df.columns):
            val = df.iloc[r - 1, c]
            if pd.isna(val):
                worksheet.write_blank(r, c, None, cell_fmt)
            elif col in numeric_cols:
                try:
                    worksheet.write_number(r, c, float(val), amt_fmt)
                except Exception:
                    worksheet.write(r, c, val, cell_fmt)
            else:
                worksheet.write(r, c, val, cell_fmt)

    # Auto width
    for c, col in enumerate(df.columns):
        max_len = len(str(col))
        if nrows:
            col_max = df[col].map(lambda x: len(str(x)) if pd.notna(x) else 0).max()
            max_len = max(max_len, int(col_max))
        worksheet.set_column(c, c, min(max_len + 2, 60))

    # Approx auto row height
    worksheet.set_row(0, 20)
    for r in range(1, nrows + 1):
        row_vals = ["" if pd.isna(v) else str(v) for v in df.iloc[r - 1].tolist()]
        longest = max((len(v) for v in row_vals), default=0)
        lines = max(1, (longest // 45) + 1)
        worksheet.set_row(r, min(15 * lines, 60))
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, nrows, ncols - 1)


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
        prev_balance = None

        # Preferred path: table-based extraction to preserve full description cell text.
        for page in pdf.pages:
            for table in extract_tables(page):
                if not table:
                    continue
                header_idx = None
                for i, row in enumerate(table):
                    row_text = " ".join([clean_text(c) for c in (row or []) if clean_text(c)]).lower()
                    if (
                        "date and time" in row_text
                        and "transaction details" in row_text
                        and "withdrawals" in row_text
                        and "deposits" in row_text
                        and "balance" in row_text
                    ):
                        header_idx = i
                        break
                if header_idx is None:
                    continue

                for row in table[header_idx + 1 :]:
                    if not row or len(row) < 7:
                        continue

                    date = parse_date(row[1]) or parse_date(row[0])
                    if not date:
                        continue

                    raw_desc_cell = str(row[2] if len(row) > 2 and row[2] is not None else "")
                    desc = ""
                    for part in raw_desc_cell.split("\n"):
                        desc = append_idfc_fragment(desc, part)
                    desc = clean_text(desc)

                    debit = parse_amount(row[4] if len(row) > 4 else None)
                    credit = parse_amount(row[5] if len(row) > 5 else None)
                    balance = parse_amount(row[6] if len(row) > 6 else None)

                    if "B/F" in desc.upper():
                        if balance is not None:
                            prev_balance = balance
                        continue

                    amount = None
                    if debit is not None and debit != 0:
                        amount = -debit
                    elif credit is not None and credit != 0:
                        amount = credit
                    elif prev_balance is not None and balance is not None:
                        amount = balance - prev_balance
                    if amount is None:
                        continue

                    prev_balance = balance if balance is not None else prev_balance
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

        if records:
            existing_keys = set()
            for r in records:
                bal_key = round(float(r["Balance"]), 2) if r.get("Balance") is not None else None
                existing_keys.add((r.get("Date"), round(float(r.get("Amount", 0.0)), 2), bal_key))
        else:
            existing_keys = set()

        # Text fallback pass for rows missed by table extraction (common on page breaks).
        prev_balance = None
        carry_prefix = []
        current_txn = None

        def is_noise_line(value):
            u = (value or "").upper()
            return any(
                h in u
                for h in (
                    "CONSOLIDATED STATEMENT",
                    "STATEMENT PERIOD",
                    "SUMMARY OF YOUR RELATIONSHIP",
                    "DATE AND TIME",
                    "VALUE DATE",
                    "TRANSACTION DETAILS",
                    "NO. (INR)",
                    "WITHDRAWALS",
                    "DEPOSITS",
                    "BALANCE",
                    "CUSTOMER ID",
                    "ACCOUNT NO.",
                    "CUSTOMER NAME",
                    "REGISTERED OFFICE:",
                    "PAGE ",
                    "ACCOUNT BRANCH",
                    "BRANCH ADDRESS",
                    "ACCOUNT OPENING DATE",
                    "ACCOUNT STATUS",
                    "CURRENCY INR",
                )
            )

        def is_txn_prefix_line(value):
            return re.match(
                r"^(IMPS|BILLPAY|ATM|UPI|NEFT|RTGS|POS|ECOM|ACH|NACH|MONTHLY|WITHDRAWAL|CASH)\b",
                (value or "").upper(),
            ) is not None

        def flush_current_txn():
            nonlocal prev_balance, current_txn, records
            if not current_txn:
                return
            line = current_txn["line"]
            nums = extract_amounts_with_decimals(line)
            dates = re.findall(r"\b\d{2}\s+[A-Za-z]{3}\s+\d{2}\b", line)
            if len(nums) < 2 or not dates:
                current_txn = None
                return
            date = parse_date(dates[-1]) or parse_date(dates[0])
            if not date:
                current_txn = None
                return

            balance = parse_amount(nums[-1])
            amt = parse_amount(nums[-2])
            if amt is None or balance is None:
                current_txn = None
                return

            desc_tail = re.sub(
                r"^\d{2}\s+[A-Za-z]{3}\s+\d{2}(?:\s+\d{2}:\d{2})?(?:\s+\d{2}\s+[A-Za-z]{3}\s+\d{2})?\s+",
                "",
                line,
            )
            desc_tail = re.sub(
                r"\s+" + re.escape(nums[-2]) + r"\s+" + re.escape(nums[-1]) + r"\s*(?:CR|DR)?\s*$",
                "",
                desc_tail,
                flags=re.I,
            ).strip()
            if re.fullmatch(r"(?:\d{1,3}(?:,\d{2,3})*\.\d{2}\s*)+(?:CR|DR)?", desc_tail, flags=re.I):
                desc_tail = ""

            desc = ""
            for part in current_txn["pre"]:
                desc = append_idfc_fragment(desc, part)
            desc = append_idfc_fragment(desc, desc_tail)
            for part in current_txn["post"]:
                desc = append_idfc_fragment(desc, part)
            desc = clean_text(desc)

            if prev_balance is None:
                prev_balance = balance
                current_txn = None
                return

            amount = amt if balance > prev_balance else -amt
            prev_balance = balance
            key = (date, round(float(amount), 2), round(float(balance), 2))
            if key not in existing_keys:
                existing_keys.add(key)
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
            current_txn = None

        for page in pdf.pages:
            text = page.extract_text() or ""
            raw_lines = [clean_text(l) for l in text.split("\n") if clean_text(l)]
            for l in raw_lines:
                nums = extract_amounts_with_decimals(l)
                if re.search(r"\bopening balance\b", l, re.I) and nums:
                    bal = parse_amount(nums[-1])
                    if bal is not None:
                        prev_balance = bal
                    continue
                if is_noise_line(l):
                    continue
                if re.match(r"^\d{2}\s+[A-Za-z]{3}\s+\d{2}\b", l):
                    flush_current_txn()
                    current_txn = {"pre": carry_prefix, "line": l, "post": []}
                    carry_prefix = []
                    continue
                if current_txn is not None:
                    if is_txn_prefix_line(l):
                        carry_prefix.append(l)
                    else:
                        current_txn["post"].append(l)
                elif is_txn_prefix_line(l):
                    carry_prefix.append(l)

        flush_current_txn()
    return records


def parse_yes(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_yes_period(first_text)

        # Preferred path: table-based extraction (more reliable for YES multi-line descriptions).
        for page in pdf.pages:
            for table in extract_tables(page):
                if not table:
                    continue
                header_idx = None
                for i, row in enumerate(table):
                    row_text = " ".join([clean_text(c) for c in (row or []) if clean_text(c)])
                    if (
                        "Transaction" in row_text
                        and "Description" in row_text
                        and "Withdrawals" in row_text
                        and "Running Balance" in row_text
                    ):
                        header_idx = i
                        break
                if header_idx is None:
                    continue

                prev_balance = None
                for row in table[header_idx + 1 :]:
                    if not row or len(row) < 7:
                        continue
                    date = parse_date(row[0])
                    if not date:
                        continue

                    # Use Description column as-is from YES table and only merge wrapped lines.
                    raw_desc_cell = str(row[2] if len(row) > 2 and row[2] is not None else "")
                    desc = ""
                    for part in raw_desc_cell.split("\n"):
                        desc = append_yes_fragment(desc, part)
                    desc = clean_text(desc)

                    debit = parse_amount(row[4] if len(row) > 4 else None)
                    credit = parse_amount(row[5] if len(row) > 5 else None)
                    balance = parse_amount(row[6] if len(row) > 6 else None)

                    if "B/F" in desc.upper():
                        if balance is not None:
                            prev_balance = balance
                        continue

                    amount = None
                    if debit is not None and debit != 0:
                        amount = -debit
                    elif credit is not None and credit != 0:
                        amount = credit
                    elif prev_balance is not None and balance is not None:
                        amount = balance - prev_balance
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

        if records:
            return records

        # Fallback path: text-line extraction for PDFs where table extraction fails.
        for page in pdf.pages:
            text = page.extract_text() or ""
            raw_lines = [l.strip() for l in text.split("\n") if l.strip()]
            prev_balance = None
            pending_prefix = []
            tx_start_idxs = [i for i, ln in enumerate(raw_lines) if re.match(r"^\d{2}/\d{2}/\d{4}\b", ln)]
            for pos, start_idx in enumerate(tx_start_idxs):
                end_idx = tx_start_idxs[pos + 1] if pos + 1 < len(tx_start_idxs) else len(raw_lines)
                line = raw_lines[start_idx]
                cont_lines = raw_lines[start_idx + 1 : end_idx]
                if "B/F" in line:
                    nums = extract_amounts_with_decimals(line)
                    if nums:
                        prev_balance = parse_amount(nums[-1])
                    # Some PDFs place next txn leading description line(s) between B/F and first txn line.
                    for cl in cont_lines:
                        cln = clean_text(cl)
                        if is_yes_footer_line(cln):
                            break
                        if re.search(r"[A-Za-z]", cln):
                            pending_prefix.append(cln)
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
                desc = re.sub(r"^\d{2}/\d{2}/\d{4}\s+\d{2}/\d{2}/\d{4}\s+", "", line)
                desc = re.sub(
                    r"\s+" + re.escape(nums[-3]) + r"\s+" + re.escape(nums[-2]) + r"\s+" + re.escape(nums[-1]) + r"$",
                    "",
                    desc,
                ).strip()
                if pending_prefix:
                    for pp in pending_prefix:
                        desc = append_yes_fragment(desc, pp)
                    pending_prefix = []

                # Keep only meaningful continuation fragments (e.g. "ER"), skip footer/noise lines.
                carry_to_next = []
                for cl in cont_lines:
                    cln = clean_text(cl)
                    if is_yes_footer_line(cln):
                        break
                    if re.fullmatch(r"[0-9/:-]+", cln):
                        # keep long numeric refs (e.g. 120407835544) out of description
                        continue
                    if not re.search(r"[A-Za-z]", cln):
                        continue
                    # This line usually belongs to the next transaction line that has empty description.
                    if cln.lower().startswith("credit interest"):
                        carry_to_next.append(cln)
                        continue
                    desc = append_yes_fragment(desc, cln)

                if carry_to_next:
                    pending_prefix.extend(carry_to_next)

                desc = clean_yes_description(desc)
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
        ctx = detect_pdf_context(pdf_path)
        parser = get_parser(pdf_path) or get_parser_by_bank(ctx["bank"])
        if not parser:
            print("   ⚠️ No parser found for this file")
            continue
        try:
            records = parser(pdf_path)
            if records:
                fallback_account = records[0].get("Account", "")
                account_name = resolve_account_name(ctx["bank"], ctx.get("customer_name"), fallback_account)
                for rec in records:
                    rec["Account"] = account_name
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
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%d-%b-%Y")

    # SB AC expenses columns A-F (no Card Variant)
    df = df[["Period", "Date", "Account", "Description", "Amount", "Balance"]]

    # Summary sheet:
    # A: Description from SB AC expenses
    # B-E derived from SB mapping using top-to-bottom, case-insensitive match on column A.
    rules_by_bank, fallback_by_bank, default_fallback = load_sb_mapping_rules()
    summary_source_df = df[
        ~(
            df["Account"].astype(str).str.strip().str.lower().eq("axis")
            & df["Description"].astype(str).str.strip().str.lower().eq("opening balance")
        )
    ].copy()

    mapped = summary_source_df.apply(
        lambda r: pd.Series(
            classify_sb_row(
                r["Description"],
                r["Amount"],
                r["Account"],
                rules_by_bank,
                fallback_by_bank,
                default_fallback,
            )
        ),
        axis=1,
    )
    mapped.columns = ["Mode", "Expense Type", "Merchant Category", "Store Name"]
    summary_df = pd.concat(
        [summary_source_df[["Period", "Account", "Description"]], mapped, summary_source_df[["Amount"]]],
        axis=1,
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="SB AC expenses", index=False)
        summary_df.to_excel(writer, sheet_name="SB Categorized Summary", index=False)
        workbook = writer.book
        format_sheet(workbook, writer.sheets["SB AC expenses"], df)
        format_sheet(workbook, writer.sheets["SB Categorized Summary"], summary_df)

    print("\\n======================================================================")
    print("✅ SB AGGREGATION COMPLETE")
    print("======================================================================")
    print(f"Total PDFs:            {len(pdf_paths)}")
    print(f"Total Transactions:    {len(df)}")
    print(f"Output File:           {OUTPUT_FILE}")
    print("======================================================================")


if __name__ == "__main__":
    main()
