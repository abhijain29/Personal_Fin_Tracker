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
MAPPING_FILE = os.path.join(PROJECT_DIR, "Reference Documents", "SB Mapping.xlsx")


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
        "%d.%m.%Y",
        "%d-%m-%y",
        "%d/%m/%y",
        "%d.%m.%y",
        "%d %b %y",
        "%d %b %Y",
        "%d-%b-%Y",
        "%d-%b-%y",
        "%d%b%Y",
        "%d%b%y",
        "%Y/%m/%d",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def append_wrapped_fragment(desc, frag):
    """
    Join wrapped cell fragments:
    - If a single word is split across lines, join without adding a space.
    - Otherwise join with a single space.
    """
    desc = clean_text(desc)
    frag = clean_text(frag)
    if not frag:
        return desc
    if not desc:
        return frag
    last = desc[-1]
    first = frag[0]
    if last == "-":
        return f"{desc[:-1]}{frag}"
    if last.isalpha() and first.isalpha():
        # Heuristic: wrapped words usually continue in lowercase (or mixed case).
        if first.islower() or last.islower():
            return f"{desc}{frag}"
    return f"{desc} {frag}".strip()


def extract_amounts_with_decimals(text):
    if not text:
        return []
    return re.findall(r"\d{1,3}(?:,\d{2,3})*\.\d{2}", text)


def format_period_from_date(d):
    if not d:
        return "Unknown"
    return d.strftime("%b-%Y")


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
        m = re.search(r"period\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})\s*-\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})", text, re.I)
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
    if bank_name == "SBI":
        m = re.search(r"As on\s*(\d{2}-\d{2}-\d{2})", text, re.I)
        if m:
            return format_period_from_date(parse_date(m.group(1)))
    if bank_name == "HSBC":
        m = re.search(r"Statement Date\s*(\d{2}[A-Za-z]{3}\d{4})", text, re.I)
        if m:
            return format_period_from_date(parse_date(m.group(1)))
    if bank_name == "IndusInd":
        m = re.search(r"Period\s*:\s*(\d{2}-[A-Za-z]{3}-\d{4})\s*To\s*(\d{2}-[A-Za-z]{3}-\d{4})", text, re.I)
        if m:
            return format_period_from_date(parse_date(m.group(2)))
    return "Unknown"


def extract_tables(page):
    return page.extract_tables() or []


def clean_text(value):
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


_KNOWN_WRAP_WORDS_CACHE = None


def load_known_wrap_words():
    """
    Build a set of known words from the "Bank Name map" worksheet to repair
    PDF extraction artifacts like "Abhi shek" -> "Abhishek".
    """
    global _KNOWN_WRAP_WORDS_CACHE
    if _KNOWN_WRAP_WORDS_CACHE is not None:
        return _KNOWN_WRAP_WORDS_CACHE
    words = set()
    try:
        df = pd.read_excel(MAPPING_FILE, sheet_name=BANK_NAME_MAP_SHEET)
        cols = list(df.columns)
        if len(cols) >= 3:
            df = df.rename(columns={cols[0]: "Bank PDF", cols[1]: "Text", cols[2]: "Output"})
        for val in df.get("Text", []):
            s = clean_text(val)
            for w in re.findall(r"[A-Za-z]{4,}", s):
                words.add(w.upper())
    except Exception:
        words = set()
    _KNOWN_WRAP_WORDS_CACHE = words
    return words


def fix_spaced_known_words(text, known_words):
    """
    Remove spaces between letter runs only when the concatenated token is a
    known word (prevents accidental merges like 'BANK LTD').
    """
    s = clean_text(text)
    if not s or not known_words:
        return s

    def _join(m):
        a = m.group(1)
        b = m.group(2)
        if (a + b).upper() in known_words:
            return a + b
        return m.group(0)

    # Multiple passes to handle cases like "AB HI SHEK" (rare).
    for _ in range(2):
        s2 = re.sub(r"\b([A-Za-z]{2,})\s+([A-Za-z]{2,})\b", _join, s)
        if s2 == s:
            break
        s = s2
    return s


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
        mc_derived_pos = clean_text(row.get("MC Derived - Positive", ""))
        merchant_category_raw = clean_text(row.get("Merchant Category", ""))
        mapped = (
            clean_text(row.get("Mode", "")) or "Uncategorized",
            clean_text(row.get("Expense Type", "")) or "Uncategorized",
            # Keep blank when empty so we can derive from "MC Derived - Positive".
            merchant_category_raw,
            clean_text(row.get("Store Name", "")) or "Unknown",
            mc_derived_pos,
        )
        if keyword:
            rules_by_bank.setdefault(bank, []).append((keyword.lower(),) + mapped)
        else:
            fallback_by_bank[bank] = mapped
        if bank == "default" and default_fallback is None:
            default_fallback = mapped
    if default_fallback is None:
        default_fallback = ("Uncategorized", "Uncategorized", "Uncategorized", "Unknown", "")
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


def derive_expense_type(expense_type, amount):
    if clean_text(expense_type).lower() != "derived":
        return expense_type
    if amount is None:
        return "Uncategorized"
    return "Money In" if float(amount) >= 0 else "Money Out"


def reverse_arrow_text(value):
    """
    Reverse a mapping string around the first -->, ignoring whitespace differences.
    Example: "Yes PJ --> IDFC PJ (Self)" -> "IDFC PJ --> Yes PJ (Self)"
    """
    s = clean_text(value)
    if "-->" not in s:
        return s
    # Keep trailing "(Self)" on the right-most side even after reversing.
    self_suffix = ""
    m_self = re.search(r"\s*\((self)\)\s*$", s, flags=re.IGNORECASE)
    if m_self:
        self_suffix = " (Self)"
        s = re.sub(r"\s*\((self)\)\s*$", "", s, flags=re.IGNORECASE).strip()

    left, right = re.split(r"\s*-->\s*", s, maxsplit=1)
    left = clean_text(left)
    right = clean_text(right)
    if not left or not right:
        return s + self_suffix
    return f"{right} --> {left}{self_suffix}"


def derive_merchant_category(merchant_category, mc_derived_positive, amount):
    # Only derive when mapping MC is empty in SB Mapping sheet.
    if clean_text(merchant_category):
        return merchant_category
    derived = clean_text(mc_derived_positive)
    if not derived:
        return "Uncategorized"
    if amount is None:
        return derived
    if float(amount) >= 0:
        return derived
    return reverse_arrow_text(derived)


def classify_sb_description(description, amount, rules, fallback):
    desc = clean_text(description).lower()
    desc_norm = normalize_match_text(desc)
    matches = []
    for keyword, mode, exp_type, merch_cat, store_name, mc_derived_pos in rules:
        key_norm = normalize_match_text(keyword)
        if not key_norm:
            continue
        if (
            keyword in desc
            or key_norm in desc_norm
            or ordered_token_match(keyword, desc)
            or unordered_token_match(keyword, desc)
        ):
            matches.append((keyword, mode, exp_type, merch_cat, store_name, mc_derived_pos))
    if matches:
        # First match wins (top-to-bottom in SB Mapping sheet).
        chosen = matches[0]
        _, mode, exp_type, merch_cat, store_name, mc_derived_pos = chosen
        exp_type = derive_expense_type(exp_type, amount)
        merch_cat = derive_merchant_category(merch_cat, mc_derived_pos, amount)
        return mode, exp_type, merch_cat, store_name
    # fallback is a tuple like: (mode, expense_type, merchant_category, store_name, mc_derived_pos)
    fb_mode, fb_exp, fb_mc, fb_store, fb_mc_pos = fallback
    fb_exp = derive_expense_type(fb_exp, amount)
    fb_mc = derive_merchant_category(fb_mc, fb_mc_pos, amount)
    return fb_mode, fb_exp, fb_mc, fb_store


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
    return classify_sb_description(description, amount, [], fallback)


def extract_yes_period(text):
    text = text or ""
    # Preferred: explicit "as on dd/mm/yyyy"
    m = re.search(r"Account Relationship Summary as on\s*(\d{2}/\d{2}/\d{4})", text, re.I)
    if m:
        d = parse_date(m.group(1))
        if d:
            return format_period_from_date(d)
    # Common: "Statement Period : dd/mm/yyyy - dd/mm/yyyy"
    m = re.search(
        r"Statement\s+Period\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})\s*(?:to|\-)\s*(\d{2}/\d{2}/\d{4})",
        text,
        re.I,
    )
    if m:
        d = parse_date(m.group(2)) or parse_date(m.group(1))
        if d:
            return format_period_from_date(d)
    # Alternate: "Period: 01 Feb 2026 - 28 Feb 2026"
    m = re.search(
        r"\bPeriod\s*:\s*(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s*-\s*(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})",
        text,
        re.I,
    )
    if m:
        d = parse_date(m.group(2)) or parse_date(m.group(1))
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


BANK_NAME_MAP_SHEET = "Bank Name map"
TRANS_DATE_SHEET = "Trans Date"


def _load_bank_name_map_rows(mapping_file):
    if not os.path.exists(mapping_file):
        return []
    try:
        df = pd.read_excel(mapping_file, sheet_name=BANK_NAME_MAP_SHEET)
    except Exception:
        return []
    cols = list(df.columns)
    if len(cols) >= 3:
        df = df.rename(columns={cols[0]: "Bank PDF", cols[1]: "Text", cols[2]: "Output"})
    rows = []
    for _, r in df.iterrows():
        bank_pdf = clean_text(r.get("Bank PDF", ""))
        text = clean_text(r.get("Text", ""))
        output = clean_text(r.get("Output", ""))
        if not bank_pdf or not text or not output:
            continue
        rows.append({"bank_pdf": bank_pdf, "text": text, "output": output})
    return rows


def load_trans_date_field_map(mapping_file):
    """
    Load bank -> table header config for transaction-date + opening balance extraction.
    Sheet: "Trans Date" with columns:
      - Bank Name
      - Transc Date Field
      - Decription Col Name
      - Description Value
      - OB Fall Back field
    """
    mapping = {}
    if not os.path.exists(mapping_file):
        return mapping
    try:
        df = pd.read_excel(mapping_file, sheet_name=TRANS_DATE_SHEET)
    except Exception:
        return mapping
    for _, row in df.iterrows():
        bank = clean_text(row.get("Bank Name", ""))
        trans_field = clean_text(row.get("Transc Date Field", ""))
        desc_col = clean_text(row.get("Decription Col Name", ""))
        desc_value = clean_text(row.get("Description Value", ""))
        ob_fallback = clean_text(row.get("OB Fall Back field", ""))
        if bank:
            mapping[bank.strip().lower()] = {
                "trans_date_field": trans_field,
                "ob_desc_col": desc_col,
                "ob_desc_value": desc_value,
                "ob_fallback_field": ob_fallback,
            }
    return mapping


def extract_opening_balance_from_pdf(pdf_path, bank_key, cfg):
    """
    Derive opening balance using the Trans Date sheet config:
    1) Find a transaction table and locate the Description column (cfg['ob_desc_col']).
       Find the row where that column contains cfg['ob_desc_value'], then read Balance.
    2) If not found, search pdf text for cfg['ob_fallback_field'] (or cfg['ob_desc_value'])
       and extract the first amount near it.
    """
    bank_key = (bank_key or "").strip().lower()
    cfg = cfg or {}
    desc_col_name = clean_text(cfg.get("ob_desc_col", ""))
    desc_value = clean_text(cfg.get("ob_desc_value", ""))
    fallback_field = clean_text(cfg.get("ob_fallback_field", "")) or desc_value

    def _find_amount_near(text, needle):
        if not text or not needle:
            return None
        t = re.sub(r"\s+", " ", text)
        n = re.sub(r"\s+", " ", needle).strip()
        m = re.search(re.escape(n) + r".{0,80}?([0-9][0-9,]*\.\d{2})", t, flags=re.IGNORECASE)
        if m:
            return parse_amount(m.group(1))
        m = re.search(r"([0-9][0-9,]*\.\d{2}).{0,40}?" + re.escape(n), t, flags=re.IGNORECASE)
        if m:
            return parse_amount(m.group(1))
        return None

    def _find_balance_on_line(text, desc_match):
        if not text or not desc_match:
            return None
        for raw in (text or "").split("\n"):
            line = clean_text(raw)
            if not line:
                continue
            if re.search(desc_match, line, flags=re.IGNORECASE):
                nums = extract_amounts_with_decimals(line)
                if nums:
                    return parse_amount(nums[-1])
        return None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Some banks (e.g. SBI) have the transaction overview on later pages.
            max_pages = 6 if bank_key == "sbi" else 3
            pages = pdf.pages[:max_pages]
            for page in pages:
                for table in extract_tables(page):
                    if not table or not table[0]:
                        continue
                    header = table[0]
                    if not desc_col_name:
                        continue
                    desc_idx = find_column_index_by_header(header, desc_col_name)
                    bal_idx = None
                    for i, cell in enumerate(header):
                        if "balance" in _norm_header_text(cell):
                            bal_idx = i
                            break
                    if desc_idx is None or bal_idx is None:
                        continue
                    for row in table[1:]:
                        if not row or max(desc_idx, bal_idx) >= len(row):
                            continue
                        cell_desc = clean_text(row[desc_idx])
                        if not cell_desc:
                            continue
                        # If Description Value isn't provided, assume the first non-empty row is opening balance.
                        if not desc_value:
                            bal = parse_amount(row[bal_idx])
                            if bal is not None:
                                return bal
                        if desc_value and re.search(re.escape(desc_value), cell_desc, flags=re.IGNORECASE):
                            bal = parse_amount(row[bal_idx])
                            if bal is not None:
                                return bal
            text = "\n".join((p.extract_text() or "") for p in pages)
            # Try matching the configured description value as a line item (common for YES: "B/F ...").
            if desc_value:
                bal = _find_balance_on_line(text, re.escape(desc_value))
                if bal is not None:
                    return bal
            # SBI statements may have OCR noise around "Opening Balance"; be lenient.
            if bank_key == "sbi":
                for raw in text.split("\n"):
                    line = clean_text(raw)
                    # Common pattern: "... on 01-02-26: 62344.19 ..."
                    m = re.search(r"\bon\s+\d{2}-\d{2}-\d{2}\s*:\s*([0-9][0-9,]*\.\d{2})", line, re.I)
                    if m:
                        val = parse_amount(m.group(1))
                        if val is not None:
                            return val
                    u = line.upper()
                    if "OPEN" in u and ("BAL" in u or "B" in u):
                        nums = extract_amounts_with_decimals(line)
                        if nums:
                            return parse_amount(nums[0])
            return _find_amount_near(text, fallback_field)
    except Exception:
        return None


def _norm_header_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def find_column_index_by_header(header_row, wanted_header):
    """
    Find a column index in a header row by matching the wanted_header text.
    Handles whitespace/newline differences.
    Returns 0-based index or None.
    """
    wanted = _norm_header_text(wanted_header)
    if not wanted:
        return None
    for idx, cell in enumerate(header_row or []):
        if _norm_header_text(cell) == wanted:
            return idx
    for idx, cell in enumerate(header_row or []):
        # allow containment match for slightly verbose headers
        cell_norm = _norm_header_text(cell)
        if wanted and cell_norm and wanted in cell_norm:
            return idx
    return None


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
    # Join without inserting a space when it looks like a wrapped single token.
    if desc and " " not in frag:
        last = desc[-1]
        first = frag[0]
        # Hyphenated wrap like "Transac-" + "tion" => "Transaction"
        if last == "-":
            return f"{desc[:-1]}{frag}"
        if last.isalpha() and first.isalpha():
            # Common PDF wraps continue with lowercase; also allow upper->lower transitions.
            if first.islower() or last.islower():
                return f"{desc}{frag}"
            # If both are uppercase, keep them separate words (e.g. "NEFT" + "IMPS").
    return f"{desc} {frag}".strip()


def detect_bank_from_text(text):
    t = (text or "").upper()
    if "AXIS BANK" in t:
        return "Axis"
    if "HDFC BANK" in t:
        return "HDFC"
    if "ICICI BANK" in t:
        return "ICICI"
    if "YES BANK" in t:
        return "Yes"
    if "STATE BANK OF INDIA" in t or re.search(r"\bSBI\b", t):
        return "SBI"
    # HSBC statements can contain other bank names inside transaction narrations; detect HSBC before IDFC.
    if "HSBC" in t:
        return "HSBC"
    if "IDFC FIRST BANK" in t or "IDFC BANK" in t:
        return "IDFC"
    if "INDUSIND" in t:
        return "IndusInd"
    return None


def detect_pdf_context(pdf_path):
    path_lower = (pdf_path or "").lower()
    # File-name hints are more reliable than scanning narration text (which may contain other bank names).
    bank_from_path = None
    if "axis" in path_lower:
        bank_from_path = "Axis"
    elif "hdfc" in path_lower:
        bank_from_path = "HDFC"
    elif "icici" in path_lower:
        bank_from_path = "ICICI"
    elif "idfc" in path_lower:
        bank_from_path = "IDFC"
    elif "yes" in path_lower:
        bank_from_path = "Yes"
    elif "sbi" in path_lower:
        bank_from_path = "SBI"
    elif "hsbc" in path_lower:
        bank_from_path = "HSBC"
    elif "indusind" in path_lower:
        bank_from_path = "IndusInd"

    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages[:2]:
                text += "\n" + (page.extract_text() or "")
    except Exception:
        return {"bank": None, "customer_name": None}
    bank_from_text = detect_bank_from_text(text)
    bank = bank_from_text or bank_from_path
    t = text.upper()
    customer_name = None
    if "PRIYANKA JAIN" in t:
        customer_name = "PRIYANKA JAIN"
    elif "ABHISHEK JAIN" in t:
        customer_name = "ABHISHEK JAIN"
    mapped_account_name = None
    try:
        mapping_rows = _load_bank_name_map_rows(MAPPING_FILE)
        pdf_comp = re.sub(r"\s+", "", text or "").upper()
        for row in mapping_rows:
            if bank and clean_text(row["bank_pdf"]).lower() not in bank.lower():
                continue
            needle = re.sub(r"\s+", "", row["text"]).upper()
            if needle and needle in pdf_comp:
                mapped_account_name = row["output"]
                break
    except Exception:
        mapped_account_name = None
    return {"bank": bank, "customer_name": customer_name, "mapped_account_name": mapped_account_name}


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
    if bank_key == "sbi":
        return parse_sbi
    if bank_key == "hsbc":
        return parse_hsbc
    if bank_key == "indusind":
        return parse_indusind
    return None


def format_sheet(workbook, worksheet, df, start_row=0, start_col=0):
    nrows, ncols = df.shape
    if ncols == 0:
        return

    header_fmt = workbook.add_format({"bold": True, "bg_color": "#F4B183", "border": 1})
    cell_fmt = workbook.add_format({"border": 1})
    wrap_cell_fmt = workbook.add_format({"border": 1, "text_wrap": True})
    amt_fmt = workbook.add_format({"border": 1, "num_format": "#,##0.00"})

    # Header styling
    for c, col in enumerate(df.columns):
        worksheet.write(start_row, start_col + c, col, header_fmt)

    # Borders + numeric formats for all data cells
    numeric_cols = {"Amount", "Balance"}
    desc_cols = {"Description"}
    for r in range(1, nrows + 1):
        for c, col in enumerate(df.columns):
            val = df.iloc[r - 1, c]
            if pd.isna(val):
                fmt = wrap_cell_fmt if col in desc_cols else cell_fmt
                worksheet.write_blank(start_row + r, start_col + c, None, fmt)
            elif col in numeric_cols:
                try:
                    worksheet.write_number(start_row + r, start_col + c, float(val), amt_fmt)
                except Exception:
                    fmt = wrap_cell_fmt if col in desc_cols else cell_fmt
                    worksheet.write(start_row + r, start_col + c, val, fmt)
            else:
                fmt = wrap_cell_fmt if col in desc_cols else cell_fmt
                worksheet.write(start_row + r, start_col + c, val, fmt)

    # Auto width
    for c, col in enumerate(df.columns):
        max_len = len(str(col))
        if nrows:
            col_max = df[col].map(lambda x: len(str(x)) if pd.notna(x) else 0).max()
            max_len = max(max_len, int(col_max))
        worksheet.set_column(start_col + c, start_col + c, min(max_len + 2, 60))

    # Approx auto row height
    worksheet.set_row(start_row, 20)
    for r in range(1, nrows + 1):
        row_vals = ["" if pd.isna(v) else str(v) for v in df.iloc[r - 1].tolist()]
        longest = max((len(v) for v in row_vals), default=0)
        lines = max(1, (longest // 45) + 1)
        worksheet.set_row(start_row + r, min(15 * lines, 60))
    if start_row == 0 and start_col == 0:
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, nrows, ncols - 1)


def parse_axis(pdf_path, trans_date_map=None):
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
            if re.search(r"Statement\s+for\s+account\s+no\.?", text, flags=re.IGNORECASE):
                statement_started = True
                m = re.search(
                    r"Statement\s+for\s+account\s+no\.?.*?from\s+(\d{2}-\d{2}-\d{4})\s+to\s+(\d{2}-\d{2}-\d{4})",
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
                date_idx = 0
                if trans_date_map:
                    mapped = trans_date_map.get("axis")
                    idx = find_column_index_by_header(header, mapped) if mapped else None
                    if idx is not None:
                        date_idx = idx
                for row in table[1:]:
                    if not row or len(row) < 5:
                        continue
                    date_cell = row[date_idx] if date_idx < len(row) else row[0]
                    date = parse_date(date_cell)
                    if not date:
                        # Axis tables sometimes wrap "Transaction Details" to the next line as a new row
                        # with a blank date cell (e.g. card number). Treat it as a continuation row.
                        if records:
                            # Pull continuation text from any non-numeric columns; some PDFs place the wrapped
                            # part under Chq/Ref No instead of Transaction Details.
                            cont_parts = []
                            for i, cell in enumerate(row):
                                if i == date_idx:
                                    continue
                                s = "" if cell is None else str(cell)
                                s = s.replace("\n", " ").strip()
                                if not s:
                                    continue
                                # Skip values that look like amounts/balances, but keep long digit strings
                                # (often card/account numbers) which may otherwise parse as numbers.
                                digits_only = re.sub(r"\s+", "", s)
                                if re.fullmatch(r"\d{10,20}", digits_only):
                                    cont_parts.append(digits_only)
                                    continue
                                if parse_amount(s) is not None:
                                    continue
                                cont_parts.append(s)
                            cont = " ".join(cont_parts).strip()
                            # Only append when the numeric columns are empty/zero-ish.
                            debit_c = parse_amount(row[3]) if len(row) > 3 else None
                            credit_c = parse_amount(row[4]) if len(row) > 4 else None
                            if cont and (debit_c in (None, 0) and credit_c in (None, 0)):
                                records[-1]["Description"] = (records[-1]["Description"] + " " + cont).strip()
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
                            # Do not emit an "Opening Balance" transaction row for Axis.
                            prev_balance = bal
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
                    elif current:
                        # Another common continuation is a wrapped card/account number on its own line.
                        digits = re.sub(r"\s+", "", line)
                        if re.fullmatch(r"\d{10,20}", digits):
                            current["Description"] = (current["Description"] + " " + digits).strip()
    return records


def parse_hdfc(pdf_path, trans_date_map=None):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "HDFC")
        for page in pdf.pages:
            for table in extract_tables(page):
                header = table[0] if table else []
                if not header or "Txn Date" not in " ".join([str(c) for c in header]):
                    continue
                date_idx = 0
                if trans_date_map:
                    mapped = trans_date_map.get("hdfc")
                    idx = find_column_index_by_header(header, mapped) if mapped else None
                    if idx is not None:
                        date_idx = idx
                for row in table[1:]:
                    if not row or len(row) < 5:
                        continue
                    date = parse_date(row[date_idx] if date_idx < len(row) else row[0])
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
        if not records:
            # Alternate ICICI layout:
            # "Statement of Transactions in Saving Account ...", text row like:
            # 1 30.03.2026 100901502072:Int.Pd:31-12-2025 to 29-03-2026 5.00 835.69
            seen = set()
            for page in pdf.pages:
                text = page.extract_text() or ""
                for raw in text.split("\n"):
                    line = clean_text(raw)
                    m = re.match(
                        r"^(?:\d+\s+)?(\d{2}\.\d{2}\.\d{4})\s+(.+?)\s+([0-9,]+\.\d{2})\s+([0-9,]+\.\d{2})\s*$",
                        line,
                    )
                    if not m:
                        continue
                    date_s, desc, amount_s, balance_s = m.groups()
                    date = parse_date(date_s)
                    amount = parse_amount(amount_s)
                    balance = parse_amount(balance_s)
                    if not date or amount is None:
                        continue
                    key = (date, desc, amount, balance)
                    if key in seen:
                        continue
                    seen.add(key)
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


def parse_idfc(pdf_path, trans_date_map=None):
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
                date_idx = None
                if trans_date_map:
                    mapped = trans_date_map.get("idfc")
                    idx = find_column_index_by_header(table[header_idx], mapped) if mapped else None
                    if idx is not None:
                        date_idx = idx

                for row in table[header_idx + 1 :]:
                    if not row or len(row) < 7:
                        continue
                    if date_idx is not None and date_idx < len(row):
                        date = parse_date(row[date_idx])
                    else:
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


def parse_yes(pdf_path, trans_date_map=None):
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
                header_row = table[header_idx] or []
                header_cells = [clean_text(c).lower() for c in header_row]

                def find_col(*needles):
                    for idx, cell in enumerate(header_cells):
                        for n in needles:
                            if n in cell:
                                return idx
                    return None

                date_col = find_col("transaction date", "tran date", "transaction", "date")
                if trans_date_map:
                    mapped = trans_date_map.get("yes")
                    idx = find_column_index_by_header(header_row, mapped) if mapped else None
                    if idx is not None:
                        date_col = idx
                desc_col = find_col("description")
                withdrawals_col = find_col("withdrawal")
                deposits_col = find_col("deposit")
                balance_col = find_col("running balance", "balance")

                for row in table[header_idx + 1 :]:
                    if not row or len(row) < 7:
                        continue
                    date_raw = None
                    if date_col is not None and date_col < len(row):
                        date_raw = row[date_col]
                    if date_raw is None:
                        date_raw = row[0]
                    date = parse_date(date_raw)
                    if not date:
                        continue

                    # Use Description column as-is from YES table and only merge wrapped lines.
                    raw_desc_cell = ""
                    if desc_col is not None and desc_col < len(row):
                        raw_desc_cell = str(row[desc_col] if row[desc_col] is not None else "")
                    desc = ""
                    for part in raw_desc_cell.split("\n"):
                        desc = append_yes_fragment(desc, part)
                    desc = clean_text(desc)

                    debit_raw = row[withdrawals_col] if withdrawals_col is not None and withdrawals_col < len(row) else (row[4] if len(row) > 4 else None)
                    credit_raw = row[deposits_col] if deposits_col is not None and deposits_col < len(row) else (row[5] if len(row) > 5 else None)
                    balance_raw = row[balance_col] if balance_col is not None and balance_col < len(row) else (row[6] if len(row) > 6 else None)
                    debit = parse_amount(debit_raw)
                    credit = parse_amount(credit_raw)
                    balance = parse_amount(balance_raw)

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


def parse_sbi(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "SBI")
        if period == "Unknown":
            m = re.search(r"As on\s+(\d{2}-\d{2}-\d{2})", first_text, re.I)
            if m:
                d = parse_date(m.group(1))
                if d:
                    period = format_period_from_date(d)

        for page in pdf.pages:
            text = page.extract_text() or ""
            if "TRANSACTION OVERVIEW" not in text.upper():
                continue
            lines = [clean_text(l) for l in text.split("\n") if clean_text(l)]
            in_overview = False
            for line in lines:
                if "TRANSACTION OVERVIEW" in line.upper():
                    in_overview = True
                    continue
                if not in_overview:
                    continue
                if line.lower().startswith("your opening balance") or line.lower().startswith("your closing balance"):
                    continue
                if not re.match(r"^\d{2}-\d{2}-\d{2}\b", line):
                    continue
                # Prefer parsing the last 3 numeric columns at end: Credit Debit Balance.
                m_cols = re.search(
                    r"\s+([0-9][0-9,]*(?:\.\d{2})?)\s+([0-9][0-9,]*(?:\.\d{2})?)\s+([0-9][0-9,]*(?:\.\d{2})?)\s*$",
                    line,
                )
                if not m_cols:
                    continue
                date = parse_date(line.split()[0])
                if not date:
                    continue
                credit = parse_amount(m_cols.group(1))
                debit = parse_amount(m_cols.group(2))
                balance = parse_amount(m_cols.group(3))
                # Description: remove leading date and trailing amounts
                desc = re.sub(r"^\d{2}-\d{2}-\d{2}\s+", "", line).strip()
                desc = re.sub(
                    r"\s+" + re.escape(m_cols.group(1)) + r"\s+" + re.escape(m_cols.group(2)) + r"\s+" + re.escape(m_cols.group(3)) + r"$",
                    "",
                    desc,
                ).strip()
                amount = None
                if credit is not None and credit != 0:
                    amount = credit
                elif debit is not None and debit != 0:
                    amount = -debit
                if amount is None:
                    continue
                records.append(
                    {
                        "Period": period,
                        "Account": "SBI",
                        "Date": date,
                        "Description": desc,
                        "Amount": amount,
                        "Balance": balance,
                    }
                )
    return records


def parse_hsbc(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "HSBC")
        prev_balance = None

        for page in pdf.pages:
            text = page.extract_text() or ""
            if "DATE TRANSACTION DETAILS" not in text.upper():
                continue
            lines = [clean_text(l) for l in text.split("\n") if clean_text(l)]
            current_date = None
            current_parts = []

            def flush():
                nonlocal prev_balance, current_date, current_parts, records
                if not current_date:
                    return
                block = " ".join(current_parts).strip()
                if not block:
                    current_date = None
                    current_parts = []
                    return
                nums = extract_amounts_with_decimals(block)
                if not nums:
                    current_date = None
                    current_parts = []
                    return
                balance = parse_amount(nums[-1])
                amount_abs = parse_amount(nums[-2]) if len(nums) >= 2 else None

                # Description should come only from Transaction Details column: remove trailing amount+balance.
                desc = re.sub(r"\s+" + re.escape(nums[-1]) + r"$", "", block).strip()
                if amount_abs is not None:
                    desc = re.sub(r"\s+" + re.escape(nums[-2]) + r"\s*$", "", desc).strip()

                u = desc.upper()
                if "BALANCE BROUGHT FORWARD" in u and balance is not None:
                    prev_balance = balance
                    current_date = None
                    current_parts = []
                    return
                if "CLOSING BALANCE" in u:
                    current_date = None
                    current_parts = []
                    return

                if balance is None or amount_abs is None:
                    current_date = None
                    current_parts = []
                    return
                if prev_balance is None:
                    prev_balance = balance
                    current_date = None
                    current_parts = []
                    return

                amount = amount_abs if balance > prev_balance else -amount_abs
                prev_balance = balance
                records.append(
                    {
                        "Period": period,
                        "Account": "HSBC",
                        "Date": current_date,
                        "Description": desc,
                        "Amount": amount,
                        "Balance": balance,
                    }
                )
                current_date = None
                current_parts = []

            for line in lines:
                # Transactions can start with either ddMonYYYY (e.g. 03Feb2026) OR yyyy/mm/dd (e.g. 2026/01/26).
                if re.match(r"^\d{2}[A-Za-z]{3}\d{4}\b", line) or re.match(r"^\d{4}/\d{2}/\d{2}\b", line):
                    flush()
                    date_token = line.split()[0]
                    current_date = parse_date(date_token)
                    tail = line[len(date_token) :].strip()
                    current_parts = [tail] if tail else []
                    continue
                if current_date:
                    # Closing balance footer should not be appended to the last transaction.
                    if line.upper().startswith("CLOSING BALANCE"):
                        flush()
                        current_date = None
                        current_parts = []
                        continue
                    current_parts.append(line)
            flush()
    return records


def parse_indusind(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        period = extract_period(first_text, "IndusInd")
        in_history = False
        prev_balance = None
        current_date = None
        current_parts = []
        current_amount_hint = None
        current_balance_hint = None

        def flush():
            nonlocal prev_balance, current_date, current_parts, records, current_amount_hint, current_balance_hint
            if not current_date:
                return
            particulars = clean_text(" ".join(current_parts))
            if not particulars:
                current_date = None
                current_parts = []
                current_amount_hint = None
                current_balance_hint = None
                return
            # Prefer using the parsed hints from the first row line so we don't pollute
            # Particulars with adjacent numeric columns.
            balance = current_balance_hint
            amt_abs = current_amount_hint
            desc = particulars

            # Skip brought forward / carried forward markers but use them to seed balance.
            u = desc.upper()
            if ("BROUGHT FORWARD" in u or "CARRIED FORWARD" in u) and balance is not None:
                prev_balance = balance
                current_date = None
                current_parts = []
                current_amount_hint = None
                current_balance_hint = None
                return

            if balance is None:
                current_date = None
                current_parts = []
                current_amount_hint = None
                current_balance_hint = None
                return

            if amt_abs is None:
                prev_balance = balance
                current_date = None
                current_parts = []
                current_amount_hint = None
                current_balance_hint = None
                return

            if prev_balance is None:
                prev_balance = balance
                current_date = None
                current_parts = []
                current_amount_hint = None
                current_balance_hint = None
                return

            amount = amt_abs if balance > prev_balance else -amt_abs
            prev_balance = balance
            records.append(
                {
                    "Period": period,
                    "Account": "IndusInd",
                    "Date": current_date,
                    "Description": desc,
                    "Amount": amount,
                    "Balance": balance,
                }
            )
            current_date = None
            current_parts = []
            current_amount_hint = None
            current_balance_hint = None

        for page in pdf.pages:
            text = page.extract_text() or ""
            if not in_history and re.search(
                r"Transaction History for Savings Account", text, re.I
            ):
                in_history = True
            if not in_history:
                continue
            lines = [clean_text(l) for l in text.split("\n") if clean_text(l)]
            for line in lines:
                if re.search(r"Transaction History for Savings Account", line, re.I):
                    continue
                if re.search(r"^Date\b", line, re.I) and "withdraw" in line.lower() and "deposit" in line.lower():
                    continue
                if re.search(r"CUSTOMER ID", line, re.I) and "ACCOUNT NUMBER" in line.upper():
                    continue
                if re.search(r"^\d{2}-[A-Za-z]{3}-\d{4}\b", line):
                    flush()
                    current_date = parse_date(line.split()[0])
                    tail = line[len(line.split()[0]) :].strip()
                    # Try to peel off trailing "<amount> <balance>" from the first row line.
                    m_cols = re.search(
                        r"\s+([0-9][0-9,]*\.\d{2})\s+([0-9][0-9,]*\.\d{2})\s*$",
                        tail,
                    )
                    if m_cols:
                        current_amount_hint = parse_amount(m_cols.group(1))
                        current_balance_hint = parse_amount(m_cols.group(2))
                        tail = re.sub(
                            r"\s+" + re.escape(m_cols.group(1)) + r"\s+" + re.escape(m_cols.group(2)) + r"\s*$",
                            "",
                            tail,
                        ).strip()
                    else:
                        # Brought forward / carried forward lines often only show the balance.
                        m_bal = re.search(r"([0-9][0-9,]*\.\d{2})\s*$", tail)
                        current_balance_hint = parse_amount(m_bal.group(1)) if m_bal else None
                        current_amount_hint = None
                    current_parts = [tail] if tail else []
                    continue
                if current_date:
                    # Stop if we reached an interest certificate section on later pages.
                    if "INTEREST CERTIFICATE" in line.upper():
                        flush()
                        in_history = False
                        current_date = None
                        current_parts = []
                        break
                    # Continuation lines belong to "Particulars" column; join wrapped words without extra spaces.
                    if current_parts:
                        current_parts[-1] = append_wrapped_fragment(current_parts[-1], line)
                    else:
                        current_parts.append(line)
            flush()
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
    if "sbi" in path:
        return parse_sbi
    if "hsbc" in path:
        return parse_hsbc
    if "indusind" in path:
        return parse_indusind
    return None


def main():
    import sys
    print("=" * 70)
    print("SAVINGS ACCOUNT MASTER PARSER")
    print("=" * 70)
    print(f"Scanning: {BASE_DIR}")

    trans_date_map = load_trans_date_field_map(MAPPING_FILE)

    all_records = []
    opening_balance_by_account = {}
    pdf_paths = []
    if len(sys.argv) > 1:
        pdf_paths = [sys.argv[1]]
    else:
        for root, dirs, files in os.walk(BASE_DIR):
            dirs[:] = [d for d in dirs if d.lower() not in {"archive", "archived"}]
            for f in files:
                if f.lower().endswith(".pdf"):
                    pdf_paths.append(os.path.join(root, f))

    for pdf_path in pdf_paths:
        print(f"\\n📄 Processing: {os.path.basename(pdf_path)}")
        ctx = detect_pdf_context(pdf_path)
        parser = get_parser_by_bank(ctx["bank"]) or get_parser(pdf_path)
        if not parser:
            print("   ⚠️ No parser found for this file")
            continue
        try:
            # Some parsers support Trans Date header mapping for table extraction.
            try:
                records = parser(pdf_path, trans_date_map=trans_date_map)
            except TypeError:
                records = parser(pdf_path)
            if records:
                fallback_account = records[0].get("Account", "")
                account_name = ctx.get("mapped_account_name") or resolve_account_name(
                    ctx["bank"], ctx.get("customer_name"), fallback_account
                )
                for rec in records:
                    rec["Account"] = account_name
                bank_key = account_to_bank_key(ctx.get("bank") or account_name)
                cfg = trans_date_map.get((ctx.get("bank") or "").strip().lower()) or trans_date_map.get(bank_key)
                ob = extract_opening_balance_from_pdf(pdf_path, bank_key, cfg)
                if ob is not None and account_name not in opening_balance_by_account:
                    opening_balance_by_account[account_name] = ob
            print(f"   ✅ Extracted {len(records)} transactions")
            all_records.extend(records)
        except Exception as e:
            print(f"   ❌ Failed: {e}")

    if not all_records:
        print("\\nNo transactions found.")
        return

    df = pd.DataFrame(all_records)
    known_words = load_known_wrap_words()
    if "Description" in df.columns and known_words:
        df["Description"] = df["Description"].astype(str).map(lambda v: fix_spaced_known_words(v, known_words))
    df["_sort_date"] = pd.to_datetime(df["Date"], errors="coerce")
    # Period is derived from transaction Date (Mon-YYYY), not from statement headers.
    df.loc[df["_sort_date"].notna(), "Period"] = df.loc[df["_sort_date"].notna(), "_sort_date"].dt.strftime("%b-%Y")

    # Opening/Closing balances per account (for summary table at O2).
    tmp_bal = df.dropna(subset=["_sort_date"]).copy()
    if not tmp_bal.empty:
        tmp_bal["_acct"] = tmp_bal["Account"].astype(str)
        tmp_bal = tmp_bal.sort_values(by=["_acct", "_sort_date"])
        opening = (
            tmp_bal.dropna(subset=["Balance"])
            .groupby("_acct", dropna=False)
            .first(numeric_only=False)
            .reset_index()[["_acct", "Balance"]]
            .rename(columns={"Balance": "Opening Balance"})
        )
        closing = (
            tmp_bal.dropna(subset=["Balance"])
            .groupby("_acct", dropna=False)
            .last(numeric_only=False)
            .reset_index()[["_acct", "Balance"]]
            .rename(columns={"Balance": "Closing Balance"})
        )
        bank_bal_df = opening.merge(closing, on="_acct", how="outer").rename(columns={"_acct": "Bank Name"})
        if opening_balance_by_account:
            bank_bal_df["Opening Balance"] = bank_bal_df.apply(
                lambda r: float(opening_balance_by_account.get(r["Bank Name"], r["Opening Balance"])),
                axis=1,
            )
    else:
        bank_bal_df = pd.DataFrame(columns=["Bank Name", "Opening Balance", "Closing Balance"])

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

    def _sort_key_after_space(val: object) -> tuple[str, str]:
        """
        Sort helper: if the value has whitespace (e.g. "AJ YES"), sort primarily by the
        token after whitespace ("YES") so "AJ YES" and "PJ YES" group together.
        """
        s = "" if val is None else str(val).strip()
        parts = s.split()
        if len(parts) >= 2:
            return (parts[-1].lower(), " ".join(parts[:-1]).lower())
        return (s.lower(), "")

    # "Pivot" summary requested: group by Period + Account + Merchant Category, sum Amount.
    # (This is written as a table since Excel pivot tables aren't generated by xlsxwriter.)
    # We'll later filter to the dominant Period for the "SB Categorized Summary" sheet.
    pivot_df = pd.DataFrame()

    # Sort tables (A table, K table, O table) using the custom "after whitespace" rule.
    # Dominant period filter for SB Categorized Summary table at A1:
    # Keep only the Period that has the maximum number of transactions.
    dominant_period = None
    if not summary_df.empty and "Period" in summary_df.columns:
        vc = summary_df["Period"].value_counts(dropna=False)
        if not vc.empty:
            dominant_period = vc.index[0]

    summary_df_sheet = summary_df
    summary_source_df_sheet = summary_source_df
    if dominant_period is not None:
        summary_df_sheet = summary_df[summary_df["Period"] == dominant_period].reset_index(drop=True)
        summary_source_df_sheet = summary_source_df[summary_source_df["Period"] == dominant_period].reset_index(drop=True)

    if not summary_df_sheet.empty and "Account" in summary_df_sheet.columns:
        summary_df_sheet = summary_df_sheet.sort_values(
            by="Account",
            key=lambda s: s.map(lambda v: _sort_key_after_space(v)),
            kind="mergesort",
        ).reset_index(drop=True)

    # Spend analysis pivot based on the (filtered) summary table.
    pivot_df = (
        summary_df_sheet.groupby(["Period", "Account", "Merchant Category"], dropna=False)["Amount"]
        .sum()
        .reset_index()
        .rename(columns={"Amount": "Sum of Amount"})
    )
    if not pivot_df.empty:
        pivot_df = pivot_df.sort_values(
            by=["Period", "Account"],
            key=lambda s: s.map(lambda v: _sort_key_after_space(v)) if s.name == "Account" else s,
            kind="mergesort",
        ).reset_index(drop=True)

    # Monthly balance table should also be for the dominant period only, and include Period.
    if not summary_source_df_sheet.empty and {"Period", "Account", "Balance"}.issubset(summary_source_df_sheet.columns):
        bank_bal_df = (
            summary_source_df_sheet.sort_values(["Account", "Date"], kind="mergesort")
            .groupby(["Period", "Account"], dropna=False, as_index=False)
            .agg(Opening_Balance=("Balance", "first"), Closing_Balance=("Balance", "last"))
            .rename(columns={"Account": "Bank Name"})
        )
        bank_bal_df = bank_bal_df.rename(
            columns={"Opening_Balance": "Opening Balance", "Closing_Balance": "Closing Balance"}
        )
    else:
        bank_bal_df = pd.DataFrame(columns=["Period", "Bank Name", "Opening Balance", "Closing Balance"])

    # Replace opening balances from PDF-derived values when available.
    if not bank_bal_df.empty and opening_balance_by_account:
        bank_bal_df["Opening Balance"] = bank_bal_df.apply(
            lambda r: float(opening_balance_by_account.get(r["Bank Name"], r["Opening Balance"])),
            axis=1,
        )

    if not bank_bal_df.empty:
        bank_bal_df = bank_bal_df.sort_values(
            by=["Period", "Bank Name"],
            key=lambda s: s.map(lambda v: _sort_key_after_space(v)) if s.name == "Bank Name" else s,
            kind="mergesort",
        ).reset_index(drop=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="SB AC expenses", index=False)
        summary_df_sheet.to_excel(writer, sheet_name="SB Categorized Summary", index=False)
        workbook = writer.book
        format_sheet(workbook, writer.sheets["SB AC expenses"], df)
        format_sheet(workbook, writer.sheets["SB Categorized Summary"], summary_df_sheet)

        ws_sum = writer.sheets["SB Categorized Summary"]

        # Spend analysis table at K (with a merged yellow header row).
        spend_hdr_fmt = workbook.add_format(
            {
                "bold": True,
                "font_size": 14,
                "font_name": "Calibri",
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#FFD966",  # yellow
                "top": 5,
                "left": 5,
                "right": 5,
                "bottom": 1,
            }
        )
        spend_header_fmt = workbook.add_format({"bold": True, "bg_color": "#F4B183", "border": 1})
        spend_cell_fmt = workbook.add_format({"border": 1})
        spend_amt_fmt = workbook.add_format({"border": 1, "num_format": "#,##0.00"})
        spend_top_fmt = workbook.add_format({"top": 5, "border": 1})
        spend_bottom_fmt = workbook.add_format({"bottom": 5, "border": 1})
        spend_left_fmt = workbook.add_format({"left": 5, "border": 1})
        spend_right_fmt = workbook.add_format({"right": 5, "border": 1})

        spend_start_row = 0
        spend_start_col = 10  # K
        spend_cols = list(pivot_df.columns)
        spend_ncols = max(1, len(spend_cols))
        ws_sum.merge_range(
            spend_start_row,
            spend_start_col,
            spend_start_row,
            spend_start_col + spend_ncols - 1,
            "Spend analysis",
            spend_hdr_fmt,
        )
        pivot_df.to_excel(
            writer,
            sheet_name="SB Categorized Summary",
            index=False,
            startrow=spend_start_row + 1,
            startcol=spend_start_col,
        )
        format_sheet(
            workbook,
            ws_sum,
            pivot_df,
            start_row=spend_start_row + 1,
            start_col=spend_start_col,
        )
        # Thick outside border around the full spend-analysis block (merged header + table).
        spend_table_rows = 1 + (len(pivot_df) + 1)  # merged header + (header + data)
        spend_end_row = spend_start_row + spend_table_rows - 1
        spend_end_col = spend_start_col + spend_ncols - 1
        # top border (merge row already thick). left/right borders for all rows:
        ws_sum.conditional_format(
            spend_start_row,
            spend_start_col,
            spend_end_row,
            spend_start_col,
            {"type": "no_errors", "format": spend_left_fmt},
        )
        ws_sum.conditional_format(
            spend_start_row,
            spend_end_col,
            spend_end_row,
            spend_end_col,
            {"type": "no_errors", "format": spend_right_fmt},
        )
        # bottom border on last row (only applies where cells exist).
        ws_sum.conditional_format(
            spend_end_row,
            spend_start_col,
            spend_end_row,
            spend_end_col,
            {"type": "no_errors", "format": spend_bottom_fmt},
        )

        # Monthly Balance table at O (with a merged yellow header row).
        bal_hdr_fmt = workbook.add_format(
            {
                "bold": True,
                "font_size": 14,
                "font_name": "Calibri",
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#FFD966",
                "top": 5,
                "left": 5,
                "right": 5,
                "bottom": 1,
            }
        )
        bal_header_fmt = workbook.add_format({"bold": True, "bg_color": "#F4B183", "border": 1})
        bal_cell_fmt = workbook.add_format({"border": 1})
        bal_amt_fmt = workbook.add_format({"border": 1, "num_format": "#,##0.00"})
        bal_diff_red = workbook.add_format({"border": 1, "num_format": "#,##0.00", "font_color": "#C00000"})
        bal_diff_green = workbook.add_format({"border": 1, "num_format": "#,##0.00", "font_color": "#006100"})
        bal_left_fmt = workbook.add_format({"left": 5, "border": 1})
        bal_right_fmt = workbook.add_format({"right": 5, "border": 1})
        bal_bottom_fmt = workbook.add_format({"bottom": 5, "border": 1})

        bal_start_row = 0
        bal_start_col = 14  # O
        bal_headers = ["Period", "Bank Name", "Opening Balance", "Closing Balance", "Difference"]
        bal_ncols = len(bal_headers)
        ws_sum.merge_range(
            bal_start_row,
            bal_start_col,
            bal_start_row,
            bal_start_col + bal_ncols - 1,
            "Monthly Balance",
            bal_hdr_fmt,
        )
        for c, h in enumerate(bal_headers):
            ws_sum.write(bal_start_row + 1, bal_start_col + c, h, bal_header_fmt)
        for r in range(len(bank_bal_df)):
            row = bank_bal_df.iloc[r]
            ws_sum.write(bal_start_row + 2 + r, bal_start_col + 0, row.get("Period"), bal_cell_fmt)
            ws_sum.write(bal_start_row + 2 + r, bal_start_col + 1, row.get("Bank Name"), bal_cell_fmt)
            for ckey, cidx in (("Opening Balance", 2), ("Closing Balance", 3)):
                val = row.get(ckey)
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    ws_sum.write_blank(bal_start_row + 2 + r, bal_start_col + cidx, None, bal_cell_fmt)
                else:
                    try:
                        ws_sum.write_number(bal_start_row + 2 + r, bal_start_col + cidx, float(val), bal_amt_fmt)
                    except Exception:
                        ws_sum.write(bal_start_row + 2 + r, bal_start_col + cidx, val, bal_cell_fmt)
            # Difference formula: Closing - Opening
            row_excel = bal_start_row + 2 + r + 1
            ws_sum.write_formula(
                bal_start_row + 2 + r,
                bal_start_col + 4,
                f"=R{row_excel}-Q{row_excel}",
                bal_amt_fmt,
            )
        if len(bank_bal_df) > 0:
            diff_first = bal_start_row + 2
            diff_last = bal_start_row + 1 + len(bank_bal_df)
            diff_col = bal_start_col + 4
            ws_sum.conditional_format(
                diff_first,
                diff_col,
                diff_last,
                diff_col,
                {"type": "cell", "criteria": "<", "value": 0, "format": bal_diff_red},
            )
            ws_sum.conditional_format(
                diff_first,
                diff_col,
                diff_last,
                diff_col,
                {"type": "cell", "criteria": ">", "value": 0, "format": bal_diff_green},
            )
        # Thick outside border around the full monthly balance block.
        bal_table_rows = 1 + (1 + len(bank_bal_df))  # merged header + (header + data)
        bal_end_row = bal_start_row + bal_table_rows - 1
        bal_end_col = bal_start_col + bal_ncols - 1
        ws_sum.conditional_format(
            bal_start_row,
            bal_start_col,
            bal_end_row,
            bal_start_col,
            {"type": "no_errors", "format": bal_left_fmt},
        )
        ws_sum.conditional_format(
            bal_start_row,
            bal_end_col,
            bal_end_row,
            bal_end_col,
            {"type": "no_errors", "format": bal_right_fmt},
        )
        ws_sum.conditional_format(
            bal_end_row,
            bal_start_col,
            bal_end_row,
            bal_end_col,
            {"type": "no_errors", "format": bal_bottom_fmt},
        )

        # Auto-size O-S for readability.
        for c in range(bal_start_col, bal_start_col + bal_ncols):
            ws_sum.set_column(c, c, 22)

    print("\\n======================================================================")
    print("✅ SB AGGREGATION COMPLETE")
    print("======================================================================")
    print(f"Total PDFs:            {len(pdf_paths)}")
    print(f"Total Transactions:    {len(df)}")
    print(f"Output File:           {OUTPUT_FILE}")
    print("======================================================================")


if __name__ == "__main__":
    main()
