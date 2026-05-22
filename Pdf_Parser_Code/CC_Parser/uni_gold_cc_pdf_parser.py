import pdfplumber
import re
from datetime import datetime

def clean_description(text):
    # Remove special currency symbols and junk
    text = text.replace("₹", "")
    text = text.replace("|", " ")
    text = text.replace("_", " ")
    text = text.replace("*", " ")
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def extract_period(text):
    # Look for "Statement Date 14 Nov, 2025"
    match = re.search(r"Statement Date\s+(\d{1,2})\s+([A-Za-z]{3}),\s+(\d{4})", text)
    if match:
        day, month, year = match.groups()
        dt = datetime.strptime(f"{day} {month} {year}", "%d %b %Y")
        return dt.strftime("%b-%y")
    
    # Alternative: "Statement Date: 14/11/2025"
    match = re.search(r"Statement Date\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", text)
    if match:
        dt = datetime.strptime(match.group(1), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    
    return ""

def parse_uni_gold_cc_pdf(pdf_path):
    """
    Parse Uni Gold Card PDF
    Format: DD/MM/YYYY DESCRIPTION DEBIT/CREDIT ₹Amount
    """
    transactions = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += "\n" + text

        period = extract_period(full_text)

        # Pattern A (older format):
        # 01/11/2025 CCCPL FRONT OFFICE II HYDERABAD IN DEBIT ₹1,79,520
        pattern_a = re.compile(
            r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(DEBIT|CREDIT)\s+₹([\d,]+(?:\.\d{2})?)",
            re.IGNORECASE,
        )

        # Pattern B (newer BOBCARD/UNI statements):
        # 16/01/2026 R1673Z UPI-ZEPTO MARKETPLACE PRIVATE INR 1,020.00 1,020.00 DR
        # There are often two amount columns; the last amount before DR/CR is the transaction amount.
        pattern_b = re.compile(
            r"(\d{2}/\d{2}/\d{4})\s+([A-Z0-9]+)\s+(.+?)\s+INR\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+(DR|CR)\b",
            re.IGNORECASE,
        )

        def _to_float(s):
            try:
                return float(str(s).replace(",", ""))
            except Exception:
                return None

        for date, desc, txn_type, amount in pattern_a.findall(full_text):
            amount_val = _to_float(amount)
            if amount_val is None:
                continue
            if str(txn_type).upper() == "CREDIT":
                amount_val = -amount_val
            transactions.append(
                {
                    "Period": period,
                    "Account": "Uni Gold Card",
                    "Date": date,
                    "Description": clean_description(desc.strip()),
                    "Amount": amount_val,
                    "Type": "Dr" if str(txn_type).upper() == "DEBIT" else "Cr",
                }
            )

        for date, _ref, desc, _src_amt, amt, drcr in pattern_b.findall(full_text):
            amount_val = _to_float(amt)
            if amount_val is None:
                continue
            drcr_u = str(drcr).upper()
            if drcr_u == "CR":
                amount_val = -amount_val
            transactions.append(
                {
                    "Period": period,
                    "Account": "Uni Gold Card",
                    "Date": date,
                    "Description": clean_description(desc.strip()),
                    "Amount": amount_val,
                    "Type": "Dr" if drcr_u == "DR" else "Cr",
                }
            )

    except Exception as e:
        print(f"❌ Error parsing Uni Gold PDF: {e}")

    return transactions
