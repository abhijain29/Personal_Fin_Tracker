import pdfplumber
import re
from datetime import datetime

def clean_description(text):
    text = text.replace("|", " ")
    text = text.replace("_", " ")
    text = text.replace("*", " ")
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def extract_period(text):
    text_norm = re.sub(r"\s+", " ", text)
    # Pattern 1: "Statement Date: 24/Nov/2025"
    match = re.search(r"Statement Date[:\-]?\s*(\d{1,2})/([A-Za-z]{3})/(\d{4})", text_norm)
    if match:
        day, month, year = match.groups()
        dt = datetime.strptime(f"{day} {month} {year}", "%d %b %Y")
        return dt.strftime("%b-%y")
    
    # Pattern 2: "Statement Date: 24/11/2025"
    match = re.search(r"Statement Date[:\-]?\s*(\d{2}/\d{2}/\d{4})", text_norm)
    if match:
        dt = datetime.strptime(match.group(1), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    
    # Pattern 3: Statement Period "25/Oct/2025 - 24/Nov/2025"
    match = re.search(r"Statement Period.*?(\d{1,2}/[A-Za-z]{3}/\d{4})\s*-\s*(\d{1,2}/[A-Za-z]{3}/\d{4})", text_norm)
    if match:
        # Use the end date
        dt = datetime.strptime(match.group(2), "%d/%b/%Y")
        return dt.strftime("%b-%y")

    # Pattern 4: Any date range in the document (fallback to end date)
    match = re.search(r"(\d{1,2}/[A-Za-z]{3}/\d{4})\s*-\s*(\d{1,2}/[A-Za-z]{3}/\d{4})", text_norm)
    if match:
        dt = datetime.strptime(match.group(2), "%d/%b/%Y")
        return dt.strftime("%b-%y")

    # Pattern 5: "24/Nov/2025" anywhere in first few lines
    match = re.search(r"(\d{1,2}/[A-Za-z]{3}/\d{4})", text_norm[:500])
    if match:
        dt = datetime.strptime(match.group(1), "%d/%b/%Y")
        return dt.strftime("%b-%y")
    
    return ""

def extract_idfc_transactions(pdf_path):
    """
    Parse IDFC FIRST Bank Credit Card PDF
    Date format in PDF: "08 Nov 25" (DD Mon YY)
    """
    transactions = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

        period = extract_period(full_text)

        # IDFC transactions can appear in 2 formats:
        # 1) Single-line: "08 Nov 25 SOME MERCHANT 216.00 DR"
        # 2) Multi-line (common in newer PDFs):
        #      "MERCHANT NAME,"
        #      "02 Mar 26 1,099.54 DR"
        #      "CITY"
        line_pattern = r"^(\d{2}\s+[A-Z][a-z]{2}\s+\d{2})\s+(.+?)\s+([\d,]+\.\d{2})\s+(DR|CR)\s*$"
        date_amt_only = r"^(\d{2}\s+[A-Z][a-z]{2}\s+\d{2})\s+([\d,]+\.\d{2})\s+(DR|CR)\s*$"

        lines = [ln.strip() for ln in full_text.splitlines() if ln and ln.strip()]
        i = 0
        while i < len(lines):
            ln = lines[i]

            m = re.match(line_pattern, ln)
            if m:
                date_str, desc, amount, txn_type = m.groups()
                date_obj = datetime.strptime(date_str, "%d %b %y")
                date = date_obj.strftime("%d/%m/%Y")
                amount_f = float(amount.replace(",", ""))
                if txn_type == "CR":
                    amount_f = -amount_f
                transactions.append(
                    {
                        "Period": period,
                        "Account": "IDFC FIRST CC",
                        "Date": date,
                        "Description": clean_description(desc),
                        "Amount": amount_f,
                        "Type": txn_type,
                    }
                )
                i += 1
                continue

            m2 = re.match(date_amt_only, ln)
            if m2:
                date_str, amount, txn_type = m2.groups()
                # Use previous line as description (merchant), and optionally append next line if it looks like a city.
                desc_parts = []
                if i - 1 >= 0:
                    prev = lines[i - 1]
                    # Avoid section headings.
                    if not re.match(r"^(Purchases|Payments|YOUR|Transaction|Card Number)", prev, re.I):
                        desc_parts.append(prev)
                if i + 1 < len(lines):
                    nxt = lines[i + 1]
                    if len(nxt) <= 40 and re.fullmatch(r"[A-Z\s\.\-&]+", nxt):
                        desc_parts.append(nxt)
                        i += 1  # consume next line

                desc = clean_description(" ".join(desc_parts).strip())
                if desc:
                    date_obj = datetime.strptime(date_str, "%d %b %y")
                    date = date_obj.strftime("%d/%m/%Y")
                    amount_f = float(amount.replace(",", ""))
                    if txn_type == "CR":
                        amount_f = -amount_f
                    transactions.append(
                        {
                            "Period": period,
                            "Account": "IDFC FIRST CC",
                            "Date": date,
                            "Description": desc,
                            "Amount": amount_f,
                            "Type": txn_type,
                        }
                    )
                i += 1
                continue

            i += 1

    return transactions
