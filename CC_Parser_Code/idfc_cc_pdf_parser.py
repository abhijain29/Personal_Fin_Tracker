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

        # IDFC uses format: "08 Nov 25" not "08/11/25"
        # Pattern: DD Mon YY Description Amount DR/CR
        # Example: 08 Nov 25 CALIFORNIA BURRITO, HYDERABAD 216.00 DR
        pattern = r"(\d{2}\s+[A-Z][a-z]{2}\s+\d{2})\s+(.+?)\s+([\d,]+\.\d{2})\s+(DR|CR)"

        for match in re.findall(pattern, full_text):
            date_str, desc, amount, txn_type = match

            # Convert "08 Nov 25" to "08/11/2025"
            date_obj = datetime.strptime(date_str, "%d %b %y")
            date = date_obj.strftime("%d/%m/%Y")

            amount = float(amount.replace(",", ""))

            # CR means payment/refund - make it negative
            if txn_type == "CR":
                amount = -amount

            transactions.append({
                "Period": period,
                "Account": "IDFC FIRST CC",
                "Date": date,
                "Description": clean_description(desc),
                "Amount": amount,
                "Type": txn_type
            })

    return transactions
