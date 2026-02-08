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
    match = re.search(r"Statement Date[:\-]?\s*(\d{2}/\d{2}/\d{4})", text)
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

        # Pattern for lines like:
        # 01/11/2025 CCCPL FRONT OFFICE II HYDERABAD IN DEBIT ₹1,79,520
        pattern = re.compile(
            r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(DEBIT|CREDIT)\s+₹([\d,]+(?:\.\d{2})?)"
        )

        matches = pattern.findall(full_text)

        for match in matches:
            date = match[0]
            description = match[1].strip()
            txn_type = match[2]
            amount = match[3]

            # Clean amount
            amount_clean = amount.replace(",", "")
            
            try:
                amount_val = float(amount_clean)
            except:
                continue

            # CREDIT means payment/refund - make it negative
            if txn_type == "CREDIT":
                amount_val = -amount_val

            transactions.append({
                "Period": period,
                "Account": "Uni Gold Card",
                "Date": date,
                "Description": clean_description(description),
                "Amount": amount_val,
                "Type": "Dr" if txn_type == "DEBIT" else "Cr"
            })

    except Exception as e:
        print(f"❌ Error parsing Uni Gold PDF: {e}")

    return transactions
