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
    # Look for "Statement Date : 13/01/2026"
    match = re.search(r"Statement Date\s*:\s*(\d{2}/\d{2}/\d{4})", text)
    if match:
        dt = datetime.strptime(match.group(1), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    
    return ""

def parse_uni_gold_upi_cc_pdf(pdf_path):
    """
    Parse Uni Gold UPI Card PDF
    Format: DD/MM/YYYY REFNO UPI-DESCRIPTION INR AMOUNT AMOUNT DR
    """
    transactions = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

        period = extract_period(full_text)

        # Pattern for UPI transactions
        # Example: 23/12/2025 F9644Z UPI-PARAS PAAL INR 20.00 20.00 DR
        pattern = r"(\d{2}/\d{2}/\d{4})\s+([A-Z0-9]+)\s+UPI-(.+?)\s+INR\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})\s+(DR|CR)"

        for match in re.findall(pattern, full_text):
            date, ref, desc, amount, txn_type = match

            amount = float(amount.replace(",", ""))

            # CR means refund - make it negative
            if txn_type == "CR":
                amount = -amount

            transactions.append({
                "Period": period,
                "Account": "Uni Gold Card UPI",
                "Date": date,
                "Description": f"UPI-{clean_description(desc)}",
                "Amount": amount,
                "Type": txn_type
            })

    return transactions
