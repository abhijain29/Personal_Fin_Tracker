import pdfplumber
import re
from datetime import datetime

def clean_description(text):
    text = text.replace("|", " ")
    text = text.replace("_", " ")
    text = text.replace("*", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def extract_period(text):
    match = re.search(r"Statement Date\s*:\s*(\d{2}/\d{2}/\d{4})", text)
    if match:
        dt = datetime.strptime(match.group(1), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    return ""

def parse_uni_gold_upi_cc_pdf(pdf_path):

    transactions = []

    with pdfplumber.open(pdf_path) as pdf:

        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"

        period = extract_period(full_text)

        pattern = r"(\d{2}/\d{2}/\d{4})\s+UPI\-(.+?)\s+INR\s+([\d,]+\.\d{2})"

        for match in re.findall(pattern, full_text):

            date, desc, amount = match

            amount = float(amount.replace(",", ""))

            transactions.append({
                "Period": period,
                "Account": "Uni Gold Card UPI",
                "Date": date,
                "Description": clean_description(desc),
                "Amount": amount,
                "Type": "Dr"
            })

    return transactions
