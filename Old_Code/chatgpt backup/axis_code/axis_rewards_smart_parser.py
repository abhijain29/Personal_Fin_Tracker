import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import re


# ---------------- TEXT BASED PARSER ------------------

def text_based_parser(pdf_path):

    transactions = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            text = page.extract_text()

            if not text:
                continue

            pattern = r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,]+\.\d{2})\s+(Dr|Cr)"

            matches = re.findall(pattern, text)

            for m in matches:

                date, desc, amt, drcr = m

                amount = float(amt.replace(",", ""))

                if drcr == "Cr":
                    amount = -amount

                transactions.append({
                    "Account": "Axis Bank Rewards CC",
                    "Date": date,
                    "Description": desc.strip(),
                    "Amount": amount,
                    "Type": drcr
                })

    return transactions



# ---------------- TABLE BASED PARSER ------------------

def table_based_parser(pdf_path):

    transactions = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            tables = page.extract_tables()

            if not tables:
                continue

            for table in tables:

                for row in table:

                    if not row or len(row) < 4:
                        continue

                    date = row[0]
                    description = row[1]
                    amount = row[-1]

                    if not date or not amount:
                        continue

                    if "DATE" in str(date).upper():
                        continue

                    try:
                        amt = float(
                            amount.replace("Dr", "")
                                  .replace("Cr", "")
                                  .replace(",", "")
                                  .strip()
                        )

                        if "Cr" in amount:
                            amt = -amt

                        transactions.append({
                            "Account": "Axis Bank Rewards CC",
                            "Date": date.strip(),
                            "Description": description.strip(),
                            "Amount": amt,
                            "Type": "Cr" if amt < 0 else "Dr"
                        })

                    except:
                        continue

    return transactions



# ---------------- OCR BASED PARSER ------------------

def ocr_based_parser(pdf_path):

    transactions = []

    images = convert_from_path(pdf_path)

    for img in images:

        text = pytesseract.image_to_string(img)

        pattern = r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,]+\.\d{2})\s+(Dr|Cr)"

        matches = re.findall(pattern, text)

        for m in matches:

            date, desc, amt, drcr = m

            amount = float(amt.replace(",", ""))

            if drcr == "Cr":
                amount = -amount

            transactions.append({
                "Account": "Axis Bank Rewards CC",
                "Date": date,
                "Description": desc.strip(),
                "Amount": amount,
                "Type": drcr
            })

    return transactions



# ---------------- SMART MASTER PARSER ------------------

def parse_axis_rewards_smart(pdf_path):

    print(f"\nTrying TEXT parser for: {pdf_path}")
    tx = text_based_parser(pdf_path)

    if tx:
        print("✔ Text parser succeeded")
        return tx

    print("⚠ Text parser failed – trying TABLE parser")
    tx = table_based_parser(pdf_path)

    if tx:
        print("✔ Table parser succeeded")
        return tx

    print("⚠ Table parser failed – trying OCR parser")
    tx = ocr_based_parser(pdf_path)

    if tx:
        print("✔ OCR parser succeeded")
        return tx

    print("❌ All parsers failed – No transactions detected")
    return []
