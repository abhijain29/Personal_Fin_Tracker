import pdfplumber
import re


def parse_uni_gold_cc_pdf(pdf_path):

    transactions = []

    try:
        with pdfplumber.open(pdf_path) as pdf:

            full_text = ""

            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += "\n" + text

        # Pattern for lines like:
        # 01/11/2025 CCCPL FRONT OFFICE II HYDERABAD IN DEBIT ₹1,79,520

        pattern = re.compile(
            r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(DEBIT|CREDIT)\s+₹([\d,]+\.\d{0,2}|[\d,]+)"
        )

        matches = pattern.findall(full_text)

        for match in matches:

            date = match[0]
            description = match[1].strip()
            txn_type = match[2]
            amount = match[3]

            amount_clean = amount.replace(",", "")

            try:
                amount_val = float(amount_clean)
            except:
                continue

            transactions.append({
                "Account": "Uni Gold Card",
                "Date": date,
                "Description": description,
                "Amount": amount_val,
                "Type": "Dr" if txn_type == "DEBIT" else "Cr"
            })

    except Exception as e:
        print("❌ Error parsing Uni Gold PDF:", e)

    return transactions


# ----------------- TEST BLOCK ------------------
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python uni_gold_cc_pdf_parser.py <pdf_path>")
        sys.exit(1)

    pdf_file = sys.argv[1]

    txns = parse_uni_gold_cc_pdf(pdf_file)

    print("\nTransactions Extracted:", len(txns))

    for t in txns:
        print(t)
