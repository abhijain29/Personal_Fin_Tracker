import pdfplumber

def parse_axis_rewards_pdf(pdf_path):

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

                    # Skip header rows
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
