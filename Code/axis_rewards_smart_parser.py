import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import re
from datetime import datetime


def clean_description(text):
    """Clean junk from description"""
    text = str(text).replace("|", " ")
    text = text.replace("_", " ")
    text = text.replace("*", " ")
    text = re.sub(r'\b(TRANSPORT|HOTELS|MERCHANT CATEGORY)\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def extract_period(text):
    """Extract period from statement text"""
    # Normalize whitespace for easier regex matching
    text_norm = re.sub(r"\s+", " ", text)

    # Pattern 1: "Statement Generation Date 18/12/2025"
    match = re.search(r"Statement Generation Date\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", text_norm, re.IGNORECASE)
    if match:
        dt = datetime.strptime(match.group(1), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    
    # Pattern 2: "Statement Date 18/11/2025"
    match = re.search(r"Statement Date\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", text_norm, re.IGNORECASE)
    if match:
        dt = datetime.strptime(match.group(1), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    
    # Pattern 3: Statement Period "20/10/2025 - 18/11/2025"
    match = re.search(r"Statement Period\s+(\d{2}/\d{2}/\d{4})\s+-\s+(\d{2}/\d{2}/\d{4})", text_norm, re.IGNORECASE)
    if match:
        dt = datetime.strptime(match.group(2), "%d/%m/%Y")
        return dt.strftime("%b-%y")

    # Pattern 4: Any date range in the document (fallback to end date)
    match = re.search(r"(\d{2}/\d{2}/\d{4})\s*[-–]\s*(\d{2}/\d{2}/\d{4})", text_norm)
    if match:
        dt = datetime.strptime(match.group(2), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    
    # Pattern 5: Find first date in document
    match = re.search(r"(\d{2}/\d{2}/\d{4})", text_norm[:2000])
    if match:
        try:
            dt = datetime.strptime(match.group(1), "%d/%m/%Y")
            return dt.strftime("%b-%y")
        except:
            pass
    
    return ""


# ---------------- TEXT BASED PARSER ------------------

def text_based_parser(pdf_path):
    """
    Parse text-based Axis Rewards PDF
    """
    transactions = []
    period = ""

    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"

            # Extract period once from full document
            if not period:
                period = extract_period(full_text)

            # If no text extracted, this is likely image-based
            if len(full_text.strip()) < 100:
                return []

            # Pattern for Axis transactions
            pattern = r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,]+\.\d{2})\s+(Dr|Cr)"
            matches = re.findall(pattern, full_text)

            for m in matches:
                date, desc, amt, drcr = m

                amount = float(amt.replace(",", ""))

                if drcr == "Cr":
                    amount = -amount

                transactions.append({
                    "Period": period,
                    "Account": "Axis Bank Rewards CC",
                    "Date": date,
                    "Description": clean_description(desc),
                    "Amount": amount,
                    "Type": drcr
                })

    except Exception as e:
        print(f"      Text parser error: {e}")
        return []

    return transactions


# ---------------- TABLE BASED PARSER ------------------

def table_based_parser(pdf_path):
    """
    Parse table-based Axis Rewards PDF
    """
    transactions = []
    period = ""

    try:
        with pdfplumber.open(pdf_path) as pdf:
            # First, extract text for period
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
            
            period = extract_period(full_text)

            # Now extract tables
            for page in pdf.pages:
                tables = page.extract_tables()

                if not tables:
                    continue

                for table in tables:
                    for row in table:
                        if not row or len(row) < 4:
                            continue

                        # Typically: Date | Description | Category | Amount Dr/Cr
                        date = str(row[0]).strip()
                        description = str(row[1]).strip() if len(row) > 1 else ""
                        amount_cell = str(row[-1]).strip()

                        # Skip header rows
                        if not date or not amount_cell:
                            continue
                        if "DATE" in date.upper() or "TRANSACTION" in date.upper():
                            continue

                        # Validate date format
                        if not re.match(r'\d{2}/\d{2}/\d{4}', date):
                            continue

                        try:
                            # Extract amount and Dr/Cr
                            amt_clean = amount_cell.replace("Dr", "").replace("Cr", "").replace(",", "").strip()
                            amt = float(amt_clean)

                            if "Cr" in amount_cell:
                                amt = -amt

                            transactions.append({
                                "Period": period,
                                "Account": "Axis Bank Rewards CC",
                                "Date": date,
                                "Description": clean_description(description),
                                "Amount": amt,
                                "Type": "Cr" if amt < 0 else "Dr"
                            })

                        except ValueError:
                            continue

    except Exception as e:
        print(f"      Table parser error: {e}")
        return []

    return transactions


# ---------------- OCR BASED PARSER ------------------

def ocr_based_parser(pdf_path):
    """
    Parse image-based (scanned) Axis Rewards PDF using OCR
    Requires: pip install pytesseract pdf2image
    Mac: brew install tesseract
    """
    transactions = []
    period = ""

    try:
        # Convert PDF pages to images
        images = convert_from_path(pdf_path)
        full_text = ""

        for img in images:
            # Extract text using OCR
            text = pytesseract.image_to_string(img, lang='eng')
            full_text += text + "\n"

        # Extract period from OCR'd text
        period = extract_period(full_text)

        # Find transactions using same pattern as text parser
        pattern = r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,]+\.\d{2})\s+(Dr|Cr)"
        matches = re.findall(pattern, full_text)

        for m in matches:
            date, desc, amt, drcr = m

            amount = float(amt.replace(",", ""))

            if drcr == "Cr":
                amount = -amount

            transactions.append({
                "Period": period,
                "Account": "Axis Bank Rewards CC",
                "Date": date,
                "Description": clean_description(desc),
                "Amount": amount,
                "Type": drcr
            })

    except ImportError:
        print("      OCR libraries not installed. Run: pip install pytesseract pdf2image")
        return []
    except Exception as e:
        print(f"      OCR parser error: {e}")
        return []

    return transactions


# ---------------- SMART MASTER PARSER ------------------

def parse_axis_rewards_smart(pdf_path):
    """
    Smart parser that tries multiple strategies in order:
    1. Text-based extraction (fastest, most reliable)
    2. Table-based extraction (for structured PDFs)
    3. OCR-based extraction (for scanned/image PDFs)
    
    Returns first successful result or empty list if all fail.
    """
    
    print(f"   📄 Trying TEXT parser...")
    tx = text_based_parser(pdf_path)

    if tx:
        print(f"   ✅ Text parser succeeded ({len(tx)} transactions)")
        return tx

    print(f"   ⚠️ Text parser failed → trying TABLE parser...")
    tx = table_based_parser(pdf_path)

    if tx:
        print(f"   ✅ Table parser succeeded ({len(tx)} transactions)")
        return tx

    print(f"   ⚠️ Table parser failed → trying OCR parser...")
    tx = ocr_based_parser(pdf_path)

    if tx:
        print(f"   ✅ OCR parser succeeded ({len(tx)} transactions)")
        return tx

    print(f"   ❌ All parsers failed - No transactions detected")
    return []


# ---------------- TEST FUNCTION ------------------

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python axis_rewards_smart_parser.py <pdf_path>")
        sys.exit(1)
    
    pdf_file = sys.argv[1]
    
    print(f"\nTesting Axis Rewards Smart Parser on: {pdf_file}")
    print("="*70)
    
    transactions = parse_axis_rewards_smart(pdf_file)
    
    if transactions:
        print(f"\n✅ Successfully extracted {len(transactions)} transactions")
        print("\nSample transactions:")
        for tx in transactions[:5]:
            print(f"  {tx['Date']} | {tx['Description'][:40]:40} | ₹{tx['Amount']:>10,.2f} | {tx['Type']}")
        
        if len(transactions) > 5:
            print(f"  ... and {len(transactions) - 5} more")
    else:
        print("\n❌ No transactions extracted")
