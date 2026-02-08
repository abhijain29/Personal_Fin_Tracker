import pdfplumber
import re
from datetime import datetime

def clean_description(text):
    text = text.replace("|", " ")
    text = text.replace("_", " ")
    text = text.replace("*", " ")
    # Remove merchant category codes
    text = re.sub(r'\b(TRANSPORT|HOTELS|MERCHANT CATEGORY)\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def extract_period(text):
    text_norm = re.sub(r"\s+", " ", text)
    # Pattern 1: "Statement Generation Date 18/12/2025"
    match = re.search(r"Statement Generation Date\s*[:\\-]?\\s*(\\d{2}/\\d{2}/\\d{4})", text_norm, re.IGNORECASE)
    if match:
        dt = datetime.strptime(match.group(1), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    
    # Pattern 2: "Statement Date 18/11/2025"
    match = re.search(r"Statement Date\s*[:\\-]?\\s*(\\d{2}/\\d{2}/\\d{4})", text_norm, re.IGNORECASE)
    if match:
        dt = datetime.strptime(match.group(1), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    
    # Pattern 3: Statement Period "20/10/2025 - 18/11/2025"
    match = re.search(r"Statement Period\s+(\d{2}/\d{2}/\d{4})\s+-\s+(\d{2}/\d{2}/\d{4})", text_norm)
    if match:
        # Use end date
        dt = datetime.strptime(match.group(2), "%d/%m/%Y")
        return dt.strftime("%b-%y")

    # Pattern 4: Any date range in the document (fallback to end date)
    match = re.search(r"(\d{2}/\d{2}/\d{4})\s*[-–]\s*(\d{2}/\d{2}/\d{4})", text_norm)
    if match:
        dt = datetime.strptime(match.group(2), "%d/%m/%Y")
        return dt.strftime("%b-%y")
    
    # Pattern 5: Look in first 1000 chars for any date
    match = re.search(r"(\d{2}/\d{2}/\d{4})", text_norm[:1000])
    if match:
        try:
            dt = datetime.strptime(match.group(1), "%d/%m/%Y")
            return dt.strftime("%b-%y")
        except:
            pass
    
    return ""

def parse_axis_pdf(pdf_path):
    """
    Unified parser for all Axis Bank credit cards (Indian Oil, Select, Rewards)
    Handles both text-based and OCR-based PDFs
    """
    transactions = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"

            # If full_text is empty or very short, PDF might be image-based (OCR needed)
            if len(full_text.strip()) < 100:
                print("   ⚠️ PDF appears to be image-based (OCR required)")
                return []

            period = extract_period(full_text)

            # Detect card type from path or PDF content
            path_lower = pdf_path.lower()
            if "select" in path_lower or "SELECT" in full_text:
                account = "Axis Bank Select CC"
            elif "indian oil" in path_lower:
                account = "Axis Bank Indian Oil CC"
            elif "rewards" in path_lower or "REWARDS" in full_text:
                account = "Axis Bank Rewards CC"
            else:
                account = "Axis Bank CC"

            # Pattern for Axis transactions
            # Format: DD/MM/YYYY DESCRIPTION CATEGORY AMOUNT Dr/Cr
            # Example: 08/12/2025 RAZ*IXIGO,GURGAON TRANSPORT 741.00 Dr
            # Example: 08/12/2025 PAYMENT RECEIVED 2,538.00 Cr
            
            pattern = r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,]+\.\d{2})\s+(Dr|Cr)"

            for match in re.findall(pattern, full_text):
                date, desc, amount, txn_type = match

                amount = float(amount.replace(",", ""))

                # Cr means payment/refund - make it negative
                if txn_type == "Cr":
                    amount = -amount

                transactions.append({
                    "Period": period,
                    "Account": account,
                    "Date": date,
                    "Description": clean_description(desc),
                    "Amount": amount,
                    "Type": txn_type
                })

    except Exception as e:
        print(f"   ❌ Error parsing Axis PDF: {e}")
        return []

    return transactions
