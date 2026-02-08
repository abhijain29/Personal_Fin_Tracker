import pdfplumber
import re
from datetime import datetime

def clean_description(text):
    # Remove common junk characters
    text = text.replace("|", " ")
    text = text.replace("_", " ")
    text = text.replace("*", " ")
    # Remove very long serial numbers
    text = re.sub(r"\b\d{11,}\b", "", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def extract_period(text):
    # Look for "Statement period : July 13, 2025 to August 12, 2025"
    match = re.search(r"Statement period\s*:\s*\w+\s+\d+,\s+\d{4}\s+to\s+(\w+)\s+\d+,\s+(\d{4})", text)
    if match:
        month = match.group(1)
        year = match.group(2)
        dt = datetime.strptime(f"{month} {year}", "%B %Y")
        return dt.strftime("%b-%y")
    
    # Alternative: "STATEMENT DATE August 12, 2025"
    match = re.search(r"STATEMENT DATE\s+(\w+)\s+\d+,\s+(\d{4})", text, re.IGNORECASE)
    if match:
        month = match.group(1)
        year = match.group(2)
        dt = datetime.strptime(f"{month} {year}", "%B %Y")
        return dt.strftime("%b-%y")
    
    # Try to extract from Statement Date field
    match = re.search(r"Statement.*?Date[:\-]?\s*(\w+)\s+\d+,\s+(\d{4})", text, re.IGNORECASE)
    if match:
        month = match.group(1)
        year = match.group(2)
        try:
            dt = datetime.strptime(f"{month} {year}", "%B %Y")
            return dt.strftime("%b-%y")
        except:
            pass
    
    return ""

def extract_icici_transactions(pdf_path):
    """
    Parse ICICI Amazon Pay Credit Card PDF
    Uses a more comprehensive extraction approach
    """
    transactions = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

        period = extract_period(full_text)

        lines = full_text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Find date anywhere in the line (handles leading noise like "100% ")
            date_match = re.search(r'(\d{2}/\d{2}/\d{4})', line)
            if not date_match:
                continue
            
            date = date_match.group(1)
            
            # Check if line has CR at the end
            is_credit = ' CR' in line
            
            # Find all amounts (format: number with optional comma and decimals)
            # Pattern matches: 84,900.00 or 29,900.00 or 1,990.00 or 549.00
            amounts = re.findall(r'([\d,]+\.\d{2})', line)
            
            if not amounts:
                continue
            
            # Prefer the maximum amount in the line to avoid missing large txns
            try:
                amount = max(float(a.replace(',', '')) for a in amounts)
            except ValueError:
                continue
            
            # Make credits negative
            if is_credit:
                amount = -amount
            
            # Extract description
            # Strategy: Remove date, serial numbers, amounts, reward points, and CR marker
            desc = line
            desc = desc.replace(date, '')  # Remove date wherever it appears
            desc = desc.replace(' CR', '')  # Remove CR marker
            
            # Remove all numbers that look like serial numbers (10+ digits)
            desc = re.sub(r'\b\d{10,}\b', '', desc)
            
            # Remove all amounts from description
            for amt in amounts:
                desc = desc.replace(amt, '')
            
            # Remove standalone numbers that are likely reward points (1-4 digits)
            # But be careful not to remove numbers that are part of descriptions
            desc = re.sub(r'\s+\d{1,4}\s+', ' ', desc)
            desc = re.sub(r'\s+-\d{1,4}\s+', ' ', desc)
            
            # Clean up
            desc = re.sub(r'\s{2,}', ' ', desc)
            desc = desc.strip()
            
            # Skip if no description left
            if not desc or len(desc) < 3:
                continue
            
            # Final cleaning
            description = clean_description(desc)
            
            transactions.append({
                "Period": period,
                "Account": "ICICI Amazon Pay CC",
                "Date": date,
                "Description": description,
                "Amount": amount,
                "Type": "CR" if is_credit else "DR"
            })

    return transactions
