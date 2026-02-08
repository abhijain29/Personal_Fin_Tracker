#!/usr/bin/env python3
"""
Debug script to find missing ICICI transactions
"""

import pdfplumber
import re
from pathlib import Path

PDF_PATH = "/Users/abhishekjain/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker/CC statements/ICICI Amazon/Aug.pdf"

print("="*70)
print("ICICI DEBUG SCRIPT")
print("="*70)
print(f"Analyzing: {PDF_PATH}\n")

with pdfplumber.open(PDF_PATH) as pdf:
    for page_num, page in enumerate(pdf.pages, 1):
        text = page.extract_text()
        
        if not text:
            continue
        
        lines = text.split('\n')
        
        print(f"\n{'='*70}")
        print(f"PAGE {page_num} - All lines with dates:")
        print('='*70)
        
        for i, line in enumerate(lines, 1):
            # Look for lines starting with dates
            if re.match(r'^\d{2}/\d{2}/\d{4}', line):
                print(f"\nLine {i}: {line}")
                
                # Check for CR
                has_cr = ' CR' in line
                print(f"  Has CR: {has_cr}")
                
                # Find amounts
                amounts = re.findall(r'([\d,]+\.\d{2})', line)
                print(f"  Amounts found: {amounts}")
                
                # Check for specific missing amounts
                if '29,900.00' in line or '1,990.00' in line:
                    print(f"  ⭐ FOUND MISSING TRANSACTION!")
                    print(f"  Full line: {line}")

print("\n" + "="*70)
print("Looking specifically for missing amounts:")
print("="*70)

with pdfplumber.open(PDF_PATH) as pdf:
    full_text = pdf.pages[0].extract_text()
    
    # Look for 29,900.00
    if '29,900.00' in full_text:
        print("\n✅ Found 29,900.00 in PDF")
        # Extract context
        idx = full_text.find('29,900.00')
        context = full_text[max(0, idx-100):idx+100]
        print(f"Context:\n{context}")
    else:
        print("\n❌ 29,900.00 NOT found in PDF text")
    
    # Look for 1,990.00
    if '1,990.00' in full_text:
        print("\n✅ Found 1,990.00 in PDF")
        idx = full_text.find('1,990.00')
        context = full_text[max(0, idx-100):idx+100]
        print(f"Context:\n{context}")
    else:
        print("\n❌ 1,990.00 NOT found in PDF text")

print("\n" + "="*70)
print("Now testing with current parser...")
print("="*70)

from icici_cc_pdf_parser import extract_icici_transactions

transactions = extract_icici_transactions(PDF_PATH)

if transactions:
    print(f"\n✅ Parser extracted {len(transactions)} transactions\n")
    for tx in transactions:
        print(f"{tx['Date']} | {tx['Description'][:40]:40} | ₹{tx['Amount']:>10,.2f}")
    
    # Check if missing amounts are in results
    amounts = [tx['Amount'] for tx in transactions]
    if 29900.00 in amounts:
        print("\n✅ 29,900.00 IS in results")
    else:
        print("\n❌ 29,900.00 MISSING from results")
    
    if 1990.00 in amounts:
        print("✅ 1,990.00 IS in results")
    else:
        print("❌ 1,990.00 MISSING from results")
else:
    print("\n❌ Parser returned no transactions")
