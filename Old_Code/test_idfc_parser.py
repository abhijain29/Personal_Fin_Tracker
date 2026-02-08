#!/usr/bin/env python3
"""
Test script for IDFC FIRST Bank Credit Card PDF parser
Usage: python3 test_idfc_parser.py <filename>
Example: python3 test_idfc_parser.py Nov.pdf
"""

import sys
from pathlib import Path

# ---------------------------------------------------------
# PATH SETUP
# ---------------------------------------------------------
BASE_DIR = (
    Path.home()
    / "Library"
    / "CloudStorage"
    / "OneDrive-Personal"
    / "Personal"
    / "Finance"
    / "projects"
    / "Claude"
)

CODE_DIR = BASE_DIR / "Code"
CC_STATEMENTS_DIR = BASE_DIR / "CC statements"  # Note: with space
IDFC_DIR = CC_STATEMENTS_DIR / "IDFC"

# Add code directory to path so we can import the parser
sys.path.insert(0, str(CODE_DIR))


def test_idfc_parser():
    # Import the parser
    try:
        from idfc_cc_pdf_parser import parse_idfc_cc_pdf
    except ImportError:
        print("❌ Could not import idfc_cc_pdf_parser.py")
        print(f"Make sure the file is in: {CODE_DIR}")
        sys.exit(1)
    
    # Get PDF filename from command line
    if len(sys.argv) < 2:
        print("Usage: python3 test_idfc_parser.py <filename>")
        print("\nExample:")
        print("  python3 test_idfc_parser.py Nov.pdf")
        print("  python3 test_idfc_parser.py Oct.pdf")
        print(f"\nWill look for PDF in: {IDFC_DIR}")
        sys.exit(1)
    
    # Construct full path to PDF
    filename = sys.argv[1]
    pdf_path = IDFC_DIR / filename
    
    if not pdf_path.exists():
        print(f"❌ File not found: {pdf_path}")
        print(f"\nChecking what's in the folder:")
        if IDFC_DIR.exists():
            files = list(IDFC_DIR.glob("*.pdf"))
            if files:
                print("Available PDFs:")
                for f in files:
                    print(f"  - {f.name}")
            else:
                print("  No PDF files found")
        else:
            print(f"  Folder doesn't exist: {IDFC_DIR}")
        sys.exit(1)
    
    print("="*70)
    print(f"Testing IDFC FIRST Bank PDF Parser")
    print("="*70)
    print(f"File: {pdf_path.name}")
    print(f"Path: {pdf_path}")
    print(f"Size: {pdf_path.stat().st_size / 1024:.1f} KB")
    print("="*70)
    print()
    
    # Parse the PDF
    print("Parsing PDF...")
    df = parse_idfc_cc_pdf(pdf_path)
    
    if df is None or df.empty:
        print("❌ Parsing failed or no transactions found")
        print("\nDebugging tips:")
        print("1. Open the PDF and check transaction format")
        print("2. Verify transactions section starts with 'YOUR TRANSACTIONS'")
        print("3. Check if dates are in format: DD Mon YY (e.g., 08 Nov 25)")
        print("4. Check if amounts end with 'DR' or 'CR'")
        sys.exit(1)
    
    # Display results
    print(f"✅ Successfully parsed!\n")
    
    print("="*70)
    print("SUMMARY")
    print("="*70)
    print(f"Total transactions:     {len(df)}")
    print(f"Date range:             {df['date'].min()} to {df['date'].max()}")
    print(f"Period:                 {df['period'].iloc[0]}")
    print(f"Account:                {df['account'].iloc[0]}")
    print()
    
    print("="*70)
    print("FINANCIAL SUMMARY")
    print("="*70)
    debits = df[df['amount'] > 0]
    credits = df[df['amount'] < 0]
    
    print(f"Total Debits:           ₹{debits['amount'].sum():>12,.2f} ({len(debits)} txns)")
    print(f"Total Credits/Refunds:  ₹{abs(credits['amount'].sum()):>12,.2f} ({len(credits)} txns)")
    print(f"Net Amount:             ₹{df['amount'].sum():>12,.2f}")
    print()
    
    print("="*70)
    print("SAMPLE TRANSACTIONS (First 10)")
    print("="*70)
    print(df.head(10).to_string(index=False))
    print()
    
    # Save to file in Output directory
    output_dir = BASE_DIR / "Output"
    output_dir.mkdir(exist_ok=True)
    output_file = output_dir / "test_idfc_output.csv"
    df.to_csv(output_file, index=False)
    print(f"💾 Full data saved to: {output_file}")
    
    print()
    print("="*70)
    print("✅ TEST COMPLETE!")
    print("="*70)
    print("\nIf everything looks good, you can now:")
    print("1. Copy more PDFs to CC statements/IDFC folder")
    print("2. Update pdf_parser_main.py to include IDFC parser")


if __name__ == "__main__":
    try:
        test_idfc_parser()
    except KeyboardInterrupt:
        print("\n\n⚠️ Test interrupted")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
