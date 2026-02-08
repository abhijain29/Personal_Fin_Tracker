from pathlib import Path
import pandas as pd
import sys

# ============================================================================
# CONFIGURATION - Updated for your folder structure
# ============================================================================

BASE_DIR = Path("/Users/abhishekjain/Library/CloudStorage/OneDrive-Personal/Personal/Finance/projects/Monthly_Fin_Tracker")

# Folder paths matching your structure
CC_STATEMENTS_DIR = BASE_DIR / "CC_statements"
SAVING_STATEMENTS_DIR = BASE_DIR / "Saving_statements"
CODE_DIR = BASE_DIR / "Code"
OUTPUT_DIR = BASE_DIR / "Output"
LOGS_DIR = BASE_DIR / "Logs"
ARCHIVE_DIR = BASE_DIR / "Archive"
ERROR_DIR = BASE_DIR / "Error"

# Create directories if they don't exist
for directory in [OUTPUT_DIR, LOGS_DIR, ARCHIVE_DIR, ERROR_DIR]:
    directory.mkdir(parents=True, exist_ok=True)

# ============================================================================
# IMPORT PARSERS
# ============================================================================

# Add code directory to path
sys.path.insert(0, str(CODE_DIR))

# Import bank-specific PDF parsers
try:
    from icici_cc_pdf_parser import parse_icici_cc_pdf
except ImportError as e:
    print(f"⚠️ Warning: Could not import icici_cc_pdf_parser: {e}")
    parse_icici_cc_pdf = None

# Add more parsers as you implement them
# from axis_cc_pdf_parser import parse_axis_cc_pdf
# from hsbc_savings_pdf_parser import parse_hsbc_savings_pdf

# ============================================================================
# PARSER REGISTRY
# ============================================================================

CC_PARSERS = {}
SAVINGS_PARSERS = {}

if parse_icici_cc_pdf:
    CC_PARSERS["ICICI Amazon"] = parse_icici_cc_pdf
    CC_PARSERS["Axis Select"] = None  # Add parser when ready

# SAVINGS_PARSERS["ICICI"] = parse_icici_savings_pdf
# SAVINGS_PARSERS["HSBC"] = parse_hsbc_savings_pdf
# SAVINGS_PARSERS["IDFC"] = parse_idfc_savings_pdf

# ============================================================================
# LOGGING SETUP
# ============================================================================

import logging
from datetime import datetime

log_file = LOGS_DIR / f"parser_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def move_to_archive(pdf_file):
    """Move successfully parsed PDF to Archive folder"""
    try:
        archive_path = ARCHIVE_DIR / pdf_file.name
        # If file exists in archive, add timestamp
        if archive_path.exists():
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            archive_path = ARCHIVE_DIR / f"{pdf_file.stem}_{timestamp}{pdf_file.suffix}"
        
        import shutil
        shutil.move(str(pdf_file), str(archive_path))
        logger.info(f"Moved to archive: {pdf_file.name}")
        return True
    except Exception as e:
        logger.error(f"Failed to archive {pdf_file.name}: {e}")
        return False


def move_to_error(pdf_file, error_msg):
    """Move failed PDF to Error folder"""
    try:
        error_path = ERROR_DIR / pdf_file.name
        if error_path.exists():
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            error_path = ERROR_DIR / f"{pdf_file.stem}_{timestamp}{pdf_file.suffix}"
        
        import shutil
        shutil.move(str(pdf_file), str(error_path))
        
        # Create error log file
        error_log = ERROR_DIR / f"{pdf_file.stem}_error.txt"
        with open(error_log, 'w') as f:
            f.write(f"Error parsing {pdf_file.name}\n")
            f.write(f"Time: {datetime.now()}\n")
            f.write(f"Error: {error_msg}\n")
        
        logger.warning(f"Moved to error folder: {pdf_file.name}")
        return True
    except Exception as e:
        logger.error(f"Failed to move to error folder {pdf_file.name}: {e}")
        return False

# ============================================================================
# MAIN PROCESSING FUNCTIONS
# ============================================================================

def process_cc_statements():
    """Process Credit Card statements"""
    
    logger.info("="*70)
    logger.info("Processing Credit Card Statements")
    logger.info("="*70)
    
    cc_data = []
    stats = {'total': 0, 'success': 0, 'failed': 0, 'transactions': 0}
    
    # Process each bank folder in CC_statements
    for bank_dir in CC_STATEMENTS_DIR.iterdir():
        if not bank_dir.is_dir():
            continue
        
        bank_name = bank_dir.name
        
        # Check if parser exists
        if bank_name not in CC_PARSERS or CC_PARSERS[bank_name] is None:
            logger.warning(f"No parser for {bank_name} (skipping)")
            continue
        
        # Find all PDFs
        pdf_files = list(bank_dir.glob("*.pdf")) + list(bank_dir.glob("*.PDF"))
        
        if not pdf_files:
            logger.info(f"No PDFs in {bank_name}")
            continue
        
        logger.info(f"\n📁 {bank_name}: {len(pdf_files)} PDF(s)")
        
        # Process each PDF
        for pdf_file in pdf_files:
            stats['total'] += 1
            logger.info(f"  📄 {pdf_file.name:40} ... ")
            
            try:
                df = CC_PARSERS[bank_name](pdf_file)
                
                if df is not None and not df.empty:
                    cc_data.append(df)
                    stats['success'] += 1
                    stats['transactions'] += len(df)
                    logger.info(f"    ✅ {len(df)} transactions")
                    
                    # Move to archive
                    move_to_archive(pdf_file)
                else:
                    stats['failed'] += 1
                    logger.warning(f"    ⚠️ No data extracted")
                    move_to_error(pdf_file, "No transactions found")
            
            except Exception as e:
                stats['failed'] += 1
                logger.error(f"    ❌ Error: {e}")
                move_to_error(pdf_file, str(e))
    
    return cc_data, stats


def process_savings_statements():
    """Process Savings Account statements (placeholder for future)"""
    
    logger.info("\n" + "="*70)
    logger.info("Processing Savings Statements")
    logger.info("="*70)
    logger.info("⚠️ Savings parsers not implemented yet")
    
    return [], {'total': 0, 'success': 0, 'failed': 0, 'transactions': 0}


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main processing function"""
    
    logger.info(f"PDF Parser Started")
    logger.info(f"Base Directory: {BASE_DIR}")
    logger.info(f"Log File: {log_file}")
    
    # Process Credit Cards
    cc_data, cc_stats = process_cc_statements()
    
    # Process Savings (future)
    # savings_data, savings_stats = process_savings_statements()
    
    # ========================================================================
    # CONSOLIDATE AND SAVE
    # ========================================================================
    
    if not cc_data:
        logger.warning("No Credit Card data parsed")
    else:
        # Combine all CC data
        final_df = pd.concat(cc_data, ignore_index=True)
        
        # Sort by date
        final_df['date_parsed'] = pd.to_datetime(final_df['date'], format='%d/%m/%Y', errors='coerce')
        final_df = final_df.sort_values('date_parsed').drop(columns=['date_parsed'])
        
        # Save to Output
        output_file = OUTPUT_DIR / "summary.csv"
        final_df.to_csv(output_file, index=False)
        
        logger.info("\n" + "="*70)
        logger.info("✅ PROCESSING COMPLETE")
        logger.info("="*70)
        logger.info(f"Files processed:        {cc_stats['total']}")
        logger.info(f"  ├─ Successful:        {cc_stats['success']}")
        logger.info(f"  └─ Failed:            {cc_stats['failed']}")
        logger.info(f"Total transactions:     {cc_stats['transactions']}")
        logger.info(f"Total amount:           ₹{final_df['amount'].sum():,.2f}")
        logger.info(f"Output file:            {output_file}")
        logger.info(f"Log file:               {log_file}")
        logger.info("="*70)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.warning("\n⚠️ Process interrupted by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
