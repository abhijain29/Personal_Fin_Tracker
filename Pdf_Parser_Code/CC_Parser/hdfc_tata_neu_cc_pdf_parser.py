import re
from datetime import datetime

import pdfplumber


def _parse_amount(s: str):
    if not s:
        return None
    s = str(s)
    s = s.replace(",", "").strip()
    # Sometimes OCR prefixes currency with 'C' (for Rs symbol) or similar noise.
    s = re.sub(r"^[^0-9\\-]+", "", s)
    try:
        return float(s)
    except Exception:
        return None


def _period_from_statement_date(text: str) -> str:
    # Example: "Statement Date 01 May, 2026" -> "May-26"
    m = re.search(r"Statement\s+Date\s+(\d{1,2})\s+([A-Za-z]{3,}),\s*(\d{4})", text, re.I)
    if not m:
        return ""
    day, month, year = m.groups()
    for fmt in ("%d %b %Y", "%d %B %Y"):
        try:
            dt = datetime.strptime(f"{day} {month} {year}", fmt)
            return dt.strftime("%b-%y")
        except Exception:
            pass
    return ""


def parse_hdfc_tata_neu_cc_pdf(pdf_path: str):
    """
    Parse HDFC Tata Neu credit card statements.

    Observed transaction format:
      04/04/2026| 19:48 UPI-ixigo C 1,557.00 l
    """
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        first_text = (pdf.pages[0].extract_text() or "") if pdf.pages else ""
        period = _period_from_statement_date(first_text) or "Unknown"

        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
            for ln in lines:
                # Transaction line begins with date + pipe.
                if not re.match(r"^\d{2}/\d{2}/\d{4}\|", ln):
                    continue

                # Strip the date/time prefix.
                m = re.match(r"^(\d{2}/\d{2}/\d{4})\|\s*(\d{2}:\d{2})\s+(.*)$", ln)
                if not m:
                    continue
                date_str, time_str, tail = m.groups()

                # Tail pattern: "<desc> <noise?> <amount> <Dr/Cr-ish?>"
                # Common: "... C 1,557.00 l" where C/l are OCR noise.
                amt_m = re.search(r"([0-9][0-9,]*\.[0-9]{2})\b", tail)
                if not amt_m:
                    continue
                amt_raw = amt_m.group(1)
                amt = _parse_amount(amt_raw)
                if amt is None:
                    continue

                desc = tail[: amt_m.start()].strip()
                # Remove trailing OCR junk tokens like solitary "C" or "|" before amount.
                desc = re.sub(r"[\|\s]+$", "", desc).strip()
                desc = re.sub(r"\s+[A-Za-z]{1}\s*$", "", desc).strip()

                # This section in the statement is purchases/debits (money out).
                records.append(
                    {
                        "Period": period,
                        "Account": "HDFC Tata Neu",
                        "Date": date_str,
                        "Description": desc,
                        "Amount": float(amt),
                        "Type": "Dr",
                    }
                )

    return records
