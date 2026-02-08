python3 - << 'EOF'
import pdfplumber

pdf_path = "../CC statements/Axis Rewards/July.pdf"

with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages):
        print(f"\n--- PAGE {i+1} TEXT ---\n")
        print(page.extract_text())
EOF
