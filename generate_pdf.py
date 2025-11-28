import fitz  # PyMuPDF library

filename = "Morgan_Stanley_Invoice.pdf"
doc = fitz.open() # Create empty PDF
page = doc.new_page()

print("Generating PDF with default fonts...")

# 1. Add Header (Blue Text)
# Removed 'fontname' argument to avoid "need font file" error
page.insert_text((50, 50), "INVOICE #2025-MS", fontsize=20, color=(0, 0, 1))

# 2. Add Body Text
page.insert_text((50, 100), "Bill To:", fontsize=12)
page.insert_text((50, 120), "Morgan Stanley Financials", fontsize=12)
page.insert_text((50, 140), "Cybersecurity Audit Dept", fontsize=12)

# 3. Add a Hidden Note at the bottom
page.insert_text((50, 500), "Note: Internal audit use only. Verified by Morgan.", fontsize=10, color=(0.5, 0.5, 0.5))

doc.save(filename)
print(f"✅ PDF Created Successfully: {filename}")