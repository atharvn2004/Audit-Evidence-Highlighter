import os
import fitz  # PyMuPDF
import cv2   # OpenCV
import os
import fitz  # PyMuPDF for PDF
import cv2   # OpenCV for Images
import pytesseract  # <--- MAKE SURE THIS IS IMPORTED

# === ADD THIS LINE ===
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# =====================

import openpyxl
# ... rest of the code ...
import pytesseract
from PIL import Image
import openpyxl
from openpyxl.styles import Border, Side
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==========================================
#           TESSERACT SETUP (CRITICAL)
# ==========================================
# If you are on WINDOWS, remove the '#' from the line below:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# If you installed Tesseract somewhere else, update the path above.
# On MAC/LINUX, usually no change is needed.
# ==========================================

class AuditEvidenceHighlighter:
    def __init__(self, search_text):
        self.search_text = search_text
        self.output_dir = "audit_output"
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def process_file(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        filename = os.path.basename(file_path)
        output_path = os.path.join(self.output_dir, f"highlighted_{filename}")
        print(f"Processing {filename}...")

        try:
            if ext == '.pdf':
                self._process_pdf(file_path, output_path)
            elif ext in ['.png', '.jpg', '.jpeg']:
                self._process_image(file_path, output_path)
            elif ext == '.xlsx':
                self._process_excel(file_path, output_path)
            elif ext == '.docx':
                self._process_word(file_path, output_path)
            else:
                print(f"Unsupported file format: {ext}")
                return
            print(f"Success! Output saved to: {output_path}")
        except Exception as e:
            print(f"Error processing {filename}: {e}")

    def _process_pdf(self, input_path, output_path):
        doc = fitz.open(input_path)
        found = False
        for page in doc:
            text_instances = page.search_for(self.search_text)
            for inst in text_instances:
                found = True
                page.draw_rect(inst, color=(1, 0, 0), width=1.5, fill_opacity=0)
        if found:
            doc.save(output_path)
        else:
            print("Text not found in PDF.")

    def _process_image(self, input_path, output_path):
        img = cv2.imread(input_path)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT)
        found = False
        n_boxes = len(data['text'])
        for i in range(n_boxes):
            if self.search_text.lower() in data['text'][i].lower():
                found = True
                (x, y, w, h) = (data['left'][i], data['top'][i], data['width'][i], data['height'][i])
                cv2.rectangle(img, (x, y), (x + w, y + h), (0, 0, 255), 2)
        if found:
            cv2.imwrite(output_path, img)
        else:
            print("Text not found in Image.")

    def _process_excel(self, input_path, output_path):
        wb = openpyxl.load_workbook(input_path)
        found = False
        red_border = Border(left=Side(style='thick', color='FF0000'), 
                            right=Side(style='thick', color='FF0000'), 
                            top=Side(style='thick', color='FF0000'), 
                            bottom=Side(style='thick', color='FF0000'))
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and self.search_text.lower() in str(cell.value).lower():
                        found = True
                        cell.border = red_border
        if found:
            wb.save(output_path)
        else:
            print("Text not found in Excel.")

    def _process_word(self, input_path, output_path):
        doc = Document(input_path)
        found = False
        def add_border(run):
            rPr = run._element.get_or_add_rPr()
            bdr = OxmlElement('w:bdr')
            bdr.set(qn('w:val'), 'single')
            bdr.set(qn('w:sz'), '4')
            bdr.set(qn('w:color'), 'FF0000')
            rPr.append(bdr)

        for paragraph in doc.paragraphs:
            if self.search_text.lower() in paragraph.text.lower():
                found = True
                for run in paragraph.runs:
                    if self.search_text.lower() in run.text.lower():
                        add_border(run)
        if found:
            doc.save(output_path)
        else:
            print("Text not found in Word Doc.")

if __name__ == "__main__":
    print("--- Audit Evidence Highlighter ---")
    f_path = input("Enter file name (e.g., invoice.pdf): ").strip().strip('"')
    query = input("Enter text to search: ").strip()
    
    if os.path.exists(f_path):
        highlighter = AuditEvidenceHighlighter(query)
        highlighter.process_file(f_path)
    else:
        print("Error: File does not exist in this folder.")