import fitz  # PyMuPDF

pdf_path = "PDF2.pdf"
doc = fitz.open(pdf_path)
page_number = 0  # NumÃ©ro de la page (commence Ã  0)
page = doc.load_page(page_number)
text = page.get_text("text")
print(text)