import fitz
pdf_path = "PDF2.pdf"
doc = fitz.open(pdf_path)
page_number = 0  # Numéro de la page (commence à 0)
page = doc.load_page(page_number)
text = page.get_text("text")
print(text)
