import fitz
import camelot
"""pdf_path="PDF2.pdf"
document=fitz.open(pdf_path)
page_number=0
page=document.load_page(page_number)
text=page.get_text("text")
print(text)
document.close()
"""
pdf_path = "PDF2.pdf"

# Utilisez la fonction read_pdf pour extrsaire les tableaux
tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')

# Affichez les tables extraites
for table in tables:
    print(table.df)
