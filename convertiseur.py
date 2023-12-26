import fitz  # PyMuPDF
import re
pdf_path = "PDF2.pdf"
doc = fitz.open(pdf_path)
page_number = 0 
page = doc.load_page(page_number)
text = page.get_text("text")
print(text)


def extraction(chemin):
    text = ""
    with fitz.open(chemin) as pdf_document:
        for page_num in range(pdf_document.page_count):
            page = pdf_document[page_num]
            text += page.get_text()

    return text

def expMatches(text, regex_pattern):
    matches = re.findall(regex_pattern, text)
    return matches

# Chemin du fichier PDF d'exemple
chemin = 'PDF2.pdf'

# Extraire le texte du PDF avec PyMuPDF
pdf_text = extraction(chemin)

# Définir le motif d'expression régulière que vous souhaitez rechercher
regex_pattern = r'\b(VEMS.*)\b' # Exemple: correspondance à un code postal de 5 chiffres

# Trouver les correspondances d'expression régulière dans le texte extrait
regex_matches = expMatches(pdf_text, regex_pattern)

# Afficher les correspondances trouvées
print("Correspondances d'expression régulière trouvées:", regex_matches)
