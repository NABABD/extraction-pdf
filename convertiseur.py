import fitz  # PyMuPDF

def extraire_formes_rectangles_pdf(chemin_pdf):
    doc = fitz.open(chemin_pdf)

    for page_num in range(doc.page_count):
        page = doc[page_num]
        formes = page.get_drawings()

        for index, forme in enumerate(formes):
            if isinstance(forme, dict) and forme.get('type') == 'rect':
                details = forme
                points = details.get('points', None)

                print(f"Page {page_num + 1}, Forme {index + 1}, Type: Rectangle, Points: {points}")
    
    doc.close()

# Utilisation de la fonction avec le chemin vers votre fichier PDF
chemin_vers_pdf = 'PDF2.pdf'
extraire_formes_rectangles_pdf(chemin_vers_pdf)
