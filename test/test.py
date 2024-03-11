import fitz  # PyMuPDF
import re
import pdfplumber

# extraire toutes les données du fichier et le mettre dans la variable texte


# Function to extract the entire text from a PDF file
def extraction(pdf_path):
    text_content = ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text_content += page.extract_text() + '\n'  # Add a newline character to separate pages
    return text_content
# Function to save the extracted text to a text file

def save_text_to_file(text, file_path):
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(text)

def expMatches(text, regex_pattern):
    matches = re.findall(regex_pattern, text)
    return matches

def tester_type(motif):
    motif_chiffres = r'^-?\d+.+$'
    motif_lettres = r'^[a-zA-Z]+$'
    motif_alphanumerique = r'^[a-zA-Z0-9]+$'
    motif_num_special = r'^[a-zA-Z0-9%]+$'

    if re.match(motif_chiffres, motif):
        return 0 #motif numerique
    elif re.match(motif_lettres, motif):
        return 1 #motif uniquement alphabetique
    elif re.match(motif_alphanumerique, motif):
        return 2 #motif chiffre et lettre
    elif re.match(motif_num_special, motif):
        return 3 #motif sans lettres
    else:
        return 4 #motif avec caracteres speciaux


chemin = 'data_repository/PDF1.pdf'

pdf_text = extraction(chemin)
save_text_to_file(pdf_text, './data_repository/test1.txt')
# Définir le motif d'expression régulière que vous souhaitez rechercher
regex_pattern = r'\b(VEMS)\b'

# Trouver les correspondances d'expression régulière dans le texte extrait
regex_matches = expMatches(pdf_text, regex_pattern)


print("Correspondances d'expression régulière trouvées:", regex_matches)


# Exemples de tests
print(tester_type("-0.62") ) 




def DonneeDunParametre(pdf_text):
    liste_param = []
    motif = fr'{parametre}\s+(.+)'
    correspondance = re.search(motif, pdf_text)

    while correspondance:
        donnees_apres = correspondance.group(1)
        liste_param.append(donnees_apres)

        if tester_type(donnees_apres) == 0:
            motif = fr'{donnees_apres}\s+(.+)'
            correspondance = re.search(motif, pdf_text)
        else:
            break  # Stop if the data type is as expected

    return liste_param



def extractionPatient(parametre, pdf_text):
    motif =fr'{parametre}\s([^\s]+)'
    correspondance = re.search(motif, pdf_text)
    return correspondance


print("//////////////////////////////////////////////")
print(DonneeDunParametre("VT",pdf_text))
print(DonneeDunParametre("VEMS",pdf_text))
print(extractionPatient("Nom",pdf_text))
#print(extractionPatient("Prénom",pdf_text))



