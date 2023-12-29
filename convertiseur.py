import fitz  # PyMuPDF
import re


# extraire toutes les données du fichier et le mettre dans la variable texte
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


chemin = 'PDF2.pdf'

pdf_text = extraction(chemin)

# Définir le motif d'expression régulière que vous souhaitez rechercher
regex_pattern = r'\b(VEMS)\b'

# Trouver les correspondances d'expression régulière dans le texte extrait
regex_matches = expMatches(pdf_text, regex_pattern)


print("Correspondances d'expression régulière trouvées:", regex_matches)


# Exemples de tests
print(tester_type("-0.62") ) 



def DonneeDunParametre(parametre):
    liste_param=[]
    motif = fr'{parametre}\s+(.+)'
    correspondance = re.search(motif, pdf_text)
    donnees_apres = correspondance.group(1)
    liste_param.append(donnees_apres)
    while tester_type(donnees_apres)==0:  #tant que la données suivant est une valeur pour le parametre on continue
        motif=fr'{donnees_apres}\s+(.+)'
        correspondance = re.search(motif, pdf_text)
        if correspondance:
            donnees_apres = correspondance.group(1)
            if tester_type(donnees_apres)==0:
                liste_param.append(donnees_apres)


    return liste_param
print("//////////////////////////////////////////////")
print(DonneeDunParametre("VT"))
print(DonneeDunParametre("VEMS"))
print(DonneeDunParametre("CVF"))


