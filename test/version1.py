import fitz  # PyMuPDF
import re
import pdfplumber

# Function to extract the entire text from a PDF file
def extract_text_from_pdf(pdf_path):
    text_content = ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text_content += page.extract_text() + '\n'  # Add a newline character to separate pages
    return text_content

# Function to save the extracted text to a text file
def save_text_to_file(text, file_path):
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(text)


# Function to extract patient data from the text content of a PDF
def extract_patient_data_from_text(text):
    # Dictionary to store the extracted data
    patient_data = {}

    # Regular expressions to find the relevant information
    name_regex = r"Nom:\s([^\s]+)"
    birth_date_regex = r"Date naissance:\s(\d{2}\/\d{2}\/\d{4})"
    test_date_regex = r"Date test\s(\d{2}\.\d{2}\.\d{2})"
    height_regex = r"Taille:\s(\d+\scm)"
    weight_regex = r"Poids:\s(\d+\.?\d*\skg)"
    bmi_regex = r"IMC:\s(\d+)"

    # Parameter headers and their corresponding regex
    parameter_headers = {
        "CVF Pré": r"CVF\s(\d+\.\d+)",
        "CVF Zscore": r"CVF.*ZScore\s(\-?\d+\.\d+)",
        # Additional parameters can be added here with their regex patterns
    }

    # Extracting information using regex
    patient_data['NOM'] = re.search(name_regex, text).group(1)
    patient_data['Date de naissance'] = re.search(birth_date_regex, text).group(1)
    patient_data['Date du test'] = re.search(test_date_regex, text).group(1)
    patient_data['taille'] = re.search(height_regex, text).group(1)
    patient_data['poids'] = re.search(weight_regex, text).group(1)
    patient_data['IMC'] = re.search(bmi_regex, text).group(1)

    # Extracting parameters
    for header, regex in parameter_headers.items():
        match = re.search(regex, text)
        patient_data[header] = match.group(1) if match else ''

    return patient_data

chemin = 'data_repository/PDF1.pdf'

pdf_text = extraction(chemin)
save_text_to_file(pdf_text, './data_repository/test1.txt')
# Définir le motif d'expression régulière que vous souhaitez rechercher
regex_pattern = r'\b(VEMS)\b'

# Trouver les correspondances d'expression régulière dans le texte extrait
regex_matches = expMatches(pdf_text, regex_pattern)


print("Correspondances d'expression régulière trouvées:", regex_matches)
