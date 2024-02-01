import fitz  # PyMuPDF
import re
import pdfplumber
import pandas 
from pandas import ExcelWriter
import xlsxwriter
from xlsxwriter import Workbook

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
    # Dictionary to store the extrac,ted data
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
    patient_data['TEST']=1
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

chemin = 'PDF1.pdf'

pdf_text = extract_text_from_pdf(chemin)
save_text_to_file(pdf_text, 'test1.txt')
# Définir le motif d'expression régulière que vous souhaitez rechercher
regex_pattern = r'\b(VEMS)\b'
# Use the existing extract_patient_data_from_text function to parse the extracted text
pdf1_data = extract_patient_data_from_text(pdf_text)
print(pdf1_data)

# Create a DataFrame with the extracted data
df = pandas.DataFrame([pdf1_data])
df2 = pandas.DataFrame([pdf1_data])

# Save the DataFrame to an Excel file
excel_path= 'PatientData.xlsx'

with pandas.ExcelWriter('multiple.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1',index=False)
    df2.to_excel(writer, sheet_name='Sheetc',index=False)

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    row = 2
    col = 0

    worksheet.write(row, col,     '2')
    worksheet.write(row, col + 1, '24')
    worksheet.write(row, col + 2, '56')
    worksheet.write(row, col + 3, '45')

"""workbook = xlsxwriter.Workbook("multiple.xlsx")
worksheet1 = workbook.add_worksheet()
"""