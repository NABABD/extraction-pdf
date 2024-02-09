#import fitz   PyMuPDF
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
#recupere la nième valeur d'un parametre 
def get_nth_word(file_path, keyword, n):
    with open(file_path, 'r') as file:
        for line in file:
            if line.startswith(keyword):
                words = line.split()
                if len(words) >= n+1:
                    nth_word = words[n]
                    return nth_word

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
    """parameter_headers = {
        "VEMS Pré": r"VEMS\s(\d+\.\d+)\s(\d+\.\d+)\s(\d+\.\d+)" ,
        "CVF Zscore": r"CVF.*ZScore\s(\-?\d+\.\d+)",
        # Additional parameters can be added here with their regex patterns
    }
    """
    # Extracting information using regex
    patient_data['TEST']=1
    patient_data['NOM'] = re.search(name_regex, text).group(1)
    patient_data['Date de naissance'] = re.search(birth_date_regex, text).group(1)
    patient_data['Date du test'] = re.search(test_date_regex, text).group(1)
    patient_data['taille'] = re.search(height_regex, text).group(1)
    patient_data['poids'] = re.search(weight_regex, text).group(1)
    patient_data['IMC'] = re.search(bmi_regex, text).group(1)

    return patient_data
    #return patient_data
def parametre(p,text):
    patient_data={}
    
    if p=='VEMS':
        #extraction spécifique
        patient_data['VEMS/Pré']=get_nth_word(text,'VEMS',2)
        patient_data['VEMS/Zscore']=get_nth_word(text,'VEMS',5)
        patient_data['VEMS/POST BD']=0
        patient_data['VBEex/Pré']=get_nth_word(text,'VBEex',1)
        patient_data['VBEex/Post BD']=0
        patient_data['VBe%VF/prè']=get_nth_word(text,'VBe%VF',1)
        patient_data['VBe%VF/POST BD']=0
        

        return patient_data
    elif p=='CVF':
        patient_data['CVF/Pré']=get_nth_word(text,'CVF',2)
        patient_data['CVF/Zscore']=get_nth_word(text,'CVF',5)
        patient_data['CVF/POST BD']=0

        return patient_data 
    
    elif p=='CRF':
        patient_data['CRF/Pré']=get_nth_word(text,'CRF',2)
        patient_data['CRF/Zscore']=get_nth_word(text,'CRF',5)
        patient_data['CRF/POST BD']=0

        return patient_data
    elif p=='CPT':
            patient_data['CPT/Pré']=get_nth_word(text,'CPT',2)
            patient_data['CPT/Zscore']=get_nth_word(text,'CPT',5)
            patient_data['CPT/POST BD']=0

            return patient_data 

    elif p=='DEMM':
        patient_data['DEMM/Pré']=get_nth_word(text,'DEMM',2)
        patient_data['DEMM/Zscore']=get_nth_word(text,'DEMM',5)
        patient_data['DEMM/POST BD']=0

        return patient_data 

    elif p=='TEF':
        patient_data['TEF/Pré']=get_nth_word(text,'TEF',1)
        patient_data['TEF/Zscore']=0
        patient_data['TEF/POST BD']=0

        return patient_data  

    




















chemin = 'PDF/PDF1.pdf'
pdf_text = extract_text_from_pdf(chemin)
save_text_to_file(pdf_text, 'test1.txt')

chemin2 = 'PDF/PDF2.pdf'
pdf_text2 = extract_text_from_pdf(chemin2)
save_text_to_file(pdf_text2, 'test2.txt')
# Définir le motif d'expression régulière que vous souhaitez rechercher
regex_pattern = r'\b(VEMS)\b'
# Use the existing extract_patient_data_from_text function to parse the extracted text
feuille1_data = extract_patient_data_from_text(pdf_text)
feuille2_data=feuille1_data.copy()
feuille3_data=feuille1_data.copy()
feuille4_data=feuille1_data.copy()
feuille5_data=feuille1_data.copy()
feuille6_data=feuille1_data.copy()

#pdf1_data = parametre(pdf_text)
feuille1_data.update(parametre('VEMS','test1.txt'))
feuille2_data.update(parametre('CVF','test1.txt'))
feuille3_data.update(parametre('CRF','test1.txt'))
feuille4_data.update(parametre('CPT','test1.txt'))
feuille5_data.update(parametre('DEMM','test1.txt'))
feuille6_data.update(parametre('TEF','test1.txt'))
print(feuille1_data)
# Create a DataFrame with the extracted data
df = pandas.DataFrame([feuille1_data])
df2 = pandas.DataFrame([feuille2_data])
df3 = pandas.DataFrame([feuille3_data])
df4 = pandas.DataFrame([feuille4_data])
df5 = pandas.DataFrame([feuille5_data])
df6 = pandas.DataFrame([feuille6_data])


#########
feuille1_data2 = extract_patient_data_from_text(pdf_text2)
feuille2_data2=feuille1_data2.copy()
feuille3_data2=feuille1_data2.copy()
feuille4_data2=feuille1_data2.copy()
feuille5_data2=feuille1_data2.copy()
feuille6_data2=feuille1_data2.copy()
feuille1_data2.update(parametre('VEMS', 'test2.txt'))
feuille2_data.update(parametre('CVF', 'test2.txt'))
feuille3_data.update(parametre('CRF', 'test2.txt'))
feuille4_data.update(parametre('CPT', 'test2.txt'))
feuille5_data.update(parametre('DEMM', 'test2.txt'))
feuille6_data.update(parametre('TEF', 'test2.txt'))
print(feuille1_data2)
#pdf1_data = parametre(pdf_text)

# Create a DataFrame with the extracted data

with pandas.ExcelWriter('Data.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='VEMS',index=False)
    df2.to_excel(writer, sheet_name='CVF',index=False)
    df3.to_excel(writer, sheet_name='CRF',index=False)
    df4.to_excel(writer, sheet_name='CPT',index=False)
    df5.to_excel(writer, sheet_name='DEMM',index=False)
    df6.to_excel(writer, sheet_name='TEF',index=False)

    workbook  = writer.book
    worksheet = writer.sheets['VEMS']
    row = 2
    
    worksheet.write(row, 0,'2')
    col = 1
    #for i in range(1,len(feuille1_data2)):
    while col <len(feuille1_data2):
        worksheet.write(row,col,list(feuille1_data2.values())[col])
        print(list(feuille1_data2.values())[col])
        col+=1
    #worksheet.write(row, col + 2, '56')
    #worksheet.write(row, col + 3, '45')