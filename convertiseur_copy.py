import os
import tkinter as tk
from tkinter import filedialog
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

    
#ouverture du dossier contenant les pdf
def choose_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    folder_path = filedialog.askdirectory()
    return folder_path
#lecture du dossier 
def list_files(folder_path):
    files = os.listdir(folder_path)
    return files

 
    
import pandas
from openpyxl import load_workbook

def main():
    folder_path = choose_folder()
    if folder_path:
        print(f"User selected folder: {folder_path}")
        files = list_files(folder_path)
        if files:
            print("Files in the selected folder:")
            combined_data = {parameter: {} for parameter in ['VEMS', 'CVF', 'CRF', 'CPT', 'DEMM', 'TEF']}
            
            for file_name in files:
                chemin = folder_path + '/' + file_name
                pdf_text = extract_text_from_pdf(chemin)
                fichier_test = 'test_' + file_name + '.txt'

                save_text_to_file(pdf_text, fichier_test)
                patient_data = extract_patient_data_from_text(pdf_text)

                # Update patient data with additional parameters
                for parameter_name in combined_data.keys():
                    parameter_data = parametre(parameter_name, fichier_test)
                    patient_data.update({f'{parameter_name}_data_{file_name}': parameter_data})

                # Update combined_data dictionary
                for parameter_name, data_dict in patient_data.items():
                    combined_data[parameter_name][file_name] = data_dict

            # Create Excel file with each sheet representing a parameter
            with pandas.ExcelWriter('combinedData.xlsx', engine='openpyxl') as writer:
                for parameter_name, data_dict in combined_data.items():
                    df = pandas.DataFrame(data_dict).transpose()
                    df.to_excel(writer, sheet_name=parameter_name, index=True, header=True)

            print("Data written to combinedData.xlsx successfully.")

if __name__ == "__main__":
    main()
