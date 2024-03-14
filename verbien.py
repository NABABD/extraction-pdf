import os
import pandas as pd
import pdfplumber
import re
import tkinter as tk
from tkinter import filedialog

def extract_text_from_pdf(pdf_path):
    text_content = ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text_content += page.extract_text() + '\n'  # Add a newline character to separate pages
    return text_content

def save_text_to_file(text, file_path):
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(text)

def get_nth_word(text, keyword, n):
    for line in text.splitlines():
        if line.startswith(keyword):
            words = line.split()
            if len(words) > n:
                return words[n]
    return None
def get_nth_word_from_line(line, n):

    words = line.split()  # Split the line into words
    if len(words) > n:
        return words[n]  # Return the nth word if it exists
    else:
        return None
def get_nth_line_starting_with_keyword(text, keyword, occurrence_number):
    current_occurrence = 0
    for line in text.splitlines():
        if line.strip().startswith(keyword):
            current_occurrence += 1
            if current_occurrence == occurrence_number:
                return line
    return None

def get_first_line_with_param(file_path,param):
    with open(file_path, 'r') as file:
        for line in file:
            if line.startswith(param):
                return line.strip()  # Return the line without leading/trailing whitespace
    return None  # Return None if no line starting with 'CVF' is found

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

    # Extracting information using regex
    patient_data['NOM'] = re.search(name_regex, text).group(1) if re.search(name_regex, text) else None
    patient_data['Date de naissance'] = re.search(birth_date_regex, text).group(1) if re.search(birth_date_regex, text) else None
    patient_data['Date du test'] = re.search(test_date_regex, text).group(1) if re.search(test_date_regex, text) else None
    patient_data['taille'] = re.search(height_regex, text).group(1) if re.search(height_regex, text) else None
    patient_data['poids'] = re.search(weight_regex, text).group(1) if re.search(weight_regex, text) else None
    patient_data['IMC'] = re.search(bmi_regex, text).group(1) if re.search(bmi_regex, text) else None

    return patient_data


def parametre(p, text):
    patient_data = {}

    if p == 'CV':
        patient_data['CV/Pré'] = get_nth_word(text, 'CV', 2) or 0
        patient_data['CV/Zscore'] = get_nth_word(text, 'CV', -1) or 0
        if (get_nth_word(text, 'CV', -1)!=get_nth_word(text, 'CV', 4)):
            patient_data['CV/POST effort'] = get_nth_word(text, 'CV', 4)
        else:
            patient_data['CV/POST effort'] = ''


    elif p=='CRF':
        patient_data['CRF/Pré'] = get_nth_word(text, 'CRF', 2) or 0
        patient_data['CRF/Zscore'] = get_nth_word(text, 'CRF', 5) or 0
        patient_data['CRF/POST BD'] = 0

    elif p == 'VEMS':
        # Extraction spécifique
        patient_data['VEMS/Pré'] = get_nth_word(text, 'VEMS', 2) or 0
        patient_data['VEMS/Zscore'] = get_nth_word(text, 'VEMS', -1) or 0
        if (get_nth_word(text, 'VEMS', -1)!=get_nth_word(text, 'VEMS', 5)):
            patient_data['VEMS/POST effort'] = get_nth_word(text, 'VEMS', 5)
        else:
            patient_data['VEMS/POST effort'] = ''

        patient_data['VBEex/Pré'] = get_nth_word(text, 'VBEex', 1) or 0
        patient_data['VBEex/POST effort'] = get_nth_word(text, 'VBEex', 2)
        patient_data['VBEex/Post BD'] = get_nth_word(text,'VBEex',5)
        patient_data['VBe%VF/pré'] = get_nth_word(text, 'VBe%VF', 1) or 0
        patient_data['VBe%VF/POST BD'] = get_nth_word(text,'VBe%VF',2)

    elif p == 'CVF':
        patient_data['CVF/Pré'] = get_nth_word(text, 'CVF', 2) or 0
        patient_data['CVF/Zscore'] = get_nth_word(text, 'CVF', -1) or 0
        if (get_nth_word(text, 'CVF', -1)!=get_nth_word(text, 'CVF', 5)):
            patient_data['CVF/POST effort'] = get_nth_word(text, 'CVF', 5)
        else:
            patient_data['CVF/POST effort'] = ''




    elif p == 'CPT-He':
        line_cpt_he = get_nth_line_starting_with_keyword(text, 'CPT', 1)
        print(f"Debug CPT-He line: {line_cpt_he}")  # Debug print

        if line_cpt_he:  # Check if line_crf_he is not None
            patient_data['CPT-He/Pré'] = get_nth_word_from_line(line_cpt_he, 1) or 0
            patient_data['CPT-He/Zscore'] = get_nth_word_from_line(line_cpt_he, 5) or 0
        else:
            patient_data['CPT-He/Pré'] = 0
            patient_data['CPT-He/Zscore'] = 0
            patient_data['CPT-He/POST BD'] = 0

        return patient_data
    elif p == 'CPT-Pl':
        line_cpt_pl = get_nth_line_starting_with_keyword(text, 'CPT', 2)
        if line_cpt_pl:  # Check if line_crf_he is not None
            patient_data['CPT-Pl/Pré'] = get_nth_word_from_line(line_cpt_pl, 2)
            patient_data['CPT-Pl/Zscore'] = get_nth_word_from_line(line_cpt_pl, -1)
        else:
            patient_data['CPT-Pl/Pré'] = 0
            patient_data['CPT-Pl/Zscore'] = 0
            patient_data['CPT-Pl/POST BD'] = 0

        return patient_data


    elif p == 'DEMM':
        patient_data['DEMM/Pré'] = get_nth_word(text, 'DEMM', 2) or 0
        patient_data['DEMM/Zscore'] = get_nth_word(text, 'DEMM', 5) or 0
        if (get_nth_word(text, 'DEMM', -1)!= get_nth_word(text, 'DEMM', 5)):

            patient_data['DEMM/POST effort'] = get_nth_word(text, 'DEMM', 5)
        else:
            patient_data['DEMM/POST effort']=''



    elif p == 'DEM25':
        patient_data['DEM25/Pré'] = get_nth_word(text, 'DEM25', 2) or 0
        patient_data['DEM25/Zscore'] = get_nth_word(text, 'DEM25', 5) or 0
        if (get_nth_word(text, 'DEM25', -1)!= get_nth_word(text, 'DEM25', 5)):

            patient_data['DEM25/POST effort'] = get_nth_word(text, 'DEM25', 5)
        else:
            patient_data['DEM25/POST effort']=''



    elif p == 'TEF':
        patient_data['TEF/Pré'] = get_nth_word(text, 'TEF', 1) or 0
        patient_data['TEF/Zscore'] = ''
        patient_data['TEF/POST effort'] = get_nth_word(text, 'TEF', 2)

    else:
        return patient_data

    return patient_data


def choose_folder():
    root = tk.Tk()
    root.withdraw()  
    folder_path = filedialog.askdirectory()
    return folder_path

def list_files(folder_path):
    files = os.listdir(folder_path)
    return files

def main():
    folder_path = choose_folder()
    if folder_path:
        print(f"User selected folder: {folder_path}")
        files = list_files(folder_path)
        if files:
            print("Files in the selected folder:")
            for file_name in files:
                print(file_name)

            n = len(files)
            with pd.ExcelWriter('doubleData.xlsx', engine='xlsxwriter') as writer:
                li_parametre = ['VEMS', 'CVF', 'CPT-He', 'CPT-Pl', 'DEMM', 'TEF', 'DEM25']
                existing_data = {sheet_name: pd.DataFrame() for sheet_name in li_parametre}

                for i in range(n):
                    chemin = folder_path + '/' + files[i]
                    print(chemin)
                    pdf_text = extract_text_from_pdf(chemin)
                    fichier_test = 'test' + str(i) + '.txt'
                    save_text_to_file(pdf_text, fichier_test)

                    for param in li_parametre:
                        feuille = extract_patient_data_from_text(pdf_text)
                        feuille.update(parametre(param, pdf_text))

                        # Check if the existing_data dictionary for the current parameter is empty
                        if existing_data[param].empty:
                            df = pd.DataFrame([feuille])
                        else:
                            # Concatenate the new DataFrame with the existing DataFrame for the current parameter
                            df = pd.concat([existing_data[param], pd.DataFrame([feuille])], ignore_index=True)

                        # Update the existing_data dictionary with the updated DataFrame for the current parameter
                        existing_data[param] = df

                        # Write the DataFrame to the Excel file
                        df.to_excel(writer, sheet_name=param, index=False)

        else:
            print("No files found in the selected folder.")
    else:
        print("No folder selected.")

if __name__ == "__main__":
    main()
