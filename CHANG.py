import os
import pandas as pd
import pdfplumber
import re
import tkinter as tk
import threading
from tkinter import filedialog, messagebox
import queue
import datetime


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
        line = line.strip()  # Trim whitespace
        if line.startswith(keyword) and (len(line) == len(keyword) or line[len(keyword)] in " \t"):
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


def get_first_line_with_param(file_path, param):
    with open(file_path, 'r') as file:
        for line in file:
            if line.startswith(param):
                return line.strip()  # Return the line without leading/trailing whitespace
    return None  # Return None if no line starting with 'CVF' is found


def validate_date(date_str):
    try:
        datetime.datetime.strptime(date_str, '%d.%m.%y')
        return True
    except ValueError:
        return False

def extract_patient_data_from_text(text):
    regex_patterns = {
        "NOM": r"Nom:\s([^\s]+)",
        "Date de naissance": r"Date naissance:\s(\d{2}/\d{2}/\d{4})",
        "Taille (cm)": r"Taille:\s(\d+)\s*cm",
        "Poids (kg)": r"Poids:\s(\d+\.?\d*)\s*kg",
        "IMC": r"IMC:\s(\d+)",
        "Age": r"Age:\s(\d+)\s*Années?",

    }

    date_du_test_pattern = re.compile(r"Date test\s(\d{2}\.\d{2}\.\d{2})\s*$", re.MULTILINE)

    patient_data = {}

    for key, pattern in regex_patterns.items():
        match = re.search(pattern, text)
        if match:
            patient_data[key] = match.group(1)

    # Initialize "Date du test" with a default value to ensure it exists in patient_data
    patient_data["Date du test"] = None  # or use "" for an empty string

    # Search for "Date du test" using the compiled pattern
    date_du_test_match = date_du_test_pattern.search(text)
    if date_du_test_match:
        patient_data["Date du test"] = date_du_test_match.group(1)

    return patient_data
def parametre(p, text):
    patient_data = {}

    if p in ['CV', 'CRF-He', 'CRFpl']:
        patient_data[f'{p}/Pré'] = get_nth_word(text, p, 2) or None  # f way to insert the value of the variable p directly into a string.
        theo_word = get_nth_word(text, p, 1)
        post_effort_word = get_nth_word(text, p, 4)
        zscore_word = get_nth_word(text, p, -1)

        if theo_word != zscore_word:
            patient_data[f'{p}/Zscore'] = zscore_word
        else:
            # If they are the same or there's no specific POST effort, leave it blank or set a default value
            patient_data[f'{p}/Zscore'] = None  # Or any default value you see fit
        # Check if the word at position 4 is not the same as the word at position -1
        # and set POST effort accordingly
        if post_effort_word != zscore_word:
            patient_data[f'{p}/POST effort'] = post_effort_word
        else:
            # If they are the same or there's no specific POST effort, leave it blank or set a default value
            patient_data[f'{p}/POST effort'] = None  # Or any default value you see fit

    elif p in ['CVF', 'DEMM', 'DEM25']:
        patient_data[f'{p}/Pré'] = get_nth_word(text, p, 2) or None  # f way to insert the value of the variable p directly into a string.
        patient_data[f'{p}/Zscore'] = get_nth_word(text, p, -1) or None

        post_effort_word = get_nth_word(text, p, 5)
        zscore_word = get_nth_word(text, p, -1)

        # Check if the word at position 4 is not the same as the word at position -1
        # and set POST effort accordingly
        if post_effort_word != zscore_word:
            patient_data[f'{p}/POST effort'] = post_effort_word
        else:
            # If they are the same or there's no specific POST effort, leave it blank or set a default value
            patient_data[f'{p}/POST effort'] = None  # Or any default value you see fit

    elif p == 'VEMS':
        # Extraction spécifique
        patient_data['VEMS/Pré'] = get_nth_word(text, 'VEMS', 2) or ''
        patient_data['VEMS/Zscore'] = get_nth_word(text, 'VEMS', -1) or ''
        if get_nth_word(text, 'VEMS', -1) != get_nth_word(text, 'VEMS', 5):
            patient_data['VEMS/POST effort'] = get_nth_word(text, 'VEMS', 5)
        else:
            patient_data['VEMS/POST effort'] = ''

        patient_data['VBEex/Pré'] = get_nth_word(text, 'VBEex', 1) or ''

        patient_data['VBEex/POST effort'] = get_nth_word(text, 'VBEex', 2) or ''
        patient_data['VBe%VF/pré'] = get_nth_word(text, 'VBe%VF', 1) or ''
        patient_data['VBe%VF/POST effort'] = get_nth_word(text, 'VBe%VF', 2) or ''

    elif p == 'CPT-He':
        line_cpt_he = get_nth_line_starting_with_keyword(text, 'CPT', 1)

        if line_cpt_he:  # Check if line_crf_he is not None
            patient_data['CPT-He/Pré'] = get_nth_word_from_line(line_cpt_he, 1)
            patient_data['CPT-He/Zscore'] = get_nth_word_from_line(line_cpt_he, 5)
        else:
            patient_data['CPT-He/Pré'] = ''
            patient_data['CPT-He/Zscore'] = ''
            patient_data['CPT-He/POST BD'] = ''

    elif p == 'CPT-Pl':
        line_cpt_pl = get_nth_line_starting_with_keyword(text, 'CPT', 2)
        if line_cpt_pl:  # Check if line_crf_he is not None
            patient_data['CPT-Pl/Pré'] = get_nth_word_from_line(line_cpt_pl, 2)
            patient_data['CPT-Pl/Zscore'] = get_nth_word_from_line(line_cpt_pl, -1)
        else:
            patient_data['CPT-Pl/Pré'] = ''
            patient_data['CPT-Pl/Zscore'] = ''
            patient_data['CPT-Pl/POST BD'] = ''

    elif p == 'TEF':
        patient_data['TEF/Pré'] = get_nth_word(text, 'TEF', 1) or ''
        patient_data['TEF/Zscore'] = ''
        patient_data['TEF/POST effort'] = get_nth_word(text, 'TEF', 2) or ''

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
    gui_queue = queue.Queue()
    def process_folder(folder_path):
        files = list_files(folder_path)
        # Check for non-PDF files
        non_pdf_files = [f for f in files if not f.lower().endswith('.pdf')]
        if non_pdf_files:
            gui_queue.put(("error_non_pdf", non_pdf_files))
            return

        # Assume no non-PDF files; proceed with processing
        if files:
            
            
            
            li_parametre = ['CV', 'VEMS', 'CRF-He', 'CRFpl', 'CVF', 'CPT-He', 'CPT-Pl', 'DEMM', 'TEF', 'DEM25']
            existing_data = {sheet_name: pd.DataFrame() for sheet_name in li_parametre}

            for file_name in files:
                if file_name.lower().endswith('.pdf'):
                    chemin = os.path.join(folder_path, file_name)
                    pdf_text = extract_text_from_pdf(chemin)

                    for param in li_parametre:
                        feuille = extract_patient_data_from_text(pdf_text)
                        
                        #print(feuille)
                        feuille.update(parametre(param, pdf_text))

                        if existing_data[param].empty:
                            df = pd.DataFrame([feuille])
                        else:
                            df = pd.concat([existing_data[param], pd.DataFrame([feuille])], ignore_index=True)

                        # Dynamically adjust the column order based on what's available
                        base_order = ['NOM', 'Date de naissance', 'Taille (cm)', 'Poids (kg)', 'IMC', 'Age']
                        # Only include 'Date du test' if it's actually a column in the current DataFrame
                        desired_order = base_order + (['Date du test'] if 'Date du test' in df.columns else []) + [col for col in df.columns if col not in base_order and col != 'Date du test']

                        df = df[desired_order]

                        existing_data[param] = df

                    #nom du fichier 
                    nom = feuille.get('NOM')
                    date_naissance = feuille.get('Date de naissance')
                    anonymat=nom[0]+date_naissance[-2]+date_naissance[-1]
                    nomfic=anonymat+'.xlsx'
                    print(anonymat)
                    #print(nom,nom[0],date_naissance,date_naissance[-2],date_naissance[-1])
                    excel_file_path = os.path.join(folder_path, nomfic)

                    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
                        for sheet_name in existing_data:
                            existing_data[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

            for param, df in existing_data.items():
                if "Age" in df.columns:
                    df["Age"] = pd.to_numeric(df["Age"], errors='coerce')  # Convert Age to numeric, safe for sorting
                    df = df.sort_values(by="Age")  # Sort the DataFrame by Age
                    existing_data[param] = df  # Update the existing data with sorted DataFrame
                df.to_excel(writer, sheet_name=param, index=False)


            gui_queue.put(("complete", excel_file_path))
        else:
            gui_queue.put(("error_no_files", None))

    def update_gui_from_queue():
        try:
            message, data = gui_queue.get_nowait()
            if message == "complete":
                # Ensure this message is clear and indicates the file path
                messagebox.showinfo("Processing Complete", f"Excel file has been successfully saved to: {data}")
            elif message == "error_non_pdf":
                messagebox.showerror("Error", "Non-PDF files detected. Please remove them and try again: " + ", ".join(data))
            elif message == "error_no_files":
                messagebox.showinfo("Error", "No files found in the selected folder.")
            elif message == "error_no_folder":
                messagebox.showinfo("Error", "No folder selected.")
            ask_for_next_folder()  # Moved out from conditions to ensure it's called regardless of the message type.
        except queue.Empty:
            pass  # If nothing in the queue, do nothing.
        finally:
            root.after(100, update_gui_from_queue)  # Re-schedule the check.

    def ask_for_next_folder():
        response = messagebox.askyesno("Query", "Do you want to process another folder?")
        if response:
            select_folder()
        else:
            root.destroy()  # This cleanly exits the program.

    def select_folder():
        folder_path = filedialog.askdirectory()
        if folder_path:
            threading.Thread(target=process_folder, args=(folder_path,)).start()
        else:
            ask_for_next_folder()  # Re-ask if no folder is selected.






    def recuperation_du_nom(fic):
        if fic.lower().endswith('.pdf'):
                        chemin = os.path.join(folder_path, file_name)
                        pdf_text = extract_text_from_pdf(chemin)

                        for param in li_parametre:
                            feuille = extract_patient_data_from_text(pdf_text)
    





    root = tk.Tk()
    root.withdraw()
    select_folder()
    update_gui_from_queue()
    root.mainloop()

if __name__ == "__main__":
    main()
