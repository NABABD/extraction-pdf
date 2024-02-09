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

def main():
    folder_path = choose_folder()
    if folder_path:
        print(f"User selected folder: {folder_path}")
        files = list_files(folder_path) #contenu du dossier
        if files:
            print("Files in the selected folder:")
            for file_name in files:
                print(file_name)
            
            n=len(files)
            with pandas.ExcelWriter('doubleData.xlsx', engine='xlsxwriter') as writer:
                li_parametre=['VEMS','CVF','CRF','CPT','DEMM','TEF']
                for i in range(0,n):
                    chemin=folder_path+'/'+files[i]
                    print(chemin)
                    pdf_text = extract_text_from_pdf(chemin)
                    fichier_test='test'+str(i)+'.txt'
                    save_text_to_file(pdf_text,fichier_test )
                    #print('test'+str(i)+'.txt')
                    # Use the existing extract_patient_data_from_text function to parse the extracted text
                    for x in range(len(li_parametre)):
                        feuille=extract_patient_data_from_text(pdf_text) 
                        feuille.update(parametre(li_parametre[x], fichier_test))
                        
                        if i==0:
                            df= pandas.DataFrame([feuille])
                            df.to_excel(writer, sheet_name=li_parametre[x],index=False)
                        
                        workbook  = writer.book
                        worksheet = writer.sheets[li_parametre [x]]
                        row = x
                        
                        worksheet.write(row, 0,'2')
                        col = 1
                        #for i in range(1,len(feuille1_data2)):
                        while col <len(feuille):
                            worksheet.write(row,col,list(feuille.values())[col])
                            print(list(feuille.values())[col])
                            col+=1 

                    """
                    feuille1 = extract_patient_data_from_text(pdf_text)
                    feuille2=feuille1_data.copy()
                    feuille3=feuille1_data.copy()
                    feuille4=feuille1_data.copy()
                    feuille5_data=feuille1_data.copy()
                    feuille6_data=feuille1_data.copy()
                    feuille1_data.update(parametre('VEMS', fichier_test))
                    feuille2_data.update(parametre('CVF',fichier_test))
                    feuille3_data.update(parametre('CRF',fichier_test))
                    feuille4_data.update(parametre('CPT',fichier_test))
                    feuille5_data.update(parametre('DEMM',fichier_test))
                    feuille6_data.update(parametre('TEF',fichier_test))
                    print(feuille1_data)                
                    if i==0:
                        df1 = pandas.DataFrame([feuille1_data])
                        df2 = pandas.DataFrame([feuille2_data])
                        df3 = pandas.DataFrame([feuille3_data])
                        df4 = pandas.DataFrame([feuille4_data])
                        df5 = pandas.DataFrame([feuille5_data])
                        df6 = pandas.DataFrame([feuille6_data])
                        df1.to_excel(writer, sheet_name='VEMS',index=False)
                        df2.to_excel(writer, sheet_name='CVF',index=False)
                        df3.to_excel(writer, sheet_name='CRF',index=False)
                        df4.to_excel(writer, sheet_name='CPT',index=False)
                        df5.to_excel(writer, sheet_name='DEMM',index=False)
                        df6.to_excel(writer, sheet_name='TEF',index=False)

                    else: 
                        workbook  = writer.book
                        worksheet = writer.sheets['VEMS']
                        row = 2
                        
                        worksheet.write(row, 0,'2')
                        col = 1
                        #for i in range(1,len(feuille1_data2)):
                        while col <len(feuille1_data):
                            worksheet.write(row,col,list(feuille1_data.values())[col])
                            print(list(feuille1_data.values())[col])
                            col+=1                                    
                        workbook  = writer.book
                        worksheet = writer.sheets['CVF']
                        row = 2
                        
                        worksheet.write(row, 0,'2')
                        col = 1
                        #for i in range(1,len(feuille1_data2)):
                        while col <len(feuille1_data):
                            worksheet.write(row,col,list(feuille1_data.values())[col])
                            print(list(feuille1_data.values())[col])
                            col+=1                                    
                        workbook  = writer.book
                        worksheet = writer.sheets['CRF']
                        row = 2
                        
                        worksheet.write(row, 0,'2')
                        col = 1
                        #for i in range(1,len(feuille1_data2)):
                        while col <len(feuille1_data):
                            worksheet.write(row,col,list(feuille1_data.values())[col])
                            print(list(feuille1_data.values())[col])
                            col+=1                                    
                        workbook  = writer.book
                        worksheet = writer.sheets['CPT']
                        row = 2
                        
                        worksheet.write(row, 0,'2')
                        col = 1
                        #for i in range(1,len(feuille1_data2)):
                        while col <len(feuille1_data):
                            worksheet.write(row,col,list(feuille1_data.values())[col])
                            print(list(feuille1_data.values())[col])
                            col+=1                                    
                        workbook  = writer.book
                        worksheet = writer.sheets['DEMM']
                        row = 2
                        
                        worksheet.write(row, 0,'2')
                        col = 1
                        #for i in range(1,len(feuille1_data2)):
                        while col <len(feuille1_data):
                            worksheet.write(row,col,list(feuille1_data.values())[col])
                            print(list(feuille1_data.values())[col])
                            col+=1                                    
                
                        workbook  = writer.book
                        worksheet = writer.sheets['TEF']
                        row = 2
                        
                        worksheet.write(row, 0,'2')
                        col = 1
                        #for i in range(1,len(feuille1_data2)):
                        while col <len(feuille1_data):
                            worksheet.write(row,col,list(feuille1_data.values())[col])
                            print(list(feuille1_data.values())[col])
                            col+=1                                    
                """



        else:
            print("No files found in the selected folder.")
    else:
        print("No folder selected.")
    
if __name__ == "__main__":
    main()