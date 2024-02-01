from images import MakeExcelFromTiff
import os
import pandas as pd
import re

class ImageFilep:
    def __init__(self, filename, file_path):
        self.filename = filename
        self.file_path = file_path

f = input ('Write the path of the folder with .tiff DRL DATA: ')
ot = input ("Write the path of the excel file: ")

folder_path = f
output_directory = ot

def make_valid_sheet_name(sheet_name):
    # Replace invalid characters with "Invalid"
    return re.sub(r'[\/\\\?\*\[\]:]', 'Invalid', sheet_name)

with pd.ExcelWriter(output_directory, engine='openpyxl') as writer:
    for patient_folder in os.listdir(folder_path):
        patient_folder_path = os.path.join(folder_path, patient_folder)

        # Check if the item in the folder is a directory (a patient folder)
        if os.path.isdir(patient_folder_path):
            # Initialize dataframes for the patient folder
            dfper_combined = pd.DataFrame()
            dft_combined = pd.DataFrame()

            # Iterate through files in the patient folder
            for root, dirs, files in os.walk(patient_folder_path):
                for filename in files:
                    if filename.endswith(".tif"):
                        file_path = os.path.join(root, filename)

                        # Create an ImageFilep object and append it to the list
                        tif_file = ImageFilep(filename, file_path)

                        # Process the TIFF file and get dataframes
                        data = MakeExcelFromTiff()
                        dft, dfper = data.start_processing(f"{tif_file.file_path}")

                        # Combine dataframes for the patient folder
                        dfper_combined = pd.concat([dfper_combined, dfper])
                        dft_combined = pd.concat([dft_combined, dft])

            # Create a sheet for the patient folder
            sheet_name = make_valid_sheet_name(patient_folder)
            p_all = dfper_combined.to_dict(orient='records')
            # Flatten the list and remove empty keys
            flattened_list = [list(item.values())[0] for item in p_all if list(item.values())[0] != '']
            patient = ['DefaultName','DefaultID','DefaultCont','DefaultObs','DefaultNumExp']
            for i in range(0, len(flattened_list), 5):
                if flattened_list[i] != 'DefaultName':
                    patient [0] = flattened_list[i]
                if flattened_list[i+1] != "DefaultID":
                    patient[1] = flattened_list[i+1]
                if flattened_list[i+2] != "DefaultCont":
                    patient[2] = flattened_list[i+2]
                if flattened_list[i + 3] != "DefaultObs":
                    patient[3] = flattened_list[i + 3]
                if flattened_list[i + 4] != "DefaultNumExp":
                    patient[4] = flattened_list[i + 4]

            dfper_new = pd.DataFrame(patient,
                                  index=['Patient name', 'ID', 'Date', 'Performing physician', 'Number of exposures'],
                                  columns=[""])

            # Write both dfper_combined and dft_combined to the Excel sheet
            dfper_new.to_excel(writer, sheet_name=sheet_name, index=True)
            dft_combined.to_excel(writer, sheet_name=sheet_name, startrow=len(dfper_new) + 2, index=True)



