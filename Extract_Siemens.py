from orthone2 import make_excel
import os
import pandas as pd

class WordFile:
    def __init__(self, filename, file_path):
        self.filename = filename
        self.file_path = file_path

def check_and_rename_sheet(writer, sheet_name):
    sheet_name_base = sheet_name
    sheet_name_counter = 1

    while sheet_name in writer.sheets:
        sheet_name_counter += 1
        sheet_name = f"{sheet_name_base} ({sheet_name_counter})"

    return sheet_name

#f = input ('Write the path of the folder with DATA: ')
folder_path = r"C:\Users\ssavv\pythonProject\tests"
pdf_files = []

# Iterate through files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        file_path = os.path.join(folder_path, filename)

        # Create a WordFile object and append it to the list
        pdf_file = WordFile(filename, file_path)
        pdf_files.append(pdf_file)


#ot = input ("Write the path of the excel file: ")
output_directory = r"C:\Users\ssavv\pythonProject\Data.xlsx"
patients =[]
with pd.ExcelWriter(output_directory, engine='openpyxl') as writer:
    for pdf_file in pdf_files:
        data = make_excel()

        df, dft, dfper, name_id1, name_id2 = data.startpro(f"{pdf_file.file_path}")

        sheet_name_df = f"{name_id1}"

        start_row_dfper = 0
        start_row_df = start_row_dfper + len(dfper) + 2
        dfper.to_excel(writer, sheet_name=sheet_name_df, startrow=start_row_dfper)
        df.to_excel(writer, sheet_name=sheet_name_df, startrow=start_row_df)

        patients.append(dft)
    dfpat = pd.concat(patients, axis=0)
    dfpat.to_excel(writer, sheet_name="Accumulated X-Ray Dose Data")

def check_and_rename_sheet(writer, sheet_name):
    while sheet_name in writer.sheets:
        # Extract the base sheet name and the existing counter (if any)
        base_sheet_name, counter = sheet_name, 1
        while True:
            new_sheet_name = f"{base_sheet_name} ({counter})"
            if new_sheet_name not in writer.sheets:
                return new_sheet_name
            counter += 1