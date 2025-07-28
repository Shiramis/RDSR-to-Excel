import tkinter as tk
from tkinter import filedialog
import pydicom
import os
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from datetime import datetime as dt
import time

class DicomFile:
    def __init__(self, filename, file_path):
        self.filename = filename
        self.file_path = file_path

def sanitize_path(path):
    return path.strip('"')

def extract_and_format_age(age_value):
    """Extract the numeric part of the age and format it."""
    if isinstance(age_value, str):
        numeric_part = ''.join(filter(str.isdigit, age_value))
        if numeric_part:
            return int(numeric_part)
        else:
            raise ValueError(f"No numeric part found in value '{age_value}'")

def extract_data(dicom_data):
    #==Totals===
    DAPtotal = []  # Dose Area Product Total
    RPt = []  # Dose (RP) Total
    dstrp = []  # Distance Source to Reference Point
    fDAPt = []  # Fluoro Dose Area Product Total
    fRPt = []  # Fluoro Dose (RP) Total
    tftime = []  # Total Fluoro Time
    aDAPt = []  # Acquisition Dose Area Product Total
    aRPt = []  # Acquisition Dose (RP) Total
    rpd = []  # Reference Point Definition
    tatime = []  # Total Acquisition Time

    def search_sequence(sequence):
        global count  # Ensure the 'count' variable is accessible globally
        count = False  # Initialize 'count' at the start of the sequence search

        for item in sequence:
            # Check if the item has the Concept Name Code Sequence (0040,A043)
            if (0x0040, 0xA043) in item:
                concept_name_code_sequence = item[(0x0040, 0xA043)].value
                code_meaning = concept_name_code_sequence[0].get((0x0008, 0x0104), None)
                if code_meaning:
                    code_meaning_value = code_meaning.value
                    if code_meaning_value == 'Dose Area Product Total':
                        extract_numeric_value(item, (0x0040, 0xA300), DAPtotal)
                    elif code_meaning_value == 'Dose (RP) Total':
                        extract_numeric_value(item, (0x0040, 0xA300), RPt)
                    elif code_meaning_value == 'Distance Source to Reference Point':
                        extract_numeric_value(item, (0x0040, 0xA300), dstrp)
                    elif code_meaning_value == 'Fluoro Dose Area Product Total':
                        extract_numeric_value(item, (0x0040, 0xA300), fDAPt)
                    elif code_meaning_value == 'Fluoro Dose (RP) Total':
                        extract_numeric_value(item, (0x0040, 0xA300), fRPt)
                    elif code_meaning_value == 'Total Fluoro Time':
                        extract_numeric_value(item, (0x0040, 0xA300), tftime)
                    elif code_meaning_value == 'Acquisition Dose Area Product Total':
                        extract_numeric_value(item, (0x0040, 0xA300), aDAPt)
                    elif code_meaning_value == 'Acquisition Dose (RP) Total':
                        extract_numeric_value(item, (0x0040, 0xA300), aRPt, allow_empty=True)
                    elif code_meaning_value == 'Reference Point Definition':
                        extract_concept_code(item, (0x0040, 0xA168), rpd)
                    elif code_meaning_value == 'Total Acquisition Time':
                        extract_numeric_value(item, (0x0040, 0xA300), tatime)
            # Recursively search within nested sequences
            if (0x0040, 0xA730) in item:
                search_sequence(item[(0x0040, 0xA730)].value)

    # Function to extract numeric value from Measured Value Sequence
    def extract_numeric_value(item, tag, target_list, allow_empty=False):
        if tag in item:
            measured_value_sequence = item[tag].value
            if measured_value_sequence and (0x0040, 0xA30A) in measured_value_sequence[0]:
                numeric_value = measured_value_sequence[0][(0x0040, 0xA30A)].value
                target_list.append(float(numeric_value))
            elif allow_empty:
                target_list.append('empty')
            else:
                target_list.append("N/A")

    # Function to extract code value from Concept Code Sequence
    def extract_concept_code(item, tag, target_list):
        if tag in item:
            concept_code_sequence = item[tag].value
            if concept_code_sequence and (0x0008, 0x0104) in concept_code_sequence[0]:
                code_value = concept_code_sequence[0][(0x0008, 0x0104)].value
                target_list.append(code_value)
            else:
                target_list.append('N/A')
    # Start searching from the main sequence
    if (0x0040, 0xA730) in dicom_data:
        search_sequence(dicom_data[(0x0040, 0xA730)].value)

    return DAPtotal, RPt, dstrp, fDAPt, fRPt, tftime, aDAPt, aRPt, rpd, tatime

def read_dicom_files(folder_path):
    data_total = []
    dicom_files = [file for file in os.listdir(folder_path) if file.endswith('') or file.endswith('.dcm')]  # Add proper file extensions if needed
    first_file_processed = False
    dose_report_found = False
    for file in dicom_files:
        file_path = os.path.join(folder_path, file)
        dicom_data = pydicom.dcmread(file_path)
        first_file_processed = True
        sop_class_uid = dicom_data.get((0x0008, 0x0016), None)
        pname = ''.join(dicom_data.get('PatientName', 'N/A'))
        pname = pname.replace('^', ' ')
        study_date_str = dicom_data.get('StudyDate', 'N/A')
        if study_date_str != 'N/A':
            content_date = dt.strptime(study_date_str, '%Y%m%d')

        if sop_class_uid and sop_class_uid.value == "1.2.840.10008.5.1.4.1.1.88.67":

            DAPtotal, RPt, dstrp, fDAPt, fRPt, tftime, aDAPt, aRPt, rpd, tatime = extract_data(dicom_data)
            dose_report_found = True
            
            # Patient Age
            if (0x0010, 0x1010) in dicom_data:
                age = dicom_data[(0x0010, 0x1010)].value
                age = extract_and_format_age(age)
            else:
                # Calculate the age using the patient's birth date and the study date
                birth_date_str = dicom_data[(0x0010, 0x0030)].value  # Patient's Birth Date
                if birth_date_str != 'N/A' and study_date_str != 'N/A':

                    birth_date = dt.strptime(birth_date_str, '%Y%m%d')
                    age = content_date.year - birth_date.year - (
                            (content_date.month, content_date.day) < (birth_date.month, birth_date.day))

                else:
                    age = 'N/A'

            data_total.append({'Patient Name': pname,
                    'Patient ID': dicom_data.get('PatientID', 'N/A'),
                    'Age (years)': age,
                    'Manufacturer': dicom_data.get('Manufacturer', 'N/A'),
                    'Station Name': dicom_data.get('StationName', 'N/A'),
                    'Institution Name': dicom_data.get('InstitutionName', 'N/A'),
                    'Study Date': study_date_str,
                    'Performing Physician': dicom_data.get('PerformingPhysicianName', 'N/A'),
                    'Dose Area Product Total (Gym²)': DAPtotal[0] if len(DAPtotal) > 0 else 'N/A',
                    'Dose (RP) Total (Gy)': RPt[0] if len(RPt) > 0 else 'N/A',
                    'Fluoro Dose Area Product Total (μGym²)': fDAPt[0] if len(fDAPt) > 0 else 'N/A',
                    'Fluoro Dose (RP) Total (Gy)': fRPt[0] if len(fRPt) > 0 else 'N/A',
                    'Total Fluoro Time (s)': tftime[0] if len(tftime) > 0 else 'N/A',
                    'Acquisition Dose Area Product Total (Gym²)': aDAPt[0] if len(aDAPt) > 0 else 'N/A',
                    'Acquisition Dose (RP) Total (Gy)': aRPt[0] if len(aRPt) > 0 else 'N/A',
                    'Reference Point Definition (cm)': rpd[0] if len(rpd) > 0 else 'N/A',
                    'Total Acquisition Time (s)': tatime[0] if len(tatime) > 0 else 'N/A'})
        else:
            pass

    if not dose_report_found:
        print(f"No dose report found in folder {folder_path}")
    if first_file_processed:
        dftotal = pd.DataFrame(data_total)

        return dftotal, pname
    else:
        return None, None

author = os.getlogin()
# open a window
root = tk.Tk()
root.withdraw()
# Select DICOM folder
folder_path = filedialog.askdirectory(title="Select Folder with DICOM DATA")
folder_path = sanitize_path(folder_path)
# Select output folder for the new Excel file, starting from the parent of the DICOM folder
initial_output_dir = os.path.dirname(folder_path)
output_folder = filedialog.askdirectory(title="Select Folder to Save the Excel File", initialdir=initial_output_dir)
output_folder = sanitize_path(output_folder)
# Automatically create an Excel file path
timestamp = time.strftime("%Y-%m-%d_%H%M%S")
output_filename = f"Dose_Report_{timestamp}.xlsx"
output_directory = os.path.join(output_folder, output_filename)

start_time = time.time()
# Reference Value Levels
print("Reference Value Levels")
def get_float(prompt):
    try:
        return float(input(prompt))
    except ValueError:
        print(f"Invalid input for {prompt.strip(':')}.")
        return None
DAP_value = get_float("Dose Area Product (DAP) (Gym²):")
DOSERP_value = get_float("Dose (RP) (Gy):")
AT_value = get_float("Acquisition Time (s):")

dicom_files = [file for file in os.listdir(folder_path) if file.endswith('') or file.endswith('.dcm')]

kbs = []
times = []
max_time = 0
total = []
coun = 0

for file in dicom_files:
    file_path = os.path.join(folder_path, file)
    # time, size etc.
    processing_start_time = time.time()
    size_kb = os.path.getsize(file_path) / 1024  # Size in KB
    folder_name = os.path.basename(os.path.dirname(file_path))
    kbs.append(size_kb)
    #--------------------
    dftotal, pname = read_dicom_files(file_path)

    if dftotal is not None and not dftotal.empty:
        total.append(dftotal)
        coun += 1
        print(f"Processed file: {coun}")
    else:
        print(f"Skipped file (no data): {file}")
    # processing time
    processed_time = time.time() - processing_start_time
    times.append(processed_time)
    # Track maximum processing time and patient index
    if processed_time > max_time:
        max_time = processed_time
        max_patient = coun
        max_patient_name = folder_name
        max_kb = size_kb
        print(f"New maximum processing time: {max_time:.2f} seconds for Patient {max_patient} ({max_patient_name}, {size_kb:.2f} KB)")
    print(f"Time taken for Patient {coun}: {processed_time:.2f} seconds")

if total:
    with pd.ExcelWriter(output_directory, engine='openpyxl') as writer:
        dft = pd.concat(total, axis=0)
        dft.replace(0, "empty", inplace=True)
        dft.columns = pd.Index(
            [f'{col}_{i}' if dft.columns.duplicated()[i] else col for i, col in enumerate(dft.columns)])
        dft.index = pd.Index(
            [f'{idx}_{i}' if dft.index.duplicated()[i] else idx for i, idx in enumerate(dft.index)])
        dfstpat = dft.style.set_properties(**{'text-align': 'left', 'white-space': 'wrap'})
        dft = dfstpat.set_properties(**{'border': '1px solid black', 'border-collapse': 'collapse'})
        dft.to_excel(writer, sheet_name="Accumulated X-Ray Dose Data", index=False)
else:
    print("No valid data to write. Excel file was not created.")

if os.path.exists(output_directory):
    wb = openpyxl.load_workbook(output_directory)
    # Ensure at least one sheet is visible
    visible_sheets = [sheet for sheet in wb.sheetnames if wb[sheet].sheet_state == 'visible']
    if not visible_sheets:
        wb.active = wb.sheetnames[0]

    sheet1 = wb["Accumulated X-Ray Dose Data"]
    # Set the height of the first row (header)
    sheet1.row_dimensions[1].height = 30
    #--------
    # commets exceeding levels
    for i in range(2, 2 + coun):

        try:
            cell_value_e = float(sheet1[f'E{i}'].value)
            if cell_value_e > 0.05:
                sheet1[f'E{i}'].fill = openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                comment = Comment(f"Exceeds trigger level: 0.05 Gym²", author)
                sheet1[f'E{i}'].comment = comment
            elif cell_value_e > DAP_value:
                sheet1[f'E{i}'].fill = openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                comment = Comment(f"Exceeds reference level: {DAP_value} Gym²", author)
                sheet1[f'E{i}'].comment = comment
        except (TypeError, ValueError):
            pass  # Ignore cells that are None or not numbers

        try:
            cell_value_f = float(sheet1[f'F{i}'].value)
            if cell_value_f > 5:
                sheet1[f'F{i}'].fill = openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                comment = Comment(f"Exceeds trigger level: 5 Gy", author)
                sheet1[f'F{i}'].comment = comment
            elif cell_value_f > DOSERP_value:
                sheet1[f'F{i}'].fill = openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                comment = Comment(f"Exceeds reference level: {DOSERP_value} Gy", author)
                sheet1[f'F{i}'].comment = comment
        except (TypeError, ValueError):
            pass  # Ignore cells that are None or not numbers

        try:
            cell_value_m = float(sheet1[f'M{i}'].value)
            if cell_value_m > 3600:
                sheet1[f'M{i}'].fill = openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                comment = Comment(f"Exceeds trigger level: 1 hour", author)
                sheet1[f'M{i}'].comment = comment
            elif cell_value_m > AT_value:
                sheet1[f'M{i}'].fill = openpyxl.styles.PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                comment = Comment(f"Exceeds reference value: {AT_value} s", author)
                sheet1[f'M{i}'].comment = comment
        except (TypeError, ValueError):
            pass  # Ignore cells that are None or not numbers
    #-------    
    alignment_settings = openpyxl.styles.Alignment(wrap_text=True, horizontal='left')
    for col in range(1, 18):
        cell1 = sheet1.cell(row=1, column=col)
        cell1.alignment = alignment_settings
    column_widths = [13.5, 13, 10.5, 12.5, 18, 14.5, 15.5, 16, 11, 19, 14, 15, 14, 17, 15, 18, 15]
    for i, width in enumerate(column_widths, start=1):
        column_letter = get_column_letter(i)
        sheet1.column_dimensions[column_letter].width = width
    wb.save(output_directory)
else:
    print("Excel file was not created. No data to load.")



print("\nProcessing done!")