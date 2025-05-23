# C-Arm Angiography and DICOM Dose Data Export

## Description

This project contains Python scripts designed to extract dose data from PDF and DICOM files generated by three different C-Arm angiography systems (Siemens Axiom, Siemens Cios Alpha, and Philips BV Pulsera) and export the data into Excel sheets.

## Using the Executables
Executable files for different C-Arms and DICOM are available for download:

- Siemens Axiom: [Export_Axiom.exe](https://github.com/Shiramis/RDSR-to-Excel/releases/download/v1.0.0/Export_Axiom.exe)
- Siemens Cios Alpha: [Export_Cios.exe](https://github.com/Shiramis/RDSR-to-Excel/releases/download/v1.0.0/Export_Cios.exe)
- Philips BV Pulsera: [Export_Philips.exe](https://github.com/Shiramis/RDSR-to-Excel/releases/download/v1.0.0/Export_Philips.exe)
- Dose Report DICOM: [Exp_RS.exe](https://github.com/Shiramis/RDSR-to-Excel/releases/download/v1.2/Exp_RS.exe)
- Diagnostic DICOM: [export_dicom.exe](https://github.com/Shiramis/RDSR-to-Excel/releases/download/v1.0.0/export_dicom.exe)

  
1. Download the appropriate executable for your needs from the releases section.
2. Run the executable:
- Double-click the downloaded .exe file and follow the prompts to provide the path to the folder with PDF/DICOM files and the path for the output Excel file.
  
## File Descriptions

- `Export_Axiom.py`: Script for processing Siemens Axiom PDF files and exporting dose data to Excel.
- `Axiom_proccess.py`: Contains the processing functions for Siemens Axiom.
- `Export_Cios.py`: Script for processing Siemens Cios Alpha PDF files and exporting dose data to Excel.
- `Cios_proccess.py`: Contains the processing functions for Siemens Cios Alpha.
- `Export_Philips.py`: Script for processing Philips BV Pulsera PDF files and exporting dose data to Excel.
- `Philips_proccess.py`: Contains the processing functions for Philips BV Pulsera.
- `ExpDoseRep.py`: Script for processing DICOM Dose Report files from patient folders and exporting dose data to Excel.
- `export_dicom.py`: Script for processing DICOM files and exporting dose data to Excel.
