import time
import os
from dotenv import load_dotenv
from sqlalchemy import create_engine, text
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import subprocess
from expfunctions import expfun
# --------------------------------------------------
# DB connection and version check
load_dotenv()                                              # read .env
DB_URL = os.getenv("DB_URL")                               # get DB_URL from .env
engine = create_engine(DB_URL, pool_pre_ping=True)
# Check if the connection is successful and print the version
with engine.connect() as connection:
    result = connection.execute(text("SELECT version();"))
    for row in result:
        print(row)
# --------------------------------------------------
# Set up output directory
def select_folder(title):
    root = tk.Tk()
    root.withdraw()  # Απόκρυψη του βασικού παραθύρου
    folder_selected = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_selected

class PdfFile:
    def __init__(self, filename, file_path):
        self.filename = filename
        self.file_path = file_path

# ------------------------------
folder_path = select_folder("📁 Select the folder with PDF DATA")
base_output_folder = select_folder("📁 Select the OUTPUT folder (Excel + JSON)")
# ------------------------------
# Initialization of empty and unavailable data
pdf_files = []
patients = []
individual =[]
index = 0
nst = 0
nfl = 0
d = {}
max_time = 0
times = []
kbs = []
# ------------------------------
# Remove string-based append and unify with PdfFile objects
for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.lower().endswith(".pdf"):
            file_path = os.path.join(root, file)
            pdf_files.append(PdfFile(file, file_path))            
# ------------------------------
# Dynamic input for export options
print("\n--- Export Options ---")
EXPORT_EXCEL = input("Export to Excel? (y/n): ").strip().lower() == 'y'
EXPORT_JSON  = input("Export to JSON? (y/n): ").strip().lower() == 'y'
EXPORT_SQL = input("Export to SQL? (y/n): ").strip().lower() == 'y'

print("\nSelected export modes:")
print(f" - Excel: {'ON' if EXPORT_EXCEL else 'OFF'}")
print(f" - JSON:  {'ON' if EXPORT_JSON else 'OFF'}")
print(f" - SQL:   {'ON' if EXPORT_SQL else 'OFF'}")
print("------------------------------\n")

# 2. Create timestamped main output folder
timestamp = time.strftime("%Y-%m-%d_%H%M%S")
main_output_dir = Path(base_output_folder) / f"Dose_Report_{timestamp}"
main_output_dir.mkdir(parents=True, exist_ok=True)

# 3. Create subfolders
excel_output_dir = main_output_dir / "excel"
json_output_dir  = main_output_dir / "json"
sql_output_dir   = main_output_dir / "sql"
# Create directories if they do not exist
excel_output_dir.mkdir(parents=True, exist_ok=True)
json_output_dir.mkdir(parents=True, exist_ok=True)
sql_output_dir.mkdir(parents=True, exist_ok=True)
# 4. Define output file path for Excel
excel_output_file = excel_output_dir / "Dose_Report.xlsx"
sql_dump_file = sql_output_dir / f"dump_{timestamp}.sql"
# -------------------------------------
print(f"Excel output: {excel_output_file}")
print(f"JSON output folder: {json_output_dir}")
print(f"SQL output folder: {sql_output_dir}")
# Excel export
if EXPORT_EXCEL:
    expfun.export_to_excel(excel_output_file, pdf_files, patients, individual, d, index, nst, nfl)
#-------------------------------------------------------------------------------------------------
# JSON export
if EXPORT_JSON:
    expfun.export_to_json_fhir(json_output_dir, pdf_files, patients, individual, d, index, nst, nfl)
#-------------------------------------------------------------------------------------------------
# SQL export
if EXPORT_SQL:
    expfun.export_to_sql(pdf_files)
    sql_dump_file = sql_output_dir / f"dump_{timestamp}.sql"
    subprocess.run([
        "C:/Program Files/PostgreSQL/17/bin/pg_dump.exe",
        "-U", "postgres",
        "-h", "localhost",
        "-d", "medical_db",
        "-f", str(sql_dump_file)
        ], check=True)
    print(f"SQL dump saved to: {sql_dump_file}")
#-------------------------------------------------------------------------------------------------
# Print completion message
print("\nProcessing done!")