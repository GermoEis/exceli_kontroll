import os
import pandas as pd
import pythoncom
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox
import threading
import psycopg2
from dotenv import load_dotenv
from pathlib import Path
import sys

# Funktsioon projektijuure leidmiseks nii skripti kui EXE puhul
def get_project_root():
    if getattr(sys, "frozen", False):
        # Kui EXE-st, siis 2 kausta ülespoole EXE failist
        return Path(sys.executable).parent.parent
    else:
        # Kui .py failina, siis 2 kausta ülespoole skriptist
        return Path(__file__).resolve().parents[1]

# Määra .env faili tee ja lae see
project_root = get_project_root()
dotenv_path = project_root / "config" / ".env"

print(f".env faili laadimise tee: {dotenv_path}")

if not dotenv_path.exists():
    raise RuntimeError(f".env faili ei leitud: {dotenv_path}")

load_dotenv(dotenv_path=dotenv_path)
print(".env fail laetud")

# Kontrolli vajalike võtmete olemasolu
required_keys = ["DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD"]
missing = [key for key in required_keys if not os.getenv(key)]
if missing:
    raise RuntimeError(f"❌ Puuduvad .env võtmed: {', '.join(missing)}")

def töötlus():
    try:
        print("Töötlus algas")
        with open("temp_folder.txt", "r", encoding="utf-8") as f:
            input_folder = f.read().strip()

        base_path = r"Z:\scan"
        folder_name = os.path.join(base_path, input_folder)
        pohifail_path = os.path.join(folder_name, "Book1.xlsx")

        df_main = pd.read_excel(pohifail_path)

        db_config = {
            "host": os.getenv("DB_HOST"),
            "port": int(os.getenv("DB_PORT", 5432)),
            "dbname": os.getenv("DB_NAME"),
            "user": os.getenv("DB_USER"),
            "password": os.getenv("DB_PASSWORD")
        }

        conn_str = (
            f"host={db_config['host']} "
            f"port={db_config['port']} "
            f"dbname={db_config['dbname']} "
            f"user={db_config['user']} "
            f"password={db_config['password']}"
        )

        with psycopg2.connect(conn_str) as conn:
            df_meta = pd.read_sql_query('SELECT nimi, isikukood FROM public."Metadata"', conn)

        meta_dict = {}
        for _, row in df_meta.iterrows():
            nimi = str(row["nimi"]).strip() if pd.notna(row["nimi"]) else ""
            isikukood = str(row["isikukood"]).strip() if pd.notna(row["isikukood"]) else ""
            if not isikukood:
                continue
            meta_dict.setdefault(isikukood, []).append(nimi)

        def värv(row):
            isikukood = str(row.iloc[14]).strip() if pd.notna(row.iloc[14]) else ""
            nimi = str(row.iloc[13]).strip() if pd.notna(row.iloc[13]) else ""

            if isikukood in meta_dict:
                meta_nimed = [n.strip() for n in meta_dict[isikukood]]
                return "green" if nimi in meta_nimed else "orange"
            return None

        pythoncom.CoInitialize()
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(pohifail_path)
        ws = wb.Sheets(1)

        print("COM Excel avatud, alustame värvimist")
        for idx, row in df_main.iterrows():
            color = värv(row)
            if color == "green":
                ws.Cells(idx + 2, 14).Interior.Color = 0x00FF00  # roheline
            elif color == "orange":
                ws.Cells(idx + 2, 14).Interior.Color = 0x00A5FF  # oranž

        wb.Save()
        wb.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()

        print("Värvimine lõpetatud, fail salvestatud")
        root.after(0, töö_valmis)

    except Exception as e:
        print(f"Töötlus ebaõnnestus: {e}")
        root.after(0, lambda e=e: töö_viga(str(e)))

def töö_valmis():
    loading_window.destroy()
    messagebox.showinfo("Teade", "Excelite võrdlus tehtud.")

def töö_viga(veateade):
    loading_window.destroy()
    messagebox.showerror("Viga", f"Töötlemisel tekkis viga:\n{veateade}")

# GUI loomine
loading_window = tk.Tk()
loading_window.title("Laeb")
loading_window.geometry("300x100")
loading_window.resizable(False, False)
label = tk.Label(loading_window, text="Tegutsen, palun oota...", font=("Arial", 12))
label.pack(pady=30)
root = loading_window

# Käivita töötlus taustal
threading.Thread(target=töötlus, daemon=True).start()
loading_window.mainloop()
