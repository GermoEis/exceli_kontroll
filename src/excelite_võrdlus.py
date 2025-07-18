import os
import pandas as pd
import pythoncom
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox
import threading
import psycopg2
from dotenv import load_dotenv

# Lae keskkonnamuutujad .env failist
load_dotenv()

def töötlus():
    try:
        print("Töötlus algas")
        with open("temp_folder.txt", "r", encoding="utf-8") as f:
            input_folder = f.read().strip()

        base_path = r"Z:\scan"
        folder_name = os.path.join(base_path, input_folder)
        pohifail_path = os.path.join(folder_name, "Book1.xlsx")

        df_main = pd.read_excel(pohifail_path)

        # ✅ PostgreSQL ühendus .env failist
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

        # ✅ Metaandmete sõnastiku loomine
        meta_dict = {}
        for _, row in df_meta.iterrows():
            nimi = str(row["nimi"]).strip() if pd.notna(row["nimi"]) else ""
            isikukood = str(row["isikukood"]).strip() if pd.notna(row["isikukood"]) else ""
            if not isikukood:
                continue
            if isikukood in meta_dict:
                meta_dict[isikukood].append(nimi)
            else:
                meta_dict[isikukood] = [nimi]

        # ✅ Värvifunktsioon
        def värv(row):
            isikukood = str(row.iloc[14]).strip() if pd.notna(row.iloc[14]) else ""
            nimi = str(row.iloc[13]).strip() if pd.notna(row.iloc[13]) else ""

            if isikukood in meta_dict:
                meta_nimed = [n.strip() for n in meta_dict[isikukood]]
                if nimi in meta_nimed:
                    return "green"
                else:
                    return "orange"
            return None

        # ✅ COM Exceli kasutamine
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

# ✅ GUI osa
loading_window = tk.Tk()
loading_window.title("Laeb")
loading_window.geometry("300x100")
loading_window.resizable(False, False)
label = tk.Label(loading_window, text="Tegutsen, palun oota...", font=("Arial", 12))
label.pack(pady=30)
root = loading_window

threading.Thread(target=töötlus, daemon=True).start()
loading_window.mainloop()
