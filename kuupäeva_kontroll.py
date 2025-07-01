import os
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import simpledialog, messagebox
import pythoncom
import win32com.client as win32

# --- Loeme kausta nime eelnevalt salvestatud failist ---
try:
    with open("temp_folder.txt", "r", encoding="utf-8") as f:
        input_folder = f.read().strip()
except FileNotFoundError:
    print("❌ Kausta nime faili ei leitud. Käivita esmalt esimene skript.")
    exit()

base_path = r"Z:\scan"
folder_name = os.path.join(base_path, input_folder)
failitee = os.path.join(folder_name, "Book1.xlsx")

if not os.path.exists(failitee):
    print("❌ Exceli faili ei leitud:", failitee)
    exit()

# --- GUI: Aasta ja nädala sisestus ---
root = tk.Tk()
root.withdraw()
try:
    aasta = int(simpledialog.askstring("Aasta", "Sisesta aasta (nt 2025):"))
    nadal = int(simpledialog.askstring("Nädal", "Sisesta nädala number (nt 19):"))
except (TypeError, ValueError):
    messagebox.showerror("Viga", "Sisestus ebaõnnestus või katkestati.")
    exit()

# --- Funktsioon, mis leiab antud ISO aasta ja nädala esmaspäeva ---
def iso_week_start(year, week):
    fourth_jan = datetime(year, 1, 4)
    delta = timedelta(days=fourth_jan.isoweekday() - 1)
    week1_monday = fourth_jan - delta
    target_monday = week1_monday + timedelta(weeks=week - 1)
    return target_monday

# --- Funktsioon, mis kontrollib, kas kuupäev kuulub valitud nädala sisse ---
def is_date_in_iso_week(date_obj, year, week):
    start = iso_week_start(year, week).replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=6, hours=23, minutes=59, seconds=59, microseconds=999999)
    date_check = date_obj.replace(hour=0, minute=0, second=0, microsecond=0)
    return start <= date_check <= end

# --- COM Excel: säilitab XML mappingud ---
pythoncom.CoInitialize()
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(failitee)
ws = wb.Sheets(1)

# --- Värvid ---
xlNone = -4142
oranž_värv = 0x00A5FF  # RGB(255,165,0)

for rida in range(2, ws.UsedRange.Rows.Count + 1):
    lahter = ws.Cells(rida, 17)  # Veerg Q
    väärtus = lahter.Value
    värv = lahter.Interior.Color

    if värv == 0x0000FF or värv == oranž_värv:
        continue

    if väärtus is None:
        lahter.Interior.Color = oranž_värv
        continue

    try:
        if isinstance(väärtus, float):
            kuup = datetime.fromordinal(datetime(1899, 12, 30).toordinal() + int(väärtus))
        elif isinstance(väärtus, datetime):
            kuup = väärtus
        else:
            s = str(väärtus).strip()
            possible_formats = ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d")
            for fmt in possible_formats:
                try:
                    kuup = datetime.strptime(s, fmt)
                    break
                except ValueError:
                    continue
            else:
                raise ValueError("Sobimatu kuupäeva formaat")
    except Exception:
        lahter.Interior.Color = oranž_värv
        continue

    kuulub = is_date_in_iso_week(kuup, aasta, nadal)

    if kuulub:
        lahter.Interior.ColorIndex = xlNone
    else:
        lahter.Interior.Color = oranž_värv

# --- Salvesta ja sulge ---
wb.Save()
wb.Close(False)
excel.Quit()
pythoncom.CoUninitialize()

messagebox.showinfo("Teade", "Kuupäevade kontroll tehtud.")
