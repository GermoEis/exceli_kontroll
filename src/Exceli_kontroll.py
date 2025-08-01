import os
import re
import pythoncom
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox
import threading
from datetime import datetime

# --- Validaatorfunktsioonid --- #
def levenshtein(s1, s2):
    len1, len2 = len(s1), len(s2)
    d = [[0] * (len2 + 1) for _ in range(len1 + 1)]
    for i in range(len1 + 1):
        d[i][0] = i
    for j in range(len2 + 1):
        d[0][j] = j
    for i in range(1, len1 + 1):
        for j in range(1, len2 + 1):
            cost = 0 if s1[i - 1] == s2[j - 1] else 1
            d[i][j] = min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost)
    return d[len1][len2]

def on_õige_nimi_formaat(nimi):
    keelatud = ";:!()_^%$#@|»*+[]/"
    for ch in keelatud:
        if ch in nimi:
            return False
    if "  " in nimi:
        parts = nimi.split("  ")
        if len(parts) > 1 and len(parts[1].strip()) == 1:
            return False
    return True

def on_õige_isikukood_formaat(isikukood):
    keelatud = ",.;:!()_^%$#@|lZCVBNMASDFGHJKLPOIUYTREWQÜÕÖÄzcvbnm'asdfghjklööäõüpoiuytrewq[]·/"
    for ch in keelatud:
        if ch in isikukood:
            return False
    if "  " in isikukood:
        parts = isikukood.split("  ")
        if len(parts) > 1 and len(parts[1].strip()) == 1:
            return False
    return True

def on_õige_pangakonto_formaat(konto):
    if " " in konto:
        return False
    if len(konto) >= 12 and konto.startswith("EE"):
        return konto[2:].isdigit()
    return False

def on_õige_kuupäeva_formaat(kuupäev):
    try:
        parts = kuupäev.split()
        if len(parts) == 1:
            datetime.strptime(parts[0], "%d.%m.%Y")
        elif len(parts) == 2:
            try:
                datetime.strptime(parts[0] + " " + parts[1], "%d.%m.%Y %H:%M:%S")
            except ValueError:
                datetime.strptime(parts[0] + " " + parts[1], "%d.%m.%Y %H:%M")
        else:
            return False
        return True
    except ValueError:
        return False

def on_õige_dokumendi_number(dok_nr):
    dok_nr = dok_nr.strip()
    if re.fullmatch(r"\d{3,}", dok_nr):
        return True  # lubatud: ainult numbrid, vähemalt 3 pikad
    if re.fullmatch(r"\d{2}-\d{6}-[A-Z]{2}( LISA)?", dok_nr):
        return True  # lubatud: vana eriformaat
    return False  # kõik muu on vigane


# --- Põhitöötlus --- #
def töötlus():
    try:
        with open("temp_folder.txt", "r", encoding="utf-8") as f:
            folder = f.read().strip()

        path = os.path.join(r"Z:\\scan", folder, "Book1.xlsx")
        if not os.path.exists(path):
            raise FileNotFoundError(f"Faili ei leitud: {path}")

        pythoncom.CoInitialize()
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(path)
        ws = wb.Sheets(1)

        red = 255
        orange = 0x0078FF

        nimiCol = 14
        isikukoodCol = 15
        kontaktitüüpCol = 16
        kuupäevCol = 17
        dokumendiNrCol = 18
        pangakontoCol = 19

        max_row = ws.UsedRange.Rows.Count

        ettevottesonad = ["OÜ","ÕÜ", "AS", "KÜ", "SIA", "OY", "HÜ", "UAB", "AB", "ÜÜ", "ÄÜ", "ASUTAMISEL", "MTÜ", "KOGUDUS", "SAATKOND", "FIE"]
        tahedsonad = ["KONTAKT", "ANDMED", "CLIENT", "INFO", "EES", "ERE", ".IQ", ".OI", "ISIKUKOOD"]

        for r in range(2, max_row + 1):
            nimi = str(ws.Cells(r, nimiCol).Value or "").strip()
            isikukood = str(ws.Cells(r, isikukoodCol).Value or "").strip()
            kontaktitüüp = str(ws.Cells(r, kontaktitüüpCol).Value or "").strip()
            if kontaktitüüp.replace('.', '', 1).isdigit():
                kontaktitüüp = str(int(float(kontaktitüüp)))                
            kuupäev = str(ws.Cells(r, kuupäevCol).Value or "").strip()
            dokumendiNr = str(ws.Cells(r, dokumendiNrCol).Value or "").strip()
            pangakonto = str(ws.Cells(r, pangakontoCol).Value or "").strip()

        

            # Kontrolli, kas veerud on tühjad – igal juhul värvitakse punaseks
            if not nimi:
                ws.Cells(r, nimiCol).Interior.Color = red
            if not isikukood:
                ws.Cells(r, isikukoodCol).Interior.Color = red
            if not kontaktitüüp:
                ws.Cells(r, kontaktitüüpCol).Interior.Color = red
            if not kuupäev:
                ws.Cells(r, kuupäevCol).Interior.Color = red
            if not dokumendiNr:
                ws.Cells(r, dokumendiNrCol).Interior.Color = red
            if not pangakonto:
                ws.Cells(r, pangakontoCol).Interior.Color = red


            # Kui isikukood tühi või "XXX", värvi ainult isikukood
            if isikukood == "" or isikukood.upper() == "XXX":
                ws.Cells(r, isikukoodCol).Interior.Color = red

            else:
                if kontaktitüüp == "81":
                    # Peab olema 11 märki ja õige vorming
                    if len(isikukood) != 11 or not on_õige_isikukood_formaat(isikukood):
                        ws.Cells(r, isikukoodCol).Interior.Color = red
                        ws.Cells(r, kontaktitüüpCol).Interior.Color = red

                elif kontaktitüüp == "80":
                    # Kontrolli ainult siis, kui pikkus on 11
                    if len(isikukood) < 11 and not on_õige_isikukood_formaat(isikukood):
                        ws.Cells(r, isikukoodCol).Interior.Color = red
                        ws.Cells(r, kontaktitüüpCol).Interior.Color = red
                    # Kui alla 11, siis EI TEE MIDAGI (ei värvi)

                else:
                    # Muud kontaktitüübid: kontrolli ainult kui pikkus on 11
                    if len(isikukood) == 11 and not on_õige_isikukood_formaat(isikukood):
                        ws.Cells(r, isikukoodCol).Interior.Color = red
                        if kontaktitüüp:
                            ws.Cells(r, kontaktitüüpCol).Interior.Color = red
                    # Kui pikkus alla 11, EI TEE MIDAGI

                # Lisakontroll: kui pole ettevõtte nimi ja kontaktitüüp on 80, aga isikukood on liiga lühike
                if not any(re.search(r'\b' + re.escape(s) + r'\b', nimi.upper()) for s in ettevottesonad):
                    if kontaktitüüp == "80" and len(isikukood) < 11:
                        ws.Cells(r, isikukoodCol).Interior.Color = red
                        ws.Cells(r, kontaktitüüpCol).Interior.Color = red





            if not kontaktitüüp:
                ws.Cells(r, kontaktitüüpCol).Interior.Color = red
            elif kontaktitüüp.strip() not in ["80", "81"]:
                ws.Cells(r, kontaktitüüpCol).Interior.Color = red


            if kuupäev == "" or not on_õige_kuupäeva_formaat(kuupäev):
                ws.Cells(r, kuupäevCol).Interior.Color = red

            if dokumendiNr and dokumendiNr.upper() not in ["N/A", "XXX"]:
                if not on_õige_dokumendi_number(dokumendiNr):
                    ws.Cells(r, dokumendiNrCol).Interior.Color = red

            if pangakonto and pangakonto.upper() not in ["N/A", "XXX"]:
                if not on_õige_pangakonto_formaat(pangakonto):
                    ws.Cells(r, pangakontoCol).Interior.Color = red

            nimi_osad = nimi.split()
            if len(nimi_osad) < 2:
                ws.Cells(r, nimiCol).Interior.Color = red
                continue

            if len(nimi_osad) == 2 and nimi_osad[0].strip().upper() == nimi_osad[1].strip().upper():
                ws.Cells(r, nimiCol).Interior.Color = orange

            if not on_õige_nimi_formaat(nimi):
                ws.Cells(r, nimiCol).Interior.Color = red

            nimi_upper = nimi.upper()
            if any(re.search(r'\b' + re.escape(s) + r'\b', nimi_upper) for s in ettevottesonad):
                continue  # Äriline nimi

            if len(nimi_osad) == 2 and levenshtein(nimi_osad[0].upper(), nimi_osad[1].upper()) == 1:
                ws.Cells(r, nimiCol).Interior.Color = red

            for s in tahedsonad:
                if f" {s} " in f" {nimi_upper} ":
                    ws.Cells(r, nimiCol).Interior.Color = red
                    break

        wb.Save()
        wb.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()
        window.after(0, kontroll_lõpetatud)

    except Exception as e:
        window.after(0, lambda: töö_viga(str(e)))

# --- GUI --- #
def kontroll_lõpetatud():
    label_info.pack_forget()
    messagebox.showinfo("Valmis", "Kontroll lõpetatud.")
    window.destroy()

def töö_viga(veateade):
    messagebox.showerror("Viga", f"Töötlemisel tekkis viga:\n{veateade}")

def käivita_töötlus():
    threading.Thread(target=töötlus).start()

window = tk.Tk()
window.title("Andmete kontroll")
label_info = tk.Label(window, text="Teen kontrolli, palun oota...")
label_info.pack(padx=60, pady=60)
window.after(100, käivita_töötlus)
window.mainloop()
