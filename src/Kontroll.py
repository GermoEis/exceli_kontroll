import os
import subprocess
import sys
import tkinter as tk
import urllib.request
import zipfile
import io
import shutil
import tkinter.messagebox as msgbox

VERSIOONI_URL = "https://raw.githubusercontent.com/GermoEis/exceli_kontroll/main/versioon.txt"
ZIP_URL = "https://github.com/GermoEis/exceli_kontroll/archive/refs/heads/main.zip"

def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.abspath(os.path.dirname(__file__))

def loe_kohalik_versioon():
    try:
        failitee = os.path.join(get_script_dir(), "config", "versioon.txt")
        with open(failitee, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return "versioon puudub"

def versioon_numbriks(v):
    v = v.lstrip("vV")
    osad = v.split(".")
    nums = []
    for x in osad:
        try:
            nums.append(int(x))
        except:
            nums.append(0)
    return tuple(nums)

def kas_uuendus_on():
    try:
        uus_versioon = urllib.request.urlopen(VERSIOONI_URL).read().decode().strip()
        lokaalne_versioon = loe_kohalik_versioon()
        uus_num = versioon_numbriks(uus_versioon)
        lokaalne_num = versioon_numbriks(lokaalne_versioon)
        return uus_num > lokaalne_num
    except Exception as e:
        print(f"Versiooni kontrolli viga: {e}")
        return False

def käivita_exe(nimi):
    tee = os.path.join(get_script_dir(), nimi)
    try:
        subprocess.Popen([tee])
    except Exception as e:
        print(f"Ei saanud käivitada {nimi}: {e}")

# --- GUI ---
aken = tk.Tk()
aken.title("Kontroll")
aken.geometry("500x430")

tk.Label(aken, text="Vali toiming:", font=("Arial", 14)).pack(pady=10)

tk.Button(aken, text="Xml muutmine", command=lambda: käivita_exe("xml_muutmine.exe"), width=50).pack(pady=5)
tk.Button(aken, text="Excelite võrdlus/kontroll suure tabeliga", command=lambda: käivita_exe("excelite_võrdlus.exe"), width=50).pack(pady=5)
tk.Button(aken, text="Kuupäevade kontroll vastavalt nädalale", command=lambda: käivita_exe("kuupäeva_kontroll.exe"), width=50).pack(pady=5)
tk.Button(aken, text="Exceli kontroll", command=lambda: käivita_exe("Exceli_kontroll.exe"), width=50).pack(pady=5)

# Versiooni silt alla nurka
versioon_sisu = loe_kohalik_versioon()
versiooni_silt = tk.Label(aken, text=f"Versioon: {versioon_sisu}", font=("Arial", 8), fg="gray")
versiooni_silt.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-5)

# Uuenda nupp + punane teade
uuenda_frame = tk.Frame(aken)
uuenda_frame.pack(pady=10)

tk.Button(uuenda_frame, text="Uuenda", command=lambda: uuenda(), width=25).pack(side="left")

uuendus_silt = tk.Label(uuenda_frame, text="Uuendus saadaval!", fg="red", font=("Arial", 10, "bold"))
uuendus_silt.pack_forget()

def kontrolli_uuendus_ja_näita():
    if kas_uuendus_on():
        uuendus_silt.pack(side="left", padx=10)
    else:
        uuendus_silt.pack_forget()

def uuenda():
    try:
        uus_versioon = urllib.request.urlopen(VERSIOONI_URL).read().decode().strip()
        lokaalne_versioon = loe_kohalik_versioon()
        if versioon_numbriks(uus_versioon) > versioon_numbriks(lokaalne_versioon):
            msgbox.showinfo("Uuendus", f"Leiti uuem versioon ({uus_versioon}). Alustame uuendamist...")

            response = urllib.request.urlopen(ZIP_URL)
            z = zipfile.ZipFile(io.BytesIO(response.read()))
            temp_dir = os.path.join(get_script_dir(), "temp_update")
            z.extractall(temp_dir)

            extracted_root = os.path.join(temp_dir, os.listdir(temp_dir)[0])

            for nimi in os.listdir(extracted_root):
                lähte = os.path.join(extracted_root, nimi)
                siht = os.path.join(get_script_dir(), nimi)
                if os.path.isfile(lähte) and not lähte.endswith(".gitignore"):
                    shutil.copy2(lähte, siht)


            shutil.rmtree(temp_dir)

            versioonitee = os.path.join(get_script_dir(), "versioon.txt")
            with open(versioonitee, "w", encoding="utf-8") as f:
                f.write(uus_versioon)

            versiooni_silt.config(text=f"Versioon: {uus_versioon}")
            uuendus_silt.pack_forget()
            msgbox.showinfo("Uuendus tehtud", "Rakendus on edukalt uuendatud.")
        else:
            msgbox.showinfo("Uuendus", "Sul on juba kõige uuem versioon.")
    except Exception as e:
        msgbox.showerror("Viga", f"Uuendamine ebaõnnestus:\n{e}")

# Kontrollime 100ms pärast GUI starti, et oleks kindlasti valmis
aken.after(100, kontrolli_uuendus_ja_näita)

tk.Button(aken, text="Sulge", command=aken.destroy, width=25).pack(pady=10)

aken.mainloop()