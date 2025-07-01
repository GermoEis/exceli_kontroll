import os
import subprocess
import sys
import tkinter as tk

def käivita_exe(nimi):
    if getattr(sys, 'frozen', False):
        # Kui exe failina, siis kasuta sys.executable kausta
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.abspath(os.path.dirname(__file__))
    tee = os.path.join(script_dir, nimi)
    print(f"Käivitame: {tee}")
    try:
        subprocess.Popen([tee])
    except Exception as e:
        print(f"Ei saanud käivitada {nimi}: {e}")

aken = tk.Tk()
aken.title("Kontroll")
aken.geometry("500x300")

tk.Label(aken, text="Vali toiming:", font=("Arial", 14)).pack(pady=10)

tk.Button(aken, text="Xml muutmine", command=lambda: käivita_exe("xml_muutmine.exe"), width=50).pack(pady=5)
tk.Button(aken, text="Excelite võrdlus/kontroll suure tabeliga", command=lambda: käivita_exe("excelite_võrdlus.exe"), width=50).pack(pady=5)
tk.Button(aken, text="Kuupäevade kontroll vastavalt nädalale", command=lambda: käivita_exe("kuupäeva_kontroll.exe"), width=50).pack(pady=5)
tk.Button(aken, text="Exceli kontroll", command=lambda: käivita_exe("Exceli_kontroll.exe"), width=50).pack(pady=5)

tk.Button(aken, text="Sulge", command=aken.destroy, width=25).pack(pady=20)
versioon = "v1.0"
tk.Label(aken, text=f"Versioon: {versioon}", font=("Arial", 8), fg="gray").place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-5)

aken.mainloop()
