import os
import pandas as pd
import psycopg2
from dotenv import load_dotenv
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import threading

# ---------- Tkinter GUI ----------

def tööta():
    try:
        # Lae .env fail
        dotenv_path = Path(__file__).resolve().parents[0] / "config" / ".env"
        load_dotenv(dotenv_path=dotenv_path)

        # Kontrolli vajalikud võtmed
        required_keys = ["DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD"]
        missing = [key for key in required_keys if not os.getenv(key)]
        if missing:
            raise RuntimeError(f"Puuduvad .env võtmed: {', '.join(missing)}")

        # PostgreSQL ühendus
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

        # Loe kaustanimi failist
        with open("temp_folder.txt", "r", encoding="utf-8") as f:
            folder = f.read().strip()

        # Exceli tee
        excel_path = Path("Z:/scan") / folder / "Book1.xlsx"
        if not excel_path.exists():
            raise FileNotFoundError(f"Ei leidnud Exceli faili: {excel_path}")

        # Loe Excel (ilma päiseta)
        df_excel = pd.read_excel(excel_path, header=None)
        if df_excel.shape[1] < 16:
            raise ValueError(f"Excelis on ainult {df_excel.shape[1]} veergu, oodatakse vähemalt 16.")

        # Nimeta veerud ümber (N = 13, O = 14, P = 15)
        df_excel = df_excel.rename(columns={
            13: "nimi",
            14: "isikukood",
            15: "tyyp"
        })
        df_excel["isikukood"] = df_excel["isikukood"].astype(str).str.strip()
        df_excel["nimi"] = df_excel["nimi"].astype(str).str.strip()
        df_excel["tyyp"] = df_excel["tyyp"].astype(str).str.strip()

        # Andmebaasi ühendus ja uuendamine
        with psycopg2.connect(conn_str) as conn:
            conn.autocommit = True
            with conn.cursor() as cur:
                # Lae olemasolevad kirjed
                cur.execute('SELECT isikukood, nimi, tyyp FROM public."Metadata"')
                db_rows = cur.fetchall()
                db_dict = {
                    (str(row[0]).strip(), str(row[1]).strip()): str(row[2]).strip() if row[2] else ""
                    for row in db_rows
                }
                lisatud = 0
                uuendatud = 0
                df_excel = df_excel.drop_duplicates(subset=["isikukood", "nimi"], keep="first")
                for _, row in df_excel.iterrows():
                    isikukood = row["isikukood"]
                    nimi = row["nimi"]
                    tyyp = row["tyyp"]

                    if not isikukood:
                        continue

                    if (isikukood, nimi) not in db_dict:
                        cur.execute(
                            '''
                            INSERT INTO public."Metadata" 
                            (isikukood, nimi, tyyp, created_at) 
                            VALUES (%s, %s, %s, CURRENT_TIMESTAMP)
                            ''',
                            (isikukood, nimi, tyyp)
                        )
                        lisatud += 1
                    elif db_dict[(isikukood, nimi)] != tyyp:
                        cur.execute(
                            '''
                            UPDATE public."Metadata" 
                            SET nimi = %s, tyyp = %s, updated_at = CURRENT_TIMESTAMP 
                            WHERE isikukood = %s
                            ''',
                            (nimi, tyyp, isikukood)
                        )
                        uuendatud += 1

        # Edukas lõpp
        root.after(0, lambda: töö_valmis(lisatud, uuendatud))

    except Exception as e:
        root.after(0, lambda: töö_viga(str(e)))


def töö_valmis(lisatud, uuendatud):
    laadimisaken.destroy()
    messagebox.showinfo("Valmis", f"✅ Andmed uuendatud.\nLisatud: {lisatud}\nUuendatud: {uuendatud}")
    root.destroy()

def töö_viga(veateade):
    laadimisaken.destroy()
    messagebox.showerror("Viga", f"❌ Tekkis viga:\n{veateade}")
    root.destroy()


# ---------- GUI käivitamine ----------
root = tk.Tk()
root.withdraw()  # Peida peamine aken

laadimisaken = tk.Toplevel()
laadimisaken.title("Töötleb...")
laadimisaken.geometry("300x100")
laadimisaken.resizable(False, False)
tk.Label(laadimisaken, text="Tegutsen, palun oota...", font=("Arial", 12)).pack(pady=30)

# Alusta töötlust eraldi lõimes
threading.Thread(target=tööta, daemon=True).start()

root.mainloop()
