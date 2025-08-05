import pandas as pd
import os
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox


def choose_file(title):
    return filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx")])


def main():
    # UI – failide valimine
    root = tk.Tk()
    root.withdraw()  # Peida põhialusaken

    messagebox.showinfo("Faili valimine", "Palun vali Book1 Excel fail.")
    file1_path = choose_file("Vali Book1 Excel fail")

    if not file1_path:
        messagebox.showerror("Viga", "Book1 faili ei valitud!")
        return

    messagebox.showinfo("Faili valimine", "Palun vali SMART-ID Excel fail.")
    file2_path = choose_file("Vali SMART-ID Excel fail")

    if not file2_path:
        messagebox.showerror("Viga", "SMART-ID faili ei valitud!")
        return

    # Lae andmed
    df1 = pd.read_excel(file1_path)
    df2 = pd.read_excel(file2_path)

    df1['SMART_ID_ACCOUNT_NUMBER'] = df1['SMART_ID_ACCOUNT_NUMBER'].astype(str).str.strip()
    df2['SMART_ID_ACCOUNT_NUMBER'] = df2['SMART_ID_ACCOUNT_NUMBER'].astype(str).str.strip()

    df1_filtered = df1[df1['xxRegNr'] == 'B06.01-200-05'].copy()
    df2_indexed = df2.set_index('SMART_ID_ACCOUNT_NUMBER')

    logs_missing_account = []
    logs_missing_in_first_table = []
    logs_changes = []

    df1_accounts = set(df1_filtered['SMART_ID_ACCOUNT_NUMBER'].dropna())
    df2_accounts = set(df2['SMART_ID_ACCOUNT_NUMBER'].dropna())

    extra_accounts = df2_accounts - df1_accounts
    for acc in sorted(extra_accounts):
        logs_missing_account.append(f"SMART_ID_ACCOUNT_NUMBER {acc} on kontrolltabelis, aga puudub meie tabelis")

    missing_in_control_table = df1_accounts - df2_accounts
    for acc in sorted(missing_in_control_table):
        logs_missing_account.append(f"SMART_ID_ACCOUNT_NUMBER {acc} on meie tabelis, aga puudub kontrolltabelis")

    # Excel automation
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False

    wb = excel.Workbooks.Open(file1_path)
    ws = wb.Sheets(1)

    headers = [ws.Cells(1, col).Value for col in range(1, ws.UsedRange.Columns.Count + 1)]
    col_idx = {
        "SMART_ID_ACCOUNT_NUMBER": headers.index("SMART_ID_ACCOUNT_NUMBER") + 1,
        "xxProductNameEnglish": headers.index("xxProductNameEnglish") + 1,
        "xxValidFrom": headers.index("xxValidFrom") + 1,
        "xxRegNr": headers.index("xxRegNr") + 1,
        "xComments": headers.index("xComments") + 1,
        "xxCustomerID": headers.index("xxCustomerID") + 1,
    }

    last_row = ws.UsedRange.Rows.Count

    for row in range(2, last_row + 1):
        regnr_val = str(ws.Cells(row, col_idx["xxRegNr"]).Value).strip()
        if regnr_val != "B06.01-200-05":
            continue

        account_cell = ws.Cells(row, col_idx["SMART_ID_ACCOUNT_NUMBER"])
        account_number = str(account_cell.Value).strip() if account_cell.Value else ""
        comment_cell = ws.Cells(row, col_idx["xComments"])

        if not account_number or account_number.lower() in ['nan', 'none', '']:
            logs_missing_account.append(f"Rida {row}: Puudub SMART_ID_ACCOUNT_NUMBER")
            continue

        if account_number not in df2_indexed.index:
            log_entry_general = f"SMART_ID_ACCOUNT_NUMBER {account_number} on meie tabelis, aga puudub kontrolltabelis"
            log_entry_row = f"Rida {row}: SMART_ID_ACCOUNT_NUMBER {account_number} puudub kontroll tabelis"

            if log_entry_general not in logs_missing_account:
                logs_missing_account.append(log_entry_row)

            comment_cell.Value = "puudub kontroll tabelis"
            continue

        match_row = df2_indexed.loc[account_number]
        if isinstance(match_row, pd.DataFrame):
            logs_changes.append(f"Rida {row}: mitu vastet kontrolltabelis kontole {account_number}, kasutati esimest")
            match_row = match_row.iloc[0]

        changed = False

        name_cell = ws.Cells(row, col_idx["xxProductNameEnglish"])
        if not name_cell.Value:
            logs_missing_in_first_table.append(f"Rida {row}: Puudub xxProductNameEnglish väärtus esimeses tabelis")

        name2 = match_row['TNIM']
        if name_cell.Value != name2:
            logs_changes.append(f"Rida {row}: Nimi asendatud '{name_cell.Value}' -> '{name2}'")
            name_cell.Value = name2
            changed = True

        date_cell = ws.Cells(row, col_idx["xxValidFrom"])
        if not date_cell.Value:
            logs_missing_in_first_table.append(f"Rida {row}: Puudub xxValidFrom väärtus esimeses tabelis")

        date2 = match_row['LOG_DTIME']
        match_date = pd.to_datetime(date2, errors='coerce')
        if pd.notna(match_date):
            logs_changes.append(f"Rida {row}: Kuupäev asendatud '{date_cell.Value}' -> '{match_date}'")
            date_cell.Value = match_date.strftime("%Y-%m-%d")
            changed = True
        else:
            logs_changes.append(f"Rida {row}: Smart_ID kuupäev ei ole kehtiv: '{date2}'")

        id_cell = ws.Cells(row, col_idx["xxCustomerID"])
        id2 = str(match_row['REGISTRATION_NUMBER']).strip() if 'REGISTRATION_NUMBER' in match_row else ""
        if id_cell.Value != id2:
            logs_changes.append(f"Rida {row}: Isikukood asendatud '{id_cell.Value}' -> '{id2}'")
            id_cell.Value = id2
            changed = True

        if changed:
            comment_cell.Value = "kontrollitud"

    wb.Save()
    wb.Close()
    excel.Quit()

    log_file_path = os.path.join(os.path.dirname(file1_path), "muudatused_logis.txt")
    with open(log_file_path, "w", encoding="utf-8") as f:
        for log in logs_missing_account + logs_missing_in_first_table + logs_changes:
            f.write(log + "\n")

    messagebox.showinfo("Valmis", f"Muudatused salvestatud!\n\nFail:\n{file1_path}\n\nLogi:\n{log_file_path}")


if __name__ == "__main__":
    main()
