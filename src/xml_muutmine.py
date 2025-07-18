import os
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import simpledialog, messagebox
import win32com.client as win32

# Kasutajalt sisendi küsimine GUI kaudu
root = tk.Tk()
root.withdraw()
input_folder = simpledialog.askstring("Kausta nimi", "Sisesta kausta nimi (nt Tulemus_W2025_22_3):")
with open("temp_folder.txt", "w", encoding="utf-8") as f:
    f.write(input_folder)

if not input_folder:
    print("❌ Kausta nime ei sisestatud, katkestan.")
    exit()

base_path = r"Z:\\scan"

# --- CONFIGURE ---
folder_name = os.path.join(base_path, input_folder)
xml_file_name = "metadata.xml"
modified_xml_name = "metadata_modified.xml"
excel_file_name = "Book1.xlsx"
# ------------------

xml_path = os.path.join(folder_name, xml_file_name)
modified_xml_path = os.path.join(folder_name, modified_xml_name)
excel_path = os.path.join(folder_name, excel_file_name)

prioritized_fields = [
    "xxProductNameEnglish",
    "xxCustomerID",
    "xxComplTypeOfContact",
    "xxValidFrom",
    "xxDocumentNumber",
    "xxSendersReqNr"
]

# 1. Muuda XML-i järjekord

# Lae algne XML
tree = ET.parse(xml_path)
root_xml = tree.getroot()

for metadata in root_xml.findall("Metadata"):
    all_elements = list(metadata)
    prioritized = []
    other = []

    for elem in all_elements:
        if elem.tag in prioritized_fields:
            prioritized.append(elem)
        else:
            other.append(elem)

    prioritized_sorted = []
    for field in prioritized_fields:
        for elem in prioritized:
            if elem.tag == field:
                prioritized_sorted.append(elem)
                break

    first_13_others = other[:13]
    remaining_others = other[13:]

    metadata.clear()
    for elem in first_13_others + prioritized_sorted + remaining_others:
        metadata.append(elem)

# Salvesta muudetud XML
ET.ElementTree(root_xml).write(modified_xml_path, encoding="utf-8", xml_declaration=True)
print(f"✅ Muudetud XML salvestatud: {modified_xml_path}")

# 2. Ava Excel ja impordi XML (nii et saab hiljem Save As → XML Data)
excel = win32.Dispatch('Excel.Application')
# excel.Visible = False  # Pane True kui tahad akent näha

wb = excel.Workbooks.Add()

try:
    wb.XmlImport(modified_xml_path, ImportMap=None, Overwrite=True, Destination=wb.Sheets(1).Cells(1,1))
except Exception as e:
    print(f"❌ XML import ebaõnnestus: {e}")
    wb.Close(False)
    excel.Quit()
    exit()

wb.SaveAs(excel_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook (xlsx)
wb.Close(False)
excel.Quit()

print(f"📊 Excel salvestatud: {excel_path}")

# Teavita kasutajat
tk.Tk().withdraw()
messagebox.showinfo("Teade", "Muudetud XML imporditi Excelisse ja salvestati.")
