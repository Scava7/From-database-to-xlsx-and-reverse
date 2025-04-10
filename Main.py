import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def scegli_file(titolo, tipi_file):
    file_path = filedialog.askopenfilename(title=titolo, filetypes=tipi_file)
    return file_path

def scegli_salvataggio(titolo, tipo_file):
    file_path = filedialog.asksaveasfilename(title=titolo, defaultextension=tipo_file[1], filetypes=[tipo_file])
    return file_path

def esporta_sqlite_in_excel():
    db_path = scegli_file("Scegli il database SQLite", [("Database SQLite", "*.db *.sqlite")])
    if not db_path:
        return

    xlsx_path = scegli_salvataggio("Salva il file Excel", ("Excel Workbook", ".xlsx"))
    if not xlsx_path:
        return

    conn = sqlite3.connect(db_path)
    xls_writer = pd.ExcelWriter(xlsx_path, engine="openpyxl")

    query = "SELECT name FROM sqlite_master WHERE type='table';"
    tables = pd.read_sql(query, conn)

    for table_name in tables['name']:
        df = pd.read_sql(f"SELECT * FROM '{table_name}'", conn)
        df.to_excel(xls_writer, sheet_name=table_name, index=False)

    xls_writer.close()
    conn.close()
    messagebox.showinfo("Completato", f"Esportazione completata in:\n{xlsx_path}")

def importa_excel_in_sqlite():
    xlsx_path = scegli_file("Scegli il file Excel", [("Excel Workbook", "*.xlsx")])
    if not xlsx_path:
        return

    db_path = scegli_salvataggio("Salva il database SQLite", ("Database SQLite", ".sqlite"))
    if not db_path:
        return

    xls = pd.ExcelFile(xlsx_path)
    conn = sqlite3.connect(db_path)

    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        df.to_sql(sheet_name, conn, index=False, if_exists="replace")

    conn.close()
    messagebox.showinfo("Completato", f"Importazione completata in:\n{db_path}")

def main():
    root = tk.Tk()
    root.title("Convertitore SQLite/Excel")

    btn_esporta = tk.Button(root, text="Esporta da SQLite a Excel", command=esporta_sqlite_in_excel)
    btn_esporta.pack(pady=10)

    btn_importa = tk.Button(root, text="Importa da Excel a SQLite", command=importa_excel_in_sqlite)
    btn_importa.pack(pady=10)

    btn_esci = tk.Button(root, text="Esci", command=root.quit)
    btn_esci.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
