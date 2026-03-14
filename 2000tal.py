import tkinter as tk
from tkinter import scrolledtext, messagebox
import os
import tempfile
import subprocess
from openpyxl import Workbook

def export_and_open_temp():
    # Läs in värden (en per rad), rensa tomrader
    lines = [r.strip() for r in input_text.get("1.0", tk.END).splitlines() if r.strip()]
    if not lines:
        messagebox.showwarning("Inget att exportera", "Klistra in dina värden först.")
        return

    # Chunk-storlek (2000 som standard)
    try:
        chunk_size = int(chunk_var.get())
        if chunk_size <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Felaktigt värde", "Antalet per kolumn måste vara ett heltal > 0.")
        return

    # Dela upp i kolumner
    chunks = [lines[i:i+chunk_size] for i in range(0, len(lines), chunk_size)]

    # Skapa tillfällig fil
    fd, filepath = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)  # stäng filhandtaget, openpyxl behöver filen fri

    # Skriv till Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Delade värden"

    for col_idx, chunk in enumerate(chunks, start=1):
        for row_idx, val in enumerate(chunk, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=str(val))
            cell.number_format = "@"  # textformat

    wb.save(filepath)

    # Öppna i Excel direkt
    try:
        os.startfile(filepath)  # Windows
    except AttributeError:
        subprocess.call(["open" if os.name == "posix" else "xdg-open", filepath])

    messagebox.showinfo("Klart!", "Excel-filen öppnades direkt. Spara den i Excel om du vill behålla den.")

# --- UI ---
root = tk.Tk()
root.title("Exportera direkt till Excel")

top = tk.Frame(root)
top.pack(fill="both", expand=True, padx=10, pady=(10,5))

lbl = tk.Label(top, text="Klistra in dina värden (en per rad):")
lbl.pack(anchor="w")

input_text = scrolledtext.ScrolledText(top, width=90, height=18)
input_text.pack(fill="both", expand=True)

bottom = tk.Frame(root)
bottom.pack(fill="x", padx=10, pady=10)

tk.Label(bottom, text="Antal rader per kolumn:").pack(side="left")
chunk_var = tk.StringVar(value="2000")
tk.Entry(bottom, textvariable=chunk_var, width=8).pack(side="left", padx=(5,15))

tk.Button(bottom, text="Öppna i Excel direkt", command=export_and_open_temp).pack(side="left")

root.mainloop()
