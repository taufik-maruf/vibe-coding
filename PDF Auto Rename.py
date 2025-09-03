import os
import shutil
import random
import openpyxl
import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox

# 🔍 Ekstraksi teks dari PDF berdasarkan kata kunci
def extract_specific_text_from_pdf(file_path, start_keyword, end_keyword):
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            page_text = page.extract_text()
            start = page_text.find(start_keyword)
            end = page_text.find(end_keyword)
            if start != -1 and end != -1:
                text += page_text[start+len(start_keyword):end]
    return text

# 🔄 Rename dan pindahkan file
def rename_and_move_file(old_path, new_dir, new_name):
    if not os.path.exists(new_dir):
        os.makedirs(new_dir)
    base, ext = os.path.splitext(old_path)
    new_path = os.path.join(new_dir, 'NEW_' + new_name + ext)
    shutil.move(old_path, new_path)

# 📊 Tulis hasil ke Excel
def write_to_excel(file_path, data):
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, row in enumerate(data, start=1):
        for j, value in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=value)
    wb.save(file_path)

# 🚀 Proses semua PDF
def process_pdfs(pdf_dir, output_excel, output_dir, start_keyword, end_keyword):
    data = [['File Name', 'Extracted Text']]
    for file_name in os.listdir(pdf_dir):
        if file_name.endswith('.pdf'):
            file_path = os.path.join(pdf_dir, file_name)
            text = extract_specific_text_from_pdf(file_path, start_keyword, end_keyword)
            data.append([file_name, text])
            new_name = text.replace('\\', '').replace('/', '').replace('-', '').replace(':', '').replace('(', '').replace(')', '').replace('\n', '').replace('Place of Birth', '').replace('Nomor Induk Mahasiswa', '').replace('Student Number', '')
            rename_and_move_file(file_path, output_dir, new_name)
    write_to_excel(output_excel, data)

# 🔍 Preview satu file acak
def preview_random_file():
    folder = entry_sumber.get()
    awal = entry_awal.get()
    akhir = entry_akhir.get()

    if not all([folder, awal, akhir]):
        messagebox.showerror("Error", "Folder sumber dan kata kunci harus diisi.")
        return

    try:
        pdf_files = [f for f in os.listdir(folder) if f.endswith('.pdf')]
        if not pdf_files:
            messagebox.showwarning("Kosong", "Tidak ada file PDF di folder sumber.")
            return

        random_file = random.choice(pdf_files)
        file_path = os.path.join(folder, random_file)
        text = extract_specific_text_from_pdf(file_path, awal, akhir)
        new_name = text.replace('\\', '').replace('/', '').replace('-', '').replace(':', '').replace('(', '').replace(')', '').replace('\n', '').replace('Place of Birth', '').replace('Nomor Induk Mahasiswa', '').replace('Student Number', '')

        messagebox.showinfo("Preview",
            f"📄 File: {random_file}\n🔍 Hasil Ekstraksi: {text}\n📁 Nama Baru: NEW_{new_name}.pdf")
    except Exception as e:
        messagebox.showerror("Gagal Preview", f"Terjadi kesalahan:\n{e}")

# 🖥️ GUI Setup
def pilih_folder(entry_field):
    folder = filedialog.askdirectory()
    if folder:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, folder)

def pilih_file_excel(entry_field):
    file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, file)

def mulai_proses():
    sumber = entry_sumber.get()
    tujuan = entry_tujuan.get()
    excel_path = entry_excel.get()
    awal = entry_awal.get()
    akhir = entry_akhir.get()

    if not all([sumber, tujuan, excel_path, awal, akhir]):
        messagebox.showerror("Error", "Semua field harus diisi.")
        return

    try:
        process_pdfs(sumber, excel_path, tujuan, awal, akhir)
        messagebox.showinfo("Sukses", "Proses selesai. File sudah dipindahkan dan Excel dibuat.")
    except Exception as e:
        messagebox.showerror("Gagal", f"Terjadi kesalahan:\n{e}")

# 🧱 Build GUI
root = tk.Tk()
root.title("PDF Renamer NEW")

tk.Label(root, text="📁 Folder Sumber PDF").grid(row=0, column=0, sticky="w")
entry_sumber = tk.Entry(root, width=50)
entry_sumber.grid(row=0, column=1)
tk.Button(root, text="Pilih", command=lambda: pilih_folder(entry_sumber)).grid(row=0, column=2)

tk.Label(root, text="📁 Folder Tujuan PDF").grid(row=1, column=0, sticky="w")
entry_tujuan = tk.Entry(root, width=50)
entry_tujuan.grid(row=1, column=1)
tk.Button(root, text="Pilih", command=lambda: pilih_folder(entry_tujuan)).grid(row=1, column=2)

tk.Label(root, text="📄 Lokasi File Excel Output").grid(row=2, column=0, sticky="w")
entry_excel = tk.Entry(root, width=50)
entry_excel.grid(row=2, column=1)
tk.Button(root, text="Pilih", command=lambda: pilih_file_excel(entry_excel)).grid(row=2, column=2)

tk.Label(root, text="🔍 Kata Kunci Awal").grid(row=3, column=0, sticky="w")
entry_awal = tk.Entry(root, width=50)
entry_awal.grid(row=3, column=1)

tk.Label(root, text="🔍 Kata Kunci Akhir").grid(row=4, column=0, sticky="w")
entry_akhir = tk.Entry(root, width=50)
entry_akhir.grid(row=4, column=1)

tk.Button(root, text="🚀 Mulai Proses", command=mulai_proses).grid(row=5, column=1, pady=10)
tk.Button(root, text="🔍 Preview Salah Satu File", command=preview_random_file).grid(row=6, column=1, pady=5)

root.mainloop()
