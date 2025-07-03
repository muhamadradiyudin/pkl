import pandas as pd
import tkinter as tk
import threading
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
from datetime import datetime

# Fungsi untuk menghitung usia
def calculate_age(birth_date):
    today = datetime.today()
    return today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))

# Fungsi untuk membaca data dari file Excel
def load_excel_data(file_path):
    file_extension = file_path.split('.')[-1]
    if file_extension in ['xls', 'xlsx', 'xlsm', 'xlsb']:
        data = pd.read_excel(file_path)
    else:
        raise ValueError("Format file tidak didukung.")

    data["SERTIFIKASI"].fillna("Belum", inplace=True)
    if "INPASSING" not in data.columns:
        data["INPASSING"] = "Belum"
    else:
        data["INPASSING"] = data["INPASSING"].replace({"Y": "Sudah", "N": "Belum"})
        data["INPASSING"].fillna("Belum", inplace=True)

    if "TANGGAL LAHIR" in data.columns:
        data["TANGGAL LAHIR"] = pd.to_datetime(data["TANGGAL LAHIR"], errors='coerce')
        data["USIA"] = data["TANGGAL LAHIR"].apply(lambda x: calculate_age(x) if not pd.isnull(x) else None)
        data["PENSIUN"] = data["USIA"].apply(lambda x: "Sudah" if x is not None and x >= 60 else "Belum")
    else:
        data["PENSIUN"] = "Data Tanggal Lahir Tidak Ada"

    return data

# Fungsi untuk menampilkan detail baris saat diklik
def show_detail(event):
    selected_item = tree.focus()
    if selected_item:
        values = tree.item(selected_item, "values")
        column_names = tree["columns"]
        detail = "\n".join(f"{col}: {val}" for col, val in zip(column_names, values))
        messagebox.showinfo("Detail Data", detail)

# Fungsi filter data
def filter_data(event=None):
    jenjang = jenjang_var.get()
    sertifikasi = sertifikasi_var.get()
    status_pegawai = status_var.get()
    inpassing = inpassing_var.get()
    jenis_kelamin = jenisKelamin_var.get()
    pensiun = pensiun_var.get()
    usia_filter = usia_var.get()

    filtered_data = data.copy()

    if jenjang != "Semua":
        filtered_data = filtered_data[filtered_data["JENJANG SEKOLAH"] == jenjang]
    if sertifikasi != "Semua":
        filtered_data = filtered_data[filtered_data["SERTIFIKASI"] == sertifikasi]
    if status_pegawai != "Semua":
        filtered_data = filtered_data[filtered_data["STATUS PEGAWAI"] == status_pegawai]
    if inpassing != "Semua":
        filtered_data = filtered_data[filtered_data["INPASSING"] == inpassing]
    if jenis_kelamin != "Semua":
        filtered_data = filtered_data[filtered_data["JENIS KELAMIN"] == jenis_kelamin]
    if pensiun != "Semua":
        filtered_data = filtered_data[filtered_data["PENSIUN"] == pensiun]
    if usia_filter != "Semua":
        batas = int(usia_filter.split(" ")[-1])
        filtered_data = filtered_data[filtered_data["USIA"] > batas]

    display_data(filtered_data)

# Fungsi untuk menampilkan data
def display_data(data):
    tree.delete(*tree.get_children())
    tree["columns"] = list(data.columns)
    tree["show"] = "headings"
    for col in data.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))

# Inisialisasi GUI
root = tk.Tk()
root.title("Aplikasi GPAIDIA")

filter_frame = tk.Frame(root)
filter_frame.pack(pady=10)

# Filter Dropdown
jenjang_var = tk.StringVar(value="Semua")
ttk.Label(filter_frame, text="Jenjang Sekolah:").grid(row=0, column=0)
ttk.Combobox(filter_frame, textvariable=jenjang_var, values=["Semua", "SD", "SMP", "SMA", "SMK"]).grid(row=0, column=1)

sertifikasi_var = tk.StringVar(value="Semua")
ttk.Label(filter_frame, text="Sertifikasi:").grid(row=0, column=2)
ttk.Combobox(filter_frame, textvariable=sertifikasi_var, values=["Semua", "Sudah", "Belum"]).grid(row=0, column=3)

status_var = tk.StringVar(value="Semua")
ttk.Label(filter_frame, text="Status Pegawai:").grid(row=1, column=0)
ttk.Combobox(filter_frame, textvariable=status_var, values=["Semua", "PNS", "NON PNS", "PPPK"]).grid(row=1, column=1)

inpassing_var = tk.StringVar(value="Semua")
ttk.Label(filter_frame, text="Inpassing:").grid(row=1, column=2)
ttk.Combobox(filter_frame, textvariable=inpassing_var, values=["Semua", "Sudah", "Belum"]).grid(row=1, column=3)

jenisKelamin_var = tk.StringVar(value="Semua")
ttk.Label(filter_frame, text="Jenis Kelamin:").grid(row=2, column=0)
ttk.Combobox(filter_frame, textvariable=jenisKelamin_var, values=["Semua", "L", "P"]).grid(row=2, column=1)

pensiun_var = tk.StringVar(value="Semua")
ttk.Label(filter_frame, text="Pensiun:").grid(row=2, column=2)
ttk.Combobox(filter_frame, textvariable=pensiun_var, values=["Semua", "Sudah", "Belum"]).grid(row=2, column=3)

usia_var = tk.StringVar(value="Semua")
ttk.Label(filter_frame, text="Usia:").grid(row=3, column=0)
ttk.Combobox(filter_frame, textvariable=usia_var, values=["Semua", "Lebih dari 20", "Lebih dari 30", "Lebih dari 40", "Lebih dari 50"]).grid(row=3, column=1)

# Treeview
tree = ttk.Treeview(root)
tree.pack(fill="both", expand=True)
tree.bind("<Double-1>", show_detail)
tree.bind("<<TreeviewSelect>>", show_detail_popup)


# Load button
load_button = ttk.Button(root, text="Load Excel", command=lambda: [
    setattr(globals(), "data", load_excel_data(filedialog.askopenfilename())),
    display_data(data)
])
load_button.pack(pady=10)

# Bind filter
for var in [jenjang_var, sertifikasi_var, status_var, inpassing_var, jenisKelamin_var, pensiun_var, usia_var]:
    var.trace("w", lambda *args: filter_data())

root.mainloop()
