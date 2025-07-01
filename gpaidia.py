import pandas as pd
import tkinter as tk
import threading
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
from datetime import datetime
# from PIL import Image, ImageTk

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
    
    # Ganti nilai kosong di kolom SERTIFIKASI dan INPASSING dengan "Belum"
    data["SERTIFIKASI"].fillna("Belum", inplace=True)
    if "INPASSING" not in data.columns:
        data["INPASSING"] = "Belum"
    else:
        # Konversi nilai "Y" menjadi "Sudah" di kolom INPASSING
        data["INPASSING"] = data["INPASSING"].replace({"Y": "Sudah", "N": "Belum"})
        data["INPASSING"].fillna("Belum", inplace=True)

    # Tambahkan kolom PENSIUN
    if "TANGGAL LAHIR" in data.columns:
        data["TANGGAL LAHIR"] = pd.to_datetime(data["TANGGAL LAHIR"], errors='coerce')
        data["USIA"] = data["TANGGAL LAHIR"].apply(lambda x: calculate_age(x) if not pd.isnull(x) else None)
        data["PENSIUN"] = data["USIA"].apply(lambda x: "Sudah" if x is not None and x >= 60 else "Belum")
    else:
        data["PENSIUN"] = "Data Tanggal Lahir Tidak Ada"

    return data

# Fungsi untuk menampilkan data di Treeview dan mengatur lebar kolom
def display_data(data, refresh_needed=True):
    tree.delete(*tree.get_children())  # Menghapus data yang ada di treeview

    # Menambahkan kolom ke treeview
    tree["columns"] = list(data.columns)
    tree["show"] = "headings"
    
    for col in data.columns:
        tree.heading(col, text=col)
        max_len = max(data[col].astype(str).apply(len).max(), len(col)) + 20
        tree.column(col, width=max_len, minwidth=200, anchor='center')
    
    # Menambahkan data ke treeview
    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))
    
    # Hitung jumlah PNS, NON PNS, dan PPPK
    count_pns = len(data[data["STATUS PEGAWAI"] == "PNS"])
    count_non_pns = len(data[data["STATUS PEGAWAI"] == "NON PNS"])
    count_pppk = len(data[data["STATUS PEGAWAI"] == "PPPK"])
    
    # Update info_label untuk menampilkan jumlah PNS, NON PNS, dan PPPK
    info_label.config(text=f"Jumlah PNS: {count_pns}, Jumlah NON PNS: {count_non_pns}, Jumlah PPPK: {count_pppk}")



# Variabel global untuk menyimpan ID timer
debounce_id = None

# Fungsi untuk menangani pencarian dengan debouncing
def on_search_change(event=None):
    global debounce_id
    if debounce_id is not None:
        root.after_cancel(debounce_id)
    debounce_id = root.after(300, filter_data)  # Tunggu 300ms setelah berhenti mengetik


# Fungsi untuk menyaring data berdasarkan jenjang sekolah, sertifikasi, status pegawai, inpassing, jenis_kelamin, dan pensiun
def filter_data(event=None):
    jenjang = jenjang_var.get()
    sertifikasi = sertifikasi_var.get()
    status_pegawai = status_var.get()
    inpassing = inpassing_var.get()
    jenis_kelamin = jenisKelamin_var.get()
    pensiun = pensiun_var.get()
    
    search_keyword = search_var.get().lower()

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
    
    if search_keyword:
        filtered_data = filtered_data[
            filtered_data.apply(lambda row: search_keyword in row.astype(str).str.lower().to_string(), axis=1)
        ]
    
    display_data(filtered_data)
    highlight_search_result(search_keyword)


# Fungsi untuk menyoroti hasil pencarian dengan mengarahkan pointer ke kata yang tepat
def highlight_search_result(keyword):
    # Hapus semua tag yang ada
    for item in tree.get_children():
        tree.item(item, tags=())
    
    if keyword:
        for item in tree.get_children():
            values = tree.item(item, "values")
            # Menggunakan kondisi if untuk mencocokkan keyword
            if any(keyword in str(value).lower() for value in values):
                tree.item(item, tags=("highlight",))
    
    # Konfigurasi tag highlight
    tree.tag_configure("highlight", background="yellow", foreground="red")


# Fungsi untuk menyimpan data ke file Excel
def save_to_excel():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            # Mengambil data yang ditampilkan di Treeview
            data_to_export = []
            for item in tree.get_children():
                values = tree.item(item, "values")
                data_to_export.append(values)
            
            # Membuat DataFrame dari data yang akan diekspor
            df_to_export = pd.DataFrame(data_to_export, columns=data.columns)
            
            # Menyimpan ke file Excel
            df_to_export.to_excel(file_path, index=False)
            messagebox.showinfo("Info", "Data berhasil disimpan ke file Excel.")
        except Exception as e:
            messagebox.showerror("Error", f"Error saat menyimpan data: {str(e)}")

# Fungsi untuk membuka dialog file dan memuat data
def load_file():
    global data
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx *.xlsm *.xlsb")])
    if file_path:
        try:
            data = load_excel_data(file_path)
            display_data(data)
        except Exception as e:
            messagebox.showerror("Error", f"Error saat memuat data: {str(e)}")

# Fungsi untuk menyembunyikan kolom dan menyimpan lebar asli kolom
def hide_column(column):
    original_widths[column] = tree.column(column, option='width')  # Simpan lebar asli kolom
    tree.column(column, width=0, stretch=tk.NO)

# Fungsi untuk menampilkan kolom dan mengatur lebar kolom
def show_column(column):
    if column in original_widths:
        tree.column(column, width=original_widths[column], minwidth=200, anchor='center')
    else:
        tree.column(column, width=100, minwidth=200, anchor='center')

# Fungsi untuk menampilkan dialog Show Columns
def show_columns():
    def on_apply():
        selected_columns = [col for col, var in check_vars.items() if var.get()]

        # Menyembunyikan semua kolom terlebih dahulu
        for col in data.columns:
            hide_column(col)

        # Menampilkan kolom yang dipilih
        for col in selected_columns:
            show_column(col)

        show_columns_dialog.destroy()

    show_columns_dialog = tk.Toplevel(root)
    show_columns_dialog.title("Show Columns")

    tk.Label(show_columns_dialog, text="Select columns to show:").pack(pady=5)

    check_vars = {}
    columns_frame = tk.Frame(show_columns_dialog)
    columns_frame.pack(fill='both', expand=True)

    columns_canvas = tk.Canvas(columns_frame)
    columns_canvas.pack(side='left', fill='both', expand=True)

    scrollbar = ttk.Scrollbar(columns_frame, orient='vertical', command=columns_canvas.yview)
    scrollbar.pack(side='right', fill='y')

    columns_canvas.configure(yscrollcommand=scrollbar.set)
    columns_canvas.bind('<Configure>', lambda e: columns_canvas.configure(scrollregion=columns_canvas.bbox('all')))

    columns_inner_frame = tk.Frame(columns_canvas)
    columns_canvas.create_window((0, 0), window=columns_inner_frame, anchor='nw')
    
    # Menyiapkan checkbox untuk kolom
    for col in data.columns:
        var = tk.BooleanVar(value=tree.column(col, option='width') > 0)
        check_vars[col] = var
        checkbutton = tk.Checkbutton(columns_inner_frame, text=col, variable=var)
        checkbutton.pack(anchor='w')

    apply_button = ttk.Button(show_columns_dialog, text="Apply", command=on_apply)
    apply_button.pack(pady=5)

    # Menyimpan data check_vars dan show_columns_dialog agar tidak dihapus oleh garbage collector
    show_columns_dialog.check_vars = check_vars
    show_columns_dialog.apply_button = apply_button

# Fungsi untuk menyegarkan data dan memperbarui tampilan kolom
def refresh_data():
    # Reset kolom pencarian
    search_var.set("")  # Mengosongkan kolom pencarian

    # Reset semua filter ke "Semua"
    jenjang_var.set("Semua")
    sertifikasi_var.set("Semua")
    status_var.set("Semua")
    inpassing_var.set("Semua")
    jenisKelamin_var.set("Semua")
    pensiun_var.set("Semua")
    
    # Mengupdate data yang ditampilkan di treeview dengan data asli
    display_data(data, refresh_needed=False)

    # Reset filter data
    filter_data()

# Fungsi untuk memuat data secara asynchronous
def load_data_async():
    threading.Thread(target=load_data).start()

# Fungsi untuk memproses data (dummy function)
def load_data():
    # Misalnya proses loading data yang berat
    print("Memuat data...")

# Inisialisasi tkinter
root = tk.Tk()
root.title("GPAIDIA")

# Inisialisasi dictionary untuk menyimpan lebar asli kolom
original_widths = {}

# Tambahkan gaya ttk.Style
style = ttk.Style(root)
style.theme_use("clam")  
style.configure("TButton", font=("Arial", 10), padding=10)
style.configure("TLabel", font=("Arial", 10))
style.configure("TCombobox", font=("Arial", 10))
style.configure("Treeview.Heading", font=("Arial", 10, "bold"), foreground="black")
style.configure("Treeview", font=("Arial", 10), rowheight=25)

# Membuat frame untuk judul aplikasi
title_frame = tk.Frame(root, bg="#007B43", padx=10, pady=10)
title_frame.pack(fill=tk.X)

# # Menambahkan logo pada frame judul aplikasi
# logo_image_path = "images/kemenag.png"  # Jalur relatif gambar
# try:
#     logo_image = Image.open(logo_image_path)  # Membuka gambar dari jalur relatif
#     logo_image = logo_image.resize((100, 100), Image.LANCZOS)  # Menggunakan Image.LANCZOS untuk memperbesar gambar
#     logo_photo = ImageTk.PhotoImage(logo_image)
#     logo_label = tk.Label(title_frame, image=logo_photo, bg="#007B43")
#     logo_label.pack(side=tk.LEFT, padx=10)
# except FileNotFoundError:
#     messagebox.showerror("Error", f"Gambar tidak ditemukan di lokasi {logo_image_path}. Pastikan gambar ada di folder 'images'.")
# except Exception as e:
#     messagebox.showerror("Error", f"Gagal memuat gambar: {str(e)}")


# Menambahkan label judul aplikasi pada frame judul aplikasi
title_label = tk.Label(title_frame, text="APLIKASI MANAJEMEN GPAIDIA", font=("Arial", 18, "bold"), bg="#007B43", fg="white")
title_label.pack(side=tk.LEFT)

# Frame untuk filter
filter_frame = tk.Frame(root)
filter_frame.pack(pady=10)

tk.Label(filter_frame, text="Jenjang Sekolah:").grid(row=0, column=0)
jenjang_var = tk.StringVar(value="Semua")
jenjang_combobox = ttk.Combobox(filter_frame, textvariable=jenjang_var, values=["Semua", "SD", "SMP", "SMA", "SMK"])
jenjang_combobox.grid(row=0, column=1)

tk.Label(filter_frame, text="Sertifikasi:").grid(row=0, column=2)
sertifikasi_var = tk.StringVar(value="Semua")
sertifikasi_combobox = ttk.Combobox(filter_frame, textvariable=sertifikasi_var, values=["Semua", "Sudah", "Belum"])
sertifikasi_combobox.grid(row=0, column=3)

tk.Label(filter_frame, text="Status Pegawai:").grid(row=1, column=0)
status_var = tk.StringVar(value="Semua")
status_combobox = ttk.Combobox(filter_frame, textvariable=status_var, values=["Semua", "PNS", "NON PNS", "PPPK"])
status_combobox.grid(row=1, column=1)

tk.Label(filter_frame, text="Inpassing:").grid(row=1, column=2)
inpassing_var = tk.StringVar(value="Semua")
inpassing_combobox = ttk.Combobox(filter_frame, textvariable=inpassing_var, values=["Semua", "Sudah", "Belum"])
inpassing_combobox.grid(row=1, column=3)

# Dropdown filter untuk Jenis Kelamin
tk.Label(filter_frame, text="Jenis Kelamin:").grid(row=2, column=0)
jenisKelamin_var = tk.StringVar(value="Semua")
jenisKelamin_combobox = ttk.Combobox(filter_frame, textvariable=jenisKelamin_var, values=["Semua", "L", "P"], state="normal")
jenisKelamin_combobox.grid(row=2, column=1)
jenisKelamin_combobox.bind("<<ComboboxSelected>>", filter_data)

tk.Label(filter_frame, text="Pensiun:").grid(row=2, column=2)
pensiun_var = tk.StringVar(value="Semua")
pensiun_combobox = ttk.Combobox(filter_frame, textvariable=pensiun_var, values=["Semua", "Sudah", "Belum"])
pensiun_combobox.grid(row=2, column=3)

#untuk mencari data
tk.Label(filter_frame, text="Cari:").grid(row=3, column=0)
search_var = tk.StringVar()
search_entry = tk.Entry(filter_frame, textvariable=search_var)
search_entry.grid(row=3, column=1)

# Binding event untuk setiap dropdown filter
jenjang_combobox.bind("<<ComboboxSelected>>", filter_data)
sertifikasi_combobox.bind("<<ComboboxSelected>>", filter_data)
status_combobox.bind("<<ComboboxSelected>>", filter_data)
inpassing_combobox.bind("<<ComboboxSelected>>", filter_data)
jenisKelamin_combobox.bind("<<ComboboxSelected>>", filter_data)
pensiun_combobox.bind("<<ComboboxSelected>>", filter_data)

# Event Binding untuk kolom pencarian
search_entry.bind("<KeyRelease>", on_search_change)

# # Tombol untuk menyaring data
# filter_button = tk.Button(filter_frame, text="Filter", command=filter_data)
# filter_button.grid(row=3, column=2)

# Membuat frame untuk tombol
button_frame = tk.Frame(root, padx=10, pady=10)
button_frame.pack(fill=tk.X)

# Menambahkan tombol "Load File"
load_button = ttk.Button(button_frame, text="Load File", command=load_file)
load_button.pack(side=tk.LEFT, padx=5)

# Menambahkan tombol "Save to Excel"
save_button = ttk.Button(button_frame, text="Save to Excel", command=save_to_excel)
save_button.pack(side=tk.LEFT, padx=5)

# Menambahkan tombol "Refresh"
refresh_button = ttk.Button(button_frame, text="Refresh", command=refresh_data)
refresh_button.pack(side=tk.LEFT, padx=5)

# Menambahkan tombol "Show Columns"
show_columns_button = ttk.Button(button_frame, text="Show Columns", command=show_columns)
show_columns_button.pack(side=tk.LEFT, padx=5)

# Treeview untuk menampilkan data
tree_frame = tk.Frame(root)
tree_frame.pack(fill="both", expand=True)

# Tambahkan scrollbar vertikal
tree_scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical")
tree_scrollbar_y.pack(side="right", fill="y")

# Tambahkan scrollbar horizontal
tree_scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal")
tree_scrollbar_x.pack(side="bottom", fill="x")

# Buat Treeview dengan scrollbar
tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set)
tree.pack(fill="both", expand=True)

# Konfigurasi scrollbar
tree_scrollbar_y.config(command=tree.yview)
tree_scrollbar_x.config(command=tree.xview)

# Menyimpan lebar asli kolom
original_widths = {}

# Membuat label informasi jumlah data
info_label = ttk.Label(root, text="Jumlah PNS: 0, Jumlah NON PNS: 0, Jumlah PPPK: 0", font=("Arial", 12))
info_label.pack(pady=10)

# Membuat frame untuk elemen hak cipta dan tombol Desk
footer_frame = tk.Frame(root, padx=10, pady=10)
footer_frame.pack(side=tk.BOTTOM, fill=tk.X)

# Menambahkan label hak cipta
copyright_label = ttk.Label(footer_frame, text="© UINMA 2024 GPAIDIA. All rights reserved.", font=("Arial", 10))
copyright_label.pack(side=tk.LEFT)

def show_desk_info():
    # Membuat jendela baru
    desk_window = tk.Toplevel()
    desk_window.title("Panduan Aplikasi Manajemen GPAIDIA")
    desk_window.geometry("500x400")

    # Menambahkan widget Text untuk menampilkan informasi
    text_widget = scrolledtext.ScrolledText(desk_window, wrap=tk.WORD, width=60, height=15)
    text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    # Menambahkan isi teks
    info_text = """
    Selamat datang di Aplikasi Manajemen GPAIDIA! Aplikasi ini dirancang untuk memudahkan pengelolaan data pegawai dengan fitur-fitur berikut:

    **Fitur Utama:**
    
    1. **Pemrosesan Usia dan Status Pensiun:**
       - Menghitung usia dan status pensiun berdasarkan tanggal lahir pegawai.
    
    2. **Pembacaan Data dari Excel:**
       - Muat data dari file Excel (.xls, .xlsx, dll.) dan ganti nilai kosong di kolom tertentu dengan "Belum" atau "Sudah".
    
    3. **Penampilan Data:**
       - Tampilkan data dalam bentuk tabel dengan kolom yang dapat disesuaikan.
       - Filter data berdasarkan jenjang sekolah, sertifikasi, status pegawai, inpassing, jenis kelamin, dan status pensiun.
    
    4. **Pencarian Data:**
       - Cari data dengan kata kunci. Hasil pencarian akan menampilkan baris yang cocok, dan highlight kata yang dicari. Harap tunggu sebentar saat pencarian dilakukan.
    
    5. **Ekspor Data:**
       - Simpan data yang ditampilkan ke file Excel.
    
    6. **Penyaringan Kolom:**
       - Pilih kolom yang ingin ditampilkan atau disembunyikan. (Fitur ini masih dalam pengembangan).

    **Pengembangan dan Kontak:**
    Aplikasi ini dikembangkan oleh Tim PKL UIN Malang:
        - Siti Rofidatus Saidah
        - Mutiara Aprillia Dzakiroh
        - Nurjihan Nabilah Ramadhani
        - An Nisa’ Puja Karimah Attamimi
    - Periode PKL: 24 Juni - 26 Juli 2024

    Untuk informasi lebih lanjut, hubungi:
    - CP: 082140717475 atau 085336520371
    """
    text_widget.insert(tk.END, info_text)
    text_widget.config(state=tk.DISABLED)  # Nonaktifkan edit pada widget

    # Menambahkan tombol Tutorial
    tutorial_button = ttk.Button(desk_window, text="Tutorial", command=show_tutorial_info)
    tutorial_button.pack(pady=10)

    # Menambahkan tombol Tutup
    close_button = ttk.Button(desk_window, text="Tutup", command=desk_window.destroy)
    close_button.pack(pady=10)

### 2. Fungsi untuk menampilkan jendela tutorial

def show_desk_info():
    # Membuat jendela baru
    desk_window = tk.Toplevel()
    desk_window.title("Panduan Aplikasi Manajemen GPAIDIA")
    desk_window.geometry("600x500")

    # Menambahkan widget ScrolledText untuk menampilkan informasi
    text_widget = scrolledtext.ScrolledText(desk_window, wrap=tk.WORD, width=80, height=20)
    text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    # Menambahkan tag untuk teks bold
    text_widget.tag_configure("bold", font=("Arial", 10, "bold"))

    # Menambahkan isi teks dengan tag bold
    info_text = (
        "Selamat datang di Aplikasi Manajemen GPAIDIA! Aplikasi ini dirancang untuk memudahkan pengelolaan data pegawai dengan fitur-fitur berikut:\n\n"
        "Fitur Utama:\n\n"
        "1. Pemrosesan Usia dan Status Pensiun:\n"
        "- Menghitung usia dan status pensiun berdasarkan tanggal lahir pegawai.\n\n"
        "2. Pembacaan Data dari Excel:\n"
        "- Muat data dari file Excel (.xls, .xlsx, dll.) dan ganti nilai kosong di kolom tertentu dengan 'Belum' atau 'Sudah'.\n\n"
        "3. Penampilan Data:\n"
        "- Tampilkan data dalam bentuk tabel dengan kolom yang dapat disesuaikan.\n"
        "- Filter data berdasarkan jenjang sekolah, sertifikasi, status pegawai, inpassing, jenis kelamin, dan status pensiun.\n\n"
        "4. Pencarian Data:\n"
        "- Cari data dengan kata kunci. Hasil pencarian akan menampilkan baris yang cocok, dan highlight kata yang dicari. Harap tunggu sebentar saat pencarian dilakukan.\n\n"
        "5. Ekspor Data:\n"
        "- Simpan data yang ditampilkan ke file Excel.\n\n"
        "6. Penyaringan Kolom:\n"
        "- Pilih kolom yang ingin ditampilkan atau disembunyikan. (Fitur ini masih dalam pengembangan).\n\n"
        "Cara Menggunakan:\n\n"
        "- Muat Data:\n"
        "Klik 'Load File' dan pilih file Excel yang ingin dimuat.\n\n"
        "- Cari Data:\n"
        "Ketik kata kunci pada kolom pencarian. Data akan di-filter dan ditampilkan setelah beberapa saat.\n\n"
        "- Terapkan Filter:\n"
        "Pilih opsi filter pada dropdown dan klik 'Filter' untuk menyaring data sesuai kriteria yang dipilih.\n\n"
        "- Refresh Data:\n"
        "Klik 'Refresh' untuk mengatur ulang filter dan pencarian, serta memperbarui tampilan data.\n\n"
        "- Tampilkan/Sembunyikan Kolom:\n"
        "Klik 'Show Columns' untuk memilih kolom yang akan ditampilkan atau disembunyikan.\n\n"
        "- Simpan Data:\n"
        "Klik 'Save to Excel' untuk menyimpan data yang ditampilkan ke file Excel.\n\n"
        "Pengembangan dan Kontak:\n"
        "Aplikasi ini dikembangkan oleh Tim PKL UIN Malang:\n"
        "- Siti Rofidatus Saidah\n"
        "- Mutiara Aprillia Dzakiroh\n"
        "- Nurjihan Nabilah Ramadhani\n"
        "- An Nisa’ Puja Karimah Attamimi\n"
        "- Periode PKL: 24 Juni - 26 Juli 2024\n\n"
        "Untuk informasi lebih lanjut, hubungi:\n"
        "- CP: 082140717475 atau 085336520371\n"
    )

    # Menggunakan tag untuk bold bagian tertentu
    text_widget.insert(tk.END, "Selamat datang di Aplikasi Manajemen GPAIDIA! Aplikasi ini dirancang untuk memudahkan pengelolaan data pegawai dengan fitur-fitur berikut:\n\n")
    text_widget.insert(tk.END, "Fitur Utama:\n\n", "bold")
    text_widget.insert(tk.END, "1. Pemrosesan Usia dan Status Pensiun:\n")
    text_widget.insert(tk.END, "- Menghitung usia dan status pensiun berdasarkan tanggal lahir pegawai.\n\n")
    text_widget.insert(tk.END, "2. Pembacaan Data dari Excel:\n")
    text_widget.insert(tk.END, "- Muat data dari file Excel (.xls, .xlsx, dll.) dan ganti nilai kosong di kolom tertentu dengan 'Belum' atau 'Sudah'.\n\n")
    text_widget.insert(tk.END, "3. Penampilan Data:\n")
    text_widget.insert(tk.END, "- Tampilkan data dalam bentuk tabel dengan kolom yang dapat disesuaikan.\n")
    text_widget.insert(tk.END, "- Filter data berdasarkan jenjang sekolah, sertifikasi, status pegawai, inpassing, jenis kelamin, dan status pensiun.\n\n")
    text_widget.insert(tk.END, "4. Pencarian Data:\n")
    text_widget.insert(tk.END, "- Cari data dengan kata kunci. Hasil pencarian akan menampilkan baris yang cocok, dan highlight kata yang dicari. Harap tunggu sebentar saat pencarian dilakukan.\n\n")
    text_widget.insert(tk.END, "5. Ekspor Data:\n")
    text_widget.insert(tk.END, "- Simpan data yang ditampilkan ke file Excel.\n\n")
    text_widget.insert(tk.END, "6. Penyaringan Kolom:\n")
    text_widget.insert(tk.END, "- Pilih kolom yang ingin ditampilkan atau disembunyikan. (Fitur ini masih dalam pengembangan).\n\n")
    text_widget.insert(tk.END, "Pengembangan dan Kontak:\n", "bold")
    text_widget.insert(tk.END, "Aplikasi ini dikembangkan oleh Tim PKL UIN Malang:\n")
    text_widget.insert(tk.END, "- Siti Rofidatus Saidah\n")
    text_widget.insert(tk.END, "- Mutiara Aprillia Dzakiroh\n")
    text_widget.insert(tk.END, "- Nurjihan Nabilah Ramadhani\n")
    text_widget.insert(tk.END, "- An Nisa’ Puja Karimah Attamimi\n")
    text_widget.insert(tk.END, "- Periode PKL: 24 Juni - 26 Juli 2024\n\n")
    text_widget.insert(tk.END, "Untuk informasi lebih lanjut, hubungi:\n")
    text_widget.insert(tk.END, "- CP: 082140717475 atau 085336520371\n")

    text_widget.config(state=tk.DISABLED)  # Nonaktifkan edit pada widget

    # Menambahkan tombol Tutorial
    tutorial_button = ttk.Button(desk_window, text="Tutorial", command=show_tutorial_info)
    tutorial_button.pack(pady=10)

    # Menambahkan tombol Tutup
    close_button = ttk.Button(desk_window, text="Tutup", command=desk_window.destroy)
    close_button.pack(pady=10)   

def show_tutorial_info():
    # Membuat jendela baru
    tutorial_window = tk.Toplevel()
    tutorial_window.title("Cara Menggunakan Aplikasi")
    tutorial_window.geometry("600x500")

    # Menambahkan widget ScrolledText untuk menampilkan tutorial
    text_widget = scrolledtext.ScrolledText(tutorial_window, wrap=tk.WORD, width=80, height=20)
    text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    # Menambahkan tag untuk teks bold
    text_widget.tag_configure("bold", font=("Arial", 10, "bold"))

    # Menambahkan isi teks dengan tag bold
    tutorial_text = (
        "Cara Menggunakan:\n\n"
        "- Muat Data:\n"
        "Klik 'Load File' dan pilih file Excel yang ingin dimuat.\n\n"
        "- Cari Data:\n"
        "Ketik kata kunci pada kolom pencarian. Data akan di-filter dan ditampilkan setelah beberapa saat.\n\n"
        "- Terapkan Filter:\n"
        "Pilih opsi filter pada dropdown dan klik 'Filter' untuk menyaring data sesuai kriteria yang dipilih. Filter akan diterapkan pada data yang ditampilkan berdasarkan kolom yang dipilih.\n\n"
        "- Refresh Data:\n"
        "Klik 'Refresh' untuk mengatur ulang filter dan pencarian, serta memperbarui tampilan data.\n\n"
        "- Tampilkan/Sembunyikan Kolom:\n"
        "Klik 'Show Columns' untuk memilih kolom yang akan ditampilkan atau disembunyikan.\n\n"
        "- Simpan Data:\n"
        "Klik 'Save to Excel' untuk menyimpan data yang ditampilkan ke file Excel.\n"
    )

    # Menggunakan tag untuk bold bagian tertentu
    text_widget.insert(tk.END, "Cara Menggunakan:\n\n", "bold")
    text_widget.insert(tk.END, "- Muat Data:\n")
    text_widget.insert(tk.END, "Klik 'Load File' dan pilih file Excel yang ingin dimuat.\n\n")
    text_widget.insert(tk.END, "- Cari Data:\n")
    text_widget.insert(tk.END, "Ketik kata kunci pada kolom pencarian. Data akan di-filter dan ditampilkan setelah beberapa saat.\n\n")
    text_widget.insert(tk.END, "- Terapkan Filter:\n")
    text_widget.insert(tk.END, "Pilih opsi filter pada dropdown dan klik 'Filter' untuk menyaring data sesuai kriteria yang dipilih. Filter akan diterapkan pada data yang ditampilkan berdasarkan kolom yang dipilih.\n\n")
    text_widget.insert(tk.END, "- Refresh Data:\n")
    text_widget.insert(tk.END, "Klik 'Refresh' untuk mengatur ulang filter dan pencarian, serta memperbarui tampilan data.\n\n")
    text_widget.insert(tk.END, "- Tampilkan/Sembunyikan Kolom:\n")
    text_widget.insert(tk.END, "Klik 'Show Columns' untuk memilih kolom yang akan ditampilkan atau disembunyikan.\n\n")
    text_widget.insert(tk.END, "- Simpan Data:\n")
    text_widget.insert(tk.END, "Klik 'Save to Excel' untuk menyimpan data yang ditampilkan ke file Excel.\n")

    text_widget.config(state=tk.DISABLED)  # Nonaktifkan edit pada widget

    # Menambahkan tombol Tutup
    close_button = ttk.Button(tutorial_window, text="Tutup", command=tutorial_window.destroy)
    close_button.pack(pady=10)

def show_welcome_message():
    messagebox.showinfo(
        "Selamat Datang",
        "Selamat datang di Aplikasi Manajemen GPAIDIA! Sebelum menggunakan aplikasi, dimohon untuk klik tombol Desk dan Tutorial. Jika sudah, bisa di-close."
    )

# Menambahkan tombol "Desk"
desk_button = ttk.Button(footer_frame, text="Desk", command=show_desk_info)
desk_button.pack(side=tk.RIGHT)

# Menambahkan tombol "Tutorial"
tutorial_button = ttk.Button(footer_frame, text="Tutorial", command=show_tutorial_info)
tutorial_button.pack(side=tk.RIGHT)

# Menampilkan pesan selamat datang saat aplikasi pertama kali dijalankan
show_welcome_message()

# Fungsi untuk menutup aplikasi dengan pesan konfirmasi
def close_app():
    if messagebox.askokcancel("Konfirmasi", "Apakah Anda yakin ingin keluar?"):
        root.destroy()

# # Menambahkan ikon aplikasi dan logo
# logo_image_path = "images/kemenag.png"  # Ganti dengan path logo yang benar
# root.iconphoto(False, logo_photo)
# root.protocol("WM_DELETE_WINDOW", close_app)

# Menjalankan aplikasi
root.mainloop()
