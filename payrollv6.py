import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from fpdf import FPDF
import os
import subprocess
from datetime import datetime
import tkinter.font as tkFont

df_global = None  # Untuk menyimpan data excel yang di-load

def open_file():
    global df_global
    file_path = filedialog.askopenfilename(
        title="Pilih file Excel Gaji",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if file_path:
        label_file.config(text=f"📂 File dipilih:\n{file_path}")
        try:
            df = pd.read_excel(file_path)
            df.fillna(0, inplace=True)
            df_global = df
            tampilkan_excel(df)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file:\n{e}")
    else:
        label_file.config(text="❌ Tidak ada file yang dipilih")



def tampilkan_excel(df):
    for widget in frame_table.winfo_children():
        widget.destroy()

    tree = ttk.Treeview(frame_table)
    tree.pack(fill=tk.BOTH, expand=True)

    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="center")

    for _, row in df.iterrows():
        tree.insert("", tk.END, values=list(row))

def generate_pdf_clicked():
    if df_global is None:
        messagebox.showwarning("Peringatan", "Silakan pilih file Excel terlebih dahulu.")
        return
    generate_slip_gaji(df_global)
    messagebox.showinfo("Sukses", "Slip gaji berhasil dibuat di folder 'slip_gaji/'")


def format_tanggal_indonesia():
    bulan_indonesia = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    sekarang = datetime.now()
    return f"{sekarang.day} {bulan_indonesia[sekarang.month - 1]} {sekarang.year}"

def generate_slip_gaji(df):
    output_dir = "slip_gaji"
    os.makedirs(output_dir, exist_ok=True)



    for _, row in df.iterrows():
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=10)

        # ======= FONT SIZE SETUP =======
        FONT_TITLE = 14
        FONT_LABEL = 11
        FONT_TEXT = 10
        FONT_FOOTER = 9

        # ======= HEADER WITH LOGO & COMPANY INFO =======
        logo_path = "logo_inix.png"  # Nama file logo di folder yang sama dengan script
        pdf.image(logo_path, x=10, y=10, w=30)  # Menyesuaikan ukuran dan posisi logo
        
        # Geser posisi teks ke kanan agar tidak tumpang tindih dengan logo
        pdf.set_xy(42, 14)
        pdf.set_font("Arial", "B", 14)  # Menyesuaikan ukuran font untuk nama perusahaan
        pdf.cell(0, 8, "PT. Inixindo Widya Utama", ln=True)

        pdf.ln(4)
        pdf.set_font("Arial", "", 10)  # Ukuran font untuk alamat dan info lainnya
        pdf.cell(0, 5, "Jl. Tenggilis Barat I/19 (D17), Surabaya 60292,", ln=True)
        pdf.cell(0, 5, "Jawa Timur, Indonesia.", ln=True)
        pdf.cell(0, 5, "Email: info@inixindosurabaya.id", ln=True)
        pdf.cell(0, 5, "Telepon: +62318477733 (Pada Jam Kerja)", ln=True)
        pdf.cell(0, 5, "WA: +628819606907 (Fast Response)", ln=True)

        pdf.ln(4)  # Memberikan jarak sebelum garis pembatas bawah kop

        # Garis pembatas bawah kop
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.5)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(5)

        # ======= SLIP GAJI CONTENT =======
        pdf.set_font("Arial", "B", FONT_TITLE)
        pdf.cell(0, 10, "SLIP GAJI KARYAWAN", ln=True, align="C")
        pdf.ln(4)

        pdf.set_font("Arial", "", FONT_TEXT)
        pdf.cell(0, 6, f"Periode: {row.get('Periode', 'N/A')}", ln=True)
        pdf.cell(0, 6, f"Total Hari Masuk: {row.get('Total Hari Masuk', 'N/A')}", ln=True)
        pdf.ln(2)

        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.cell(0, 6, f"Nama: {row.get('Nama', '')}", ln=True)
        pdf.cell(0, 6, f"NIK: {row.get('NIK', '')}", ln=True)
        pdf.ln(3)

        pdf.set_draw_color(0, 0, 0)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(6)



        # ======= PENDAPATAN =======
        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.set_fill_color(230, 230, 230)  # Abu-abu muda
        pdf.cell(90, 8, "Pendapatan", border=1, align="C", fill=True)
        pdf.cell(0, 8, "Nominal", border=1, align="C", fill=True)
        pdf.ln()

        
        pdf.set_font("Arial", "", FONT_TEXT)
        for label in ["Gaji Pokok", "Tunjangan Kehadiran", "Lembur", "Komisi", "Bonus"]:
            pdf.cell(90, 6, label, border=1)
            pdf.cell(0, 6, f"Rp {row.get(label, 0):,}", border=1, align="R")
            pdf.ln()

        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.cell(90, 7, "TOTAL PENDAPATAN (A)", border=1, align="C")
        pdf.cell(0, 7, f"Rp {row.get('TOTAL PENDAPATAN (A)', 0):,}", border=1, align="R")
        pdf.ln(12)

        # ======= POTONGAN =======
        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.set_fill_color(230, 230, 230)
        pdf.cell(90, 8, "Potongan", border=1, align="C", fill=True)
        pdf.cell(0, 8, "Nominal", border=1, align="C", fill=True)
        pdf.ln()
        
        pdf.set_font("Arial", "", FONT_TEXT)
        for label in ["Potongan Tidak Masuk", "PPH 21", "Potongan Lainnya"]:
            pdf.cell(90, 6, label, border=1)
            pdf.cell(0, 6, f"Rp {row.get(label, 0):,}", border=1, align="R")
            pdf.ln()

        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.cell(90, 7, "TOTAL POTONGAN (B)", border=1, align="C")
        pdf.cell(0, 7, f"Rp {row.get('Total Potongan (B)', 0):,}", border=1, align="R")
        pdf.ln(8)

        # ======= TAKE HOME PAY =======
        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.set_fill_color(230, 230, 230)
        pdf.cell(90, 8, "Take Home Pay", border=1, align="C", fill=True)
        pdf.cell(0, 8, f"Rp {row.get('Take Home Pay', 0):,}", border=1, align="R", fill=True)
        pdf.ln(20)


        # ======= FOOTER (2 columns: Mengetahui & Diterima oleh) =======
        tanggal_str = format_tanggal_indonesia()
        pdf.set_font("Arial", "", FONT_TEXT)
        pdf.cell(0, 6, tanggal_str, ln=True)
        pdf.ln(1)

        col_width = 90  # half of 180 page width
        pdf.set_font("Arial", "", FONT_TEXT)
        pdf.cell(col_width, 6, "Mengetahui,", align="L")
        pdf.cell(col_width, 6, "Diterima oleh,", align="R")
        pdf.ln(20)

        pdf.cell(col_width, 6, "Bambang Soerjohandoko", align="L")
        pdf.cell(col_width, 6, row.get('Nama', ''), align="R")
        pdf.ln(6)

        pdf.set_font("Arial", "", FONT_TEXT)
        pdf.cell(col_width, 6, "(Direktur Utama)", align="L")
        pdf.cell(col_width, 6, "", align="R")
        pdf.ln(20)

        # ======= DISCLAIMER =======
        pdf.set_font("Arial", "I", FONT_FOOTER)
        pdf.multi_cell(0, 5,
            "Dokumen slip gaji ini diterbitkan secara resmi oleh PT Inixindo Widya Utama "
            "melalui sistem informasi internal. Dokumen ini sah dan berlaku tanpa memerlukan tanda tangan basah maupun stempel."
        )

        filename = os.path.join(output_dir, f"{row['Nama'].replace(' ', '_')}_Slip_Gaji.pdf")
        pdf.output(filename)


def open_folder():
    output_dir = "slip_gaji"
    if os.path.exists(output_dir):
        if os.name == 'nt':  # Windows
            os.startfile(output_dir)
        else:  # Mac or Linux
            subprocess.call(["open", output_dir])
    else:
        messagebox.showwarning("Peringatan", "Folder 'slip_gaji' tidak ditemukan!")


       

root = tk.Tk()
root.title("Slip Gaji - Generate PDF")
root.geometry("1000x650")
root.configure(bg="#f0f2f5")


style = ttk.Style(root)
style.theme_use("default")
style.configure("Treeview",
                background="#ffffff",
                foreground="#333333",
                rowheight=30,
                fieldbackground="#ffffff",
                font=("Segoe UI", 10))
style.map("Treeview", background=[("selected", "#007acc")], foreground=[("selected", "#ffffff")])
style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"), background="#e0e0e0", foreground="#333")

header_frame = tk.Frame(root, bg="#f0f2f5")
header_frame.pack(pady=10)

tk.Label(header_frame, text="📄 Aplikasi Slip Gaji", font=("Segoe UI", 18, "bold"), bg="#f0f2f5").pack()
tk.Label(header_frame, text="Upload file Excel dan generate slip gaji dalam format PDF.",
         font=("Segoe UI", 10), bg="#f0f2f5", fg="#555").pack()

btn_frame = tk.Frame(root, bg="#f0f2f5")
btn_frame.pack(pady=10)



tk.Button(btn_frame, text="📁 Pilih File Excel", command=open_file, width=20, bg="#4caf50", fg="white",
          font=("Segoe UI", 10, "bold")).grid(row=0, column=0, padx=10)
tk.Button(btn_frame, text="🖨️ Generate Slip Gaji", command=generate_pdf_clicked, width=20, bg="#2196f3", fg="white",
          font=("Segoe UI", 10, "bold")).grid(row=0, column=1, padx=10)
tk.Button(btn_frame, text="📂 Buka Folder Slip", command=open_folder, width=20, bg="#ff9800", fg="white",
          font=("Segoe UI", 10, "bold")).grid(row=0, column=2, padx=10)

label_file = tk.Label(root, text="📂 Belum ada file yang dipilih", font=("Segoe UI", 10), bg="#f0f2f5", anchor="w")
label_file.pack(pady=5, anchor="w", padx=20)


frame_table = tk.Frame(root)
frame_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

scroll_x = tk.Scrollbar(frame_table, orient=tk.HORIZONTAL)
scroll_y = tk.Scrollbar(frame_table, orient=tk.VERTICAL)

tree = ttk.Treeview(frame_table, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

scroll_y.config(command=tree.yview)
scroll_x.config(command=tree.xview)

root.mainloop()
