import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from fpdf import FPDF
import os
import subprocess
from datetime import datetime
import tkinter.font as tkFont
import uuid
import sys
from PyPDF2 import PdfWriter, PdfReader
import io


# Predefined login credentials
LOGIN_CREDENTIALS = {
    "username": "admin",
    "password": "inixindo123"
}

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

df_global = None
output_dir = ""

# Button references untuk enable/disable
btn_pilih_excel = None
btn_pilih_lokasi = None
btn_generate_pdf = None
btn_buka_folder = None
btn_blast_email = None

def update_button_states():
    """Update status enable/disable button berdasarkan langkah yang sudah diselesaikan"""
    # Button 1: Pilih File Excel - selalu enable
    btn_pilih_excel.state(['!disabled'])
    
    # Button 2: Pilih Lokasi - enable setelah file excel dipilih
    if df_global is not None:
        btn_pilih_lokasi.state(['!disabled'])
    else:
        btn_pilih_lokasi.state(['disabled'])
    
    # Button 3: Generate PDF - enable setelah file excel dan lokasi dipilih
    if df_global is not None and output_dir:
        btn_generate_pdf.state(['!disabled'])
    else:
        btn_generate_pdf.state(['disabled'])
    
    # Button 4: Buka Folder - enable setelah PDF digenerate
    if df_global is not None and output_dir and os.path.exists(output_dir) and any(f.endswith('.pdf') for f in os.listdir(output_dir)):
        btn_buka_folder.state(['!disabled'])
    else:
        btn_buka_folder.state(['disabled'])
    
    # Button 5: Blast Email - enable setelah PDF digenerate
    if df_global is not None and output_dir and os.path.exists(output_dir) and any(f.endswith('.pdf') for f in os.listdir(output_dir)):
        btn_blast_email.state(['!disabled'])
    else:
        btn_blast_email.state(['disabled'])

# Login Window
def create_login_window():
    login_window = tk.Toplevel()
    login_window.title("Login - Slip Gaji Inixindo")
    login_window.geometry("350x200")
    login_window.resizable(False, False)
    
    # Center the window
    window_width = 350
    window_height = 200
    screen_width = login_window.winfo_screenwidth()
    screen_height = login_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    login_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # Make it modal
    login_window.grab_set()
    
    # Login Frame
    login_frame = tk.Frame(login_window, padx=20, pady=20)
    login_frame.pack(expand=True, fill="both")
    
    tk.Label(login_frame, text="Username:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=(0, 5))
    username_entry = tk.Entry(login_frame, font=("Segoe UI", 10))
    username_entry.grid(row=0, column=1, sticky="ew", pady=(0, 5))
    
    tk.Label(login_frame, text="Password:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=(0, 5))
    password_entry = tk.Entry(login_frame, show="*", font=("Segoe UI", 10))
    password_entry.grid(row=1, column=1, sticky="ew", pady=(0, 5))
    
    def attempt_login():
        username = username_entry.get()
        password = password_entry.get()
        
        if username == LOGIN_CREDENTIALS["username"] and password == LOGIN_CREDENTIALS["password"]:
            login_window.destroy()
            root.deiconify()  # Show the main window
            update_button_states()  # Update button states setelah login
        else:
            messagebox.showerror("Login Gagal", "Username atau password salah!")
    
    login_button = tk.Button(login_frame, text="Login", command=attempt_login, 
                           bg="#4caf50", fg="white", font=("Segoe UI", 10, "bold"))
    login_button.grid(row=2, column=0, columnspan=2, pady=(10, 0), sticky="ew")
    
    # Bind Enter key to login
    login_window.bind('<Return>', lambda event: attempt_login())
    
    # Set focus to username field
    username_entry.focus_set()
    
    return login_window

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
            update_button_states()  # Update button states setelah file dipilih
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file:\n{e}")
    else:
        label_file.config(text="❌ Tidak ada file yang dipilih")

def tampilkan_excel(df):
    tree.delete(*tree.get_children())
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"

    font = tkFont.Font(font=("Segoe UI", 10))
    header_font = tkFont.Font(font=("Segoe UI", 11, "bold"))

    # Daftar kolom yang tidak perlu format Rupiah
    exclude_columns = ['Periode', 'Nama', 'NIK', 'Email', 'Total Hari Masuk']
    
    for col in df.columns:
        tree.heading(col, text=col)
        
        max_width = header_font.measure(col) + 40
        
        # Format khusus untuk kolom numerik (selain yang dikecualikan)
        if col not in exclude_columns and pd.api.types.is_numeric_dtype(df[col]):
            for item in df[col]:
                formatted = f"Rp {int(item):,}" if pd.notnull(item) else ""
                item_width = font.measure(formatted)
                if item_width > max_width:
                    max_width = item_width + 40
        else:
            for item in df[col].astype(str):
                item_width = font.measure(item)
                if item_width > max_width:
                    max_width = item_width + 40

        tree.column(col, anchor="center", width=max_width, minwidth=max_width, stretch=True)

    # Insert data dengan format Rupiah
    for _, row in df.iterrows():
        formatted_values = []
        for col in df.columns:
            if col not in exclude_columns and pd.api.types.is_numeric_dtype(df[col]):
                value = row[col]
                formatted_values.append(f"Rp {int(value):,}" if pd.notnull(value) else "")
            else:
                formatted_values.append(str(row[col]))
        
        tree.insert("", "end", values=formatted_values)

    tree.xview_moveto(0)

def select_pdf_loc():
    global output_dir
    folder_path = filedialog.askdirectory(title="Pilih Lokasi Penyimpanan Slip Gaji")
    if folder_path:
        output_dir = folder_path
        messagebox.showinfo("Lokasi Tersimpan", f"File PDF akan disimpan di:\n{output_dir}")
        update_button_states()  # Update button states setelah lokasi dipilih

def generate_pdf_clicked():
    global output_dir
    
    if df_global is None:
        messagebox.showwarning("Peringatan", "Silakan pilih file Excel terlebih dahulu.")
        return
    
    if not output_dir:
        messagebox.showwarning("Peringatan", "Silakan pilih lokasi penyimpanan terlebih dahulu.")
        return

    generate_slip_gaji(df_global)
    messagebox.showinfo("Sukses", f"Slip gaji berhasil dibuat di folder '{output_dir}'")
    update_button_states()  # Update button states setelah PDF digenerate

def format_tanggal_indonesia():
    bulan_indonesia = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    sekarang = datetime.now()
    return f"{sekarang.day} {bulan_indonesia[sekarang.month - 1]} {sekarang.year}"

def generate_slip_gaji(df):
    os.makedirs(output_dir, exist_ok=True)

    for _, row in df.iterrows():
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=10)

        # Font setup
        FONT_TITLE = 14
        FONT_LABEL = 11
        FONT_TEXT = 10
        FONT_FOOTER = 9

        # Border setup
        border_margin = 6
        page_width = 210
        border_width = 0.5

        # Draw border
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(border_width)
        pdf.rect(border_margin, border_margin, 
                page_width - 2*border_margin, 
                297 - 2*border_margin)

        content_x = 10 + border_margin
        content_y = 10 + border_margin
        pdf.set_xy(content_x, content_y)

        # Header with logo
        logo_path = resource_path("logo_inix.png")
        pdf.image(logo_path, x=content_x-5, y=content_y, w=30)
        
        # Company info
        pdf.set_xy(content_x + 28, content_y + 4)
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 8, "PT. Inixindo Widya Utama", ln=True)

        pdf.ln(2)
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 5, "Jl. Tenggilis Barat I/19 (D17), Surabaya 60292,", ln=True)
        pdf.cell(0, 5, "Jawa Timur, Indonesia.", ln=True)
        pdf.cell(0, 5, "Email: info@inixindosurabaya.id", ln=True)
        pdf.cell(0, 5, "Telepon: +62318477733 (Pada Jam Kerja)", ln=True)
        pdf.cell(0, 5, "WA: +628819606907 (Fast Response)", ln=True)

        pdf.ln(4)
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.5)
        pdf.line(content_x, pdf.get_y(), page_width - content_x, pdf.get_y())
        pdf.ln(5)

        # Slip gaji content
        pdf.set_font("Arial", "B", FONT_TITLE)
        pdf.cell(0, 10, "SLIP GAJI KARYAWAN", ln=True, align="C")
        pdf.ln(4)

        pdf.set_font("Arial", "", FONT_TEXT)
        pdf.cell(0, 6, f"Periode: {row.get('Periode', 'N/A')}", ln=True)
        pdf.cell(0, 6, f"Total Hari Masuk: {row.get('Total Hari Masuk', 'N/A')}", ln=True)
        pdf.cell(0, 6, f"Terlambat: {row.get('Terlambat', 'N/A')}x | Alpha: {row.get('Alpha', 'N/A')}x", ln=True)
        pdf.ln(2)

        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.cell(0, 6, f"Nama: {row.get('Nama', '')}", ln=True)
        pdf.cell(0, 6, f"NIK: {row.get('NIK', '')}", ln=True)
        pdf.cell(0, 6, f"Email: {row.get('Email', '')}", ln=True)
        pdf.ln(3)

        pdf.set_draw_color(0, 0, 0)
        pdf.line(content_x, pdf.get_y(), page_width - content_x, pdf.get_y())
        pdf.ln(6)

        # Pendapatan section
        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.set_fill_color(235, 235, 235)
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

        # Potongan section
        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.set_fill_color(235, 235, 235)
        pdf.cell(90, 8, "Potongan", border=1, align="C", fill=True)
        pdf.cell(0, 8, "Nominal", border=1, align="C", fill=True)
        pdf.ln()
        
        pdf.set_font("Arial", "", FONT_TEXT)
        for label in ["Tidak Masuk", "PPH 21", "Potongan Lainnya"]:
            pdf.cell(90, 6, label, border=1)
            pdf.cell(0, 6, f"Rp {row.get(label, 0):,}", border=1, align="R")
            pdf.ln()

        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.cell(90, 7, "TOTAL POTONGAN (B)", border=1, align="C")
        pdf.cell(0, 7, f"Rp {row.get('Total Potongan (B)', 0):,}", border=1, align="R")
        pdf.ln(12)

        # Take Home Pay
        pdf.set_font("Arial", "B", FONT_LABEL)
        pdf.set_fill_color(235, 235, 235)
        pdf.cell(90, 8, "Take Home Pay", border=1, align="C", fill=True)
        pdf.cell(0, 8, f"Rp {row.get('Take Home Pay', 0):,}", border=1, align="R", fill=True)
        pdf.ln(20)

        # Footer
        tanggal_str = format_tanggal_indonesia()
        pdf.set_font("Arial", "", FONT_TEXT)
        pdf.cell(0, 6, tanggal_str, ln=True)
        pdf.ln(1)

        col_width = 90
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

        # Disclaimer
        pdf.set_font("Arial", "I", FONT_FOOTER)
        pdf.multi_cell(0, 5,
            "Dokumen slip gaji ini diterbitkan secara resmi oleh PT. Inixindo Widya Utama "
            "melalui sistem informasi internal. Dokumen ini sah dan berlaku tanpa memerlukan tanda tangan basah maupun stempel."
        )
       
        periode = row['Periode']
        safe_name = row['Nama'].replace(' ', '_')
        filename = os.path.join(output_dir, f"{safe_name}_Slip_Gaji_{periode}.pdf")
        
        # Simpan PDF ke buffer
        pdf_bytes = pdf.output(dest='S').encode('latin1')
        
        # Buat PDF yang dipassword
        pdf_reader = PdfReader(io.BytesIO(pdf_bytes))
        pdf_writer = PdfWriter()
        
        # Tambahkan semua halaman ke writer
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)
        
        # Enkripsi dengan NIK sebagai password
        nik_password = str(row.get('NIK', '123456'))  # Default password jika NIK kosong
        pdf_writer.encrypt(nik_password)
        
        # Simpan PDF yang sudah dienkripsi
        with open(filename, "wb") as f:
            pdf_writer.write(f)

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import ssl

def blast_email():
    global df_global, output_dir
    
    if df_global is None:
        messagebox.showwarning("Peringatan", "Silakan pilih file Excel terlebih dahulu.")
        return
    
    if not output_dir:
        messagebox.showwarning("Peringatan", "Silakan generate slip gaji terlebih dahulu.")
        return

    # Konfigurasi email
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    sender_email = "masbrownid@gmail.com"
    sender_password = "goia mmwe atit zydp"
    
    confirm = messagebox.askyesno("Konfirmasi", 
                                "Anda akan mengirim email ke semua karyawan.\n"
                                f"Email pengirim: {sender_email}\n"
                                "Lanjutkan?")
    if not confirm:
        return
    
    progress_window = tk.Toplevel()
    progress_window.title("Progress Pengiriman Email")
    progress_window.geometry("400x200")
    
    progress_label = tk.Label(progress_window, text="Mengirim email...", font=("Segoe UI", 12))
    progress_label.pack(pady=20)
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=len(df_global))
    progress_bar.pack(fill="x", padx=20, pady=10)
    
    status_label = tk.Label(progress_window, text="", font=("Segoe UI", 10))
    status_label.pack(pady=10)
    
    def send_emails():
        context = ssl.create_default_context()
        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls(context=context)
                server.login(sender_email, sender_password)
            
                total_sent = 0
                for i, (_, row) in enumerate(df_global.iterrows()):
                    recipient_email = row['Email']
                    if not pd.isna(recipient_email) and "@" in str(recipient_email):
                        try:
                            # Buat email
                            msg = MIMEMultipart()
                            msg['From'] = sender_email
                            msg['To'] = recipient_email
                        
                            periode = row['Periode']
                            msg['Subject'] = f"Slip Gaji {periode} - {row['Nama']}"
                        
                            # Isi email
                            body = f"""
                            <html>
                                <body>
                                    <p>Yth. {row['Nama']},</p>
                                    <p>Berikut kami sampaikan slip gaji Anda untuk periode {periode}.</p>
                                    <p>Take Home Pay: <b>Rp {row['Take Home Pay']:,}</b></p>
                                    <p>Password untuk membuka file PDF adalah NIK Anda: <b>{row['NIK']}</b></p>
                                    <p>Slip gaji dapat diunduh pada lampiran email ini.</p>
                                    <br>
                                    <p>Hormat kami,</p>
                                    <p>HRD PT. Inixindo Widya Utama</p>
                                </body>
                            </html>
                            """
                            msg.attach(MIMEText(body, 'html'))
                        
                            # Cari file PDF yang sudah digenerate
                            safe_name = row['Nama'].replace(' ', '_')
                            safe_periode = periode
                            pdf_filename = f"{safe_name}_Slip_Gaji_{safe_periode}.pdf"
                            pdf_path = os.path.join(output_dir, pdf_filename)
                        
                            if os.path.exists(pdf_path):
                                # Baca file PDF yang sudah dienkripsi
                                with open(pdf_path, "rb") as f:
                                    file_data = f.read()
                                
                                # Lampirkan PDF
                                part = MIMEApplication(file_data, Name=pdf_filename)
                                part['Content-Disposition'] = f'attachment; filename="{pdf_filename}"'
                                msg.attach(part)
                            
                                # Kirim email
                                server.sendmail(sender_email, recipient_email, msg.as_string())
                                total_sent += 1
                            
                                # Update progress
                                progress_var.set(i+1)
                                status_label.config(text=f"Mengirim ke {row['Nama']} ({recipient_email})")
                                progress_window.update()
                            else:
                                print(f"File PDF tidak ditemukan: {pdf_path}")
                        
                        except Exception as e:
                            print(f"Gagal mengirim ke {recipient_email}: {str(e)}")
            
                messagebox.showinfo("Sukses", f"Email berhasil dikirim ke {total_sent} karyawan")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengirim email: {str(e)}")
        finally:
            progress_window.destroy()
    
    # Jalankan di thread terpisah
    import threading
    threading.Thread(target=send_emails, daemon=True).start()

def open_folder():
    if os.path.exists(output_dir):
        if os.name == 'nt':
            os.startfile(output_dir)
        else:
            subprocess.call(["open", output_dir])
    else:
        messagebox.showwarning("Peringatan", "Folder tidak ditemukan!")

# Main GUI Setup
root = tk.Tk()
root.title("Slip Gaji - Generate PDF")
root.geometry("1000x650")
root.configure(bg="#f0f2f5")

# Hide main window initially
root.withdraw()

# Create login window
create_login_window()

style = ttk.Style(root)
style.theme_use("default")

# Tambahkan konfigurasi untuk button disabled
style.configure('TButton', 
               font=("Segoe UI", 10, "bold"),
               padding=6)

style.map('TButton',
          foreground=[('disabled', 'gray')],
          background=[('disabled', 'white')]
          )

# Style khusus untuk masing-masing tombol
style.configure('Excel.TButton', background='#4caf50', foreground='white')
style.configure('Lokasi.TButton', background='#ff8c00', foreground='white')
style.configure('Generate.TButton', background='#2196f3', foreground='white')
style.configure('Folder.TButton', background='#ff5722', foreground='white')
style.configure('Email.TButton', background='#9c27b0', foreground='white')

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

tk.Label(header_frame, text="📄 App Slip Gaji Inixindo Surabaya", font=("Segoe UI", 18, "bold"), bg="#f0f2f5").pack()
tk.Label(header_frame, text="Upload file Excel dan generate slip gaji dalam format PDF.",
         font=("Segoe UI", 10), bg="#f0f2f5", fg="#555").pack()

btn_frame = tk.Frame(root, bg="#f0f2f5")
btn_frame.pack(pady=10)

# Buat button dengan referensi global
btn_pilih_excel = ttk.Button(btn_frame, text="📁 Pilih File Excel", command=open_file, style='Excel.TButton')
btn_pilih_excel.grid(row=0, column=0, padx=10)

btn_pilih_lokasi = ttk.Button(btn_frame, text="📁 Pilih Lokasi Slip Gaji", command=select_pdf_loc, style='Lokasi.TButton')
btn_pilih_lokasi.grid(row=0, column=1, padx=10)
btn_pilih_lokasi.state(['disabled'])

btn_generate_pdf = ttk.Button(btn_frame, text="🖨️ Generate Slip Gaji", command=generate_pdf_clicked, style='Generate.TButton')
btn_generate_pdf.grid(row=0, column=2, padx=10)
btn_generate_pdf.state(['disabled'])

btn_buka_folder = ttk.Button(btn_frame, text="📂 Buka Folder Slip Gaji", command=open_folder, style='Folder.TButton')
btn_buka_folder.grid(row=0, column=3, padx=10)
btn_buka_folder.state(['disabled'])

btn_blast_email = ttk.Button(btn_frame, text="📧 Blasting Email", command=blast_email, style='Email.TButton')
btn_blast_email.grid(row=0, column=4, padx=10)
btn_blast_email.state(['disabled'])

label_file = tk.Label(root, text="❌ Tidak ada file yang dipilih", font=("Segoe UI", 10), fg="gray", bg="#f0f2f5")
label_file.pack(pady=5)

# Frame untuk Scroll
scroll_frame = tk.Frame(root)
scroll_frame.pack(fill="both", expand=True)

tree_frame = tk.Frame(scroll_frame)
tree_frame.pack(fill="both", expand=True)

# Scrollbars
scrollbar_y = tk.Scrollbar(tree_frame, orient="vertical")
scrollbar_y.pack(side="right", fill="y")

scrollbar_x = tk.Scrollbar(tree_frame, orient="horizontal")
scrollbar_x.pack(side="bottom", fill="x")

tree = ttk.Treeview(tree_frame, show="headings", height=8, 
                   yscrollcommand=scrollbar_y.set, 
                   xscrollcommand=scrollbar_x.set)
tree.pack(fill="both", expand=True)

scrollbar_y.config(command=tree.yview)
scrollbar_x.config(command=tree.xview)

root.mainloop()