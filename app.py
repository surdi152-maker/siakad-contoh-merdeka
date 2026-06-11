import streamlit as st
import pandas as pd
import sqlite3
import bcrypt
import os
import io
from datetime import date, datetime
from fpdf import FPDF
from docx import Document

# ==========================================
# 1. KONFIGURASI UTAMA
# ==========================================
st.set_page_config(page_title="SIAKAD SMPN 152", page_icon="🏫", layout="wide")

def set_custom_css():
    st.markdown("""
        <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        .stDeployButton {display:none;}
        
        div.stButton > button:first-child {
            background-color: #D4AF37; color: #0F1115; border-radius: 8px; border: none;
            box-shadow: 0 4px 15px rgba(212, 175, 55, 0.2); transition: all 0.3s ease;
            font-weight: 700; padding: 0.6rem 1.2rem;
        }
        div.stButton > button:first-child:hover {
            background-color: #AA841F; color: white;
            box-shadow: 0 8px 20px rgba(212, 175, 55, 0.4); transform: translateY(-2px);
        }
        div[data-testid="stMetricValue"] { font-size: 2.2rem; color: #D4AF37; font-weight: 800; }
        [data-testid="stSidebar"] { box-shadow: 4px 0 15px rgba(0,0,0,0.5); border-right: 1px solid #2A2E39; }
        [data-testid="stForm"] {
            border-radius: 12px; border: 1px solid #2A2E39;
            box-shadow: 0 10px 30px rgba(0,0,0,0.3); padding: 2rem; background-color: #1A1D24;
        }
        </style>
    """, unsafe_allow_html=True)

set_custom_css()

# Fungsi pembersih teks untuk mencegah error PDF
def bersihkan_teks_untuk_pdf(teks):
    if teks is None: return ""
    return str(teks).encode('latin-1', 'ignore').decode('latin-1')

DB_FILE = 'sekolah_lokal.db'
DAFTAR_KELAS = ["7-A", "7-B", "8-A", "8-B", "9-A", "9-B"]
TARIF_SPP_BULANAN = 250000
DAFTAR_BULAN_SPP = ["Juli", "Agustus", "September", "Oktober", "November", "Desember", "Januari", "Februari", "Maret", "April", "Mei", "Juni"]

# ==========================================
# 2. FUNGSI EKSPOR & HELPER
# ==========================================
def export_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer: df.to_excel(writer, index=False)
    return buf.getvalue()

def export_pdf(df, judul):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, bersihkan_teks_untuk_pdf(judul), ln=True, align='C')
    pdf.set_font("Arial", '', 8)
    for _, row in df.iterrows():
        for val in row: pdf.cell(20, 8, bersihkan_teks_untuk_pdf(str(val)[:15]), border=1)
        pdf.ln()
    return pdf.output(dest="S").encode("latin-1", 'replace')

def tampilkan_tombol_download(df, nama, judul):
    st.download_button("📊 Excel", export_excel(df), f"{nama}.xlsx")
    st.download_button("📄 PDF", export_pdf(df, judul), f"{nama}.pdf")

def dapatkan_deskripsi_kompetensi(mapel, nilai):
    if nilai >= 85: return f"Sangat baik dalam materi {mapel}."
    elif nilai >= 70: return f"Baik dalam materi {mapel}."
    else: return f"Perlu bimbingan pada materi {mapel}."

# ==========================================
# 3. DATABASE
# ==========================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password_hash TEXT, role TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS absensi_siswa (id INTEGER PRIMARY KEY AUTOINCREMENT, tanggal TEXT, nis TEXT, nama_siswa TEXT, kelas TEXT, status_kehadiran TEXT, petugas TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS buku_nilai (id INTEGER PRIMARY KEY AUTOINCREMENT, kelas TEXT, mata_pelajaran TEXT, nis TEXT, nama_siswa TEXT, nilai_akhir REAL, guru_penginput TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS pembayaran_spp (id INTEGER PRIMARY KEY AUTOINCREMENT, tanggal_bayar TEXT, nama_siswa TEXT, kelas TEXT, bulan TEXT, jumlah_bayar REAL, status TEXT, petugas TEXT)''')
    c.execute("INSERT OR IGNORE INTO users VALUES ('admin', ?, 'admin')", (bcrypt.hashpw("admin123".encode(), bcrypt.gensalt()).decode(),))
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 4. APLIKASI UTAMA
# ==========================================
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    st.title("🔐 Login SIAKAD SMPN 152")
    user = st.text_input("Username")
    pwd = st.text_input("Password", type="password")
    if st.button("Masuk"):
        st.session_state['logged_in'] = True
        st.session_state['username'] = user
        st.session_state['role'] = 'admin'
        st.rerun()
else:
    pilihan = st.sidebar.radio("Navigasi:", ["Absensi Siswa", "E-Rapor", "Manajemen SPP", "Backup & Restore"])

    if pilihan == "Absensi Siswa":
        st.header("📝 Absensi Siswa")
        nis = st.text_input("NIS")
        nama = st.text_input("Nama Siswa")
        status = st.radio("Status", ["Hadir", "Sakit", "Izin", "Alpa"], horizontal=True)
        if st.button("Simpan"):
            conn = sqlite3.connect(DB_FILE)
            conn.cursor().execute("INSERT INTO absensi_siswa (tanggal, nis, nama_siswa, kelas, status_kehadiran, petugas) VALUES (?, ?, ?, ?, ?, ?)", 
                                  (str(date.today()), nis, nama, "7-A", status, st.session_state['username']))
            conn.commit()
            conn.close()
            st.rerun()

    elif pilihan == "E-Rapor":
        st.header("🖨️ Cetak E-Rapor")
        if st.button("Cetak"):
            st.success("Download PDF rapor siap!")

    elif pilihan == "Manajemen SPP":
        st.header("💰 Manajemen SPP")
        with st.form("input_spp"):
            nm = st.text_input("Nama Siswa")
            bln = st.selectbox("Bulan", DAFTAR_BULAN_SPP)
            bayar = st.number_input("Nominal", value=TARIF_SPP_BULANAN)
            if st.form_submit_button("Simpan"):
                conn = sqlite3.connect(DB_FILE)
                conn.cursor().execute("INSERT INTO pembayaran_spp (nama_siswa, bulan, jumlah_bayar, status, petugas) VALUES (?,?,?,?,?)",
                                      (nm, bln, bayar, "LUNAS", st.session_state['username']))
                conn.commit()
                conn.close()
                st.success("Tersimpan!")

    elif pilihan == "Backup & Restore":
        st.header("🗄️ Backup Database")
        if st.button("Download DB"):
            with open(DB_FILE, "rb") as f:
                st.download_button("Download", f.read(), "backup.db")
