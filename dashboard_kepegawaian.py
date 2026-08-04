# =====================================================================
# DASHBOARD KEPEGAWAIAN - SEKRETARIAT PEMERINTAHAN DAERAH
# ---------------------------------------------------------------------
# =====================================================================

# --------------------------- IMPORT LIBRARY ----------------------------
import streamlit as st                  # Library utama untuk membangun web app
import pandas as pd                     # Untuk manipulasi data tabular (CSV/excel)
import numpy as np                      # Untuk menghasilkan data dummy KPI yang kosong
import plotly.express as px             # Untuk visualisasi grafik interaktif
import plotly.graph_objects as go       # Untuk visualisasi grafik yang lebih kompleks
import re                               # Untuk manipulasi string dan regex (misal parsing kode kehadiran)
from datetime import date, datetime     #untuk manipulasi tanggal dan waktu
from dateutil.relativedelta import relativedelta  #Untuk menghitung selisih tahun, bulan, hari secara akurat
import os                               # Untuk mengecek keberadaan file CSV
import csv                              # Untuk quoting CSV agar aman dibuka ulang di Excel (mencegah NIK berubah jadi tanggal)
from io import BytesIO                  # Untuk ekspor data ke Excel 
from reportlab.lib.pagesizes import A4  # Untuk ekspor data ke PDF 
from reportlab.lib import colors        # Untuk styling tabel PDF
from reportlab.lib.styles import getSampleStyleSheet    # Untuk styling teks PDF
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image   # Untuk membuat elemen PDF seperti tabel, paragraf, gambar, dll
from reportlab.lib.units import cm      # Untuk konversi ukuran ke cm saat membuat PDF
from pathlib import Path                # Digunakan untuk membuat dan mengelola path (lokasi file/folder) project

# ==========================================
# PATH CONFIGURATION
# ==========================================
BASE_DIR = Path(__file__).resolve().parent.parent

# Folder
SOURCE_CODE_DIR = BASE_DIR / "01_source_code"
DATASET_DIR = BASE_DIR / "03_dataset"
DATABASE_DIR = BASE_DIR / "04_database_operasional"
ASSET_DIR = BASE_DIR / "05_asset"

# Dataset
DATA_MASTER_PATH = DATASET_DIR / "employee" / "Data_Master.xlsx"

# Database Operasional
PEGAWAI_MASTER_PATH = DATABASE_DIR / "employee" / "pegawai_kelola.csv"
PEGAWAI_ARSIP_PATH = DATABASE_DIR / "employee" / "arsip_pegawai.csv"

KEHADIRAN_PATH = DATABASE_DIR / "attendance" / "kehadiran_final.csv"
KPI_PATH = DATABASE_DIR / "performance" / "kpi_7aspek_final.csv"

LOG_AKTIVITAS_PATH = DATABASE_DIR / "activity" / "log_aktivitas.csv"  # sudah tidak dipakai: history activity kini disimpan di session_state (in-memory), bukan file

# Asset
LOGO_PATH = ASSET_DIR / "LOGO.png"

# =====================================================================
# KONFIGURASI HALAMAN STREAMLIT
# =====================================================================
# set_page_config harus dipanggil PALING AWAL sebelum elemen lain.
st.set_page_config(
    page_title="DASHBOARD KEPEGAWAIAN BIRO UMUM",   # Judul tab browser
    layout="wide",                        # Tampilan lebar (full width)
    initial_sidebar_state="expanded"      # Sidebar otomatis terbuka
)



# =====================================================================
# STYLING (CSS) - mempercantik tampilan dengan warna pemerintahan
# =====================================================================
st.markdown("""
    <style>        

    /* Import font Google */
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');

    /* Mengatur font global */
    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
    }

    /* Background utama */
    .stApp {
        background-color: #f4f7fb;
    }

    /* Kartu metrik (KPI Cards) */
    .kpi-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 20px;
        border-radius: 14px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .kpi-card h3 {
        font-size: 14px; font-weight: 500; margin: 0; opacity: 0.9;
    }
    .kpi-card h1 {
        font-size: 32px; font-weight: 700; margin: 6px 0 0 0;
    }

    /* Header judul halaman */
    .page-title {
        background: linear-gradient(90deg, #1e3c72, #2a5298);
        color: white;
        padding: 18px 24px;
        border-radius: 12px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(30,60,114,0.25);
    }
    .page-title h1 { margin: 0; font-size: 26px; font-weight: 700; }
    .page-title p  { margin: 4px 0 0 0; opacity: 0.85; font-size: 13px; }

    /* Tombol */
    .stButton>button {
        background: linear-gradient(90deg, #1e3c72, #2a5298);
        color: white; border: none; border-radius: 8px;
        padding: 8px 20px; font-weight: 600;
    }
    .stButton>button:hover {
        background: linear-gradient(90deg, #2a5298, #1e3c72);
        transform: translateY(-1px);
    }

    /* Label input di area konten utama */
    div[data-testid="stVerticalBlock"] label,
    div[data-testid="stVerticalBlock"] label p {
        color: #111111 !important;
    }

    /* Teks radio button di area konten utama */
    div[data-testid="stRadio"] label p {
        color: #111111 !important;
    }

    /* Text input seperti Cari nama */
    div[data-testid="stTextInput"] label,
    div[data-testid="stTextInput"] label p {
        color: #111111 !important;
    }

    /* Selectbox seperti Pilih pegawai / Pilih bulan */
    div[data-testid="stSelectbox"] label,
    div[data-testid="stSelectbox"] label p {
        color: #111111 !important;
    }

    /* File uploader label */
    div[data-testid="stFileUploader"] label,
    div[data-testid="stFileUploader"] label p {
        color: #111111 !important;
    }

    /* Date input label */
    div[data-testid="stDateInput"] label,
    div[data-testid="stDateInput"] label p {
        color: #111111 !important;
    }

    /* Header page-title tetap putih */
    .page-title,
    .page-title h1,
    .page-title p {
        color: white !important;
    }

    /* KPI card tetap putih */
    .kpi-card,
    .kpi-card h3,
    .kpi-card h1 {
        color: white !important;
    }    

    /* =========================================================
    SIDEBAR / NAVBAR PROFESIONAL
    ========================================================= */

    /* Sidebar utama */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #16294a 0%, #1e3c72 55%, #2a5298 100%) !important;
        box-shadow: 8px 0 24px rgba(0, 0, 0, 0.20) !important;
        border-right: none !important;
    }

    /* Padding sidebar */
    section[data-testid="stSidebar"] > div {
        padding: 20px 16px !important;
    }

    /* Semua teks sidebar */
    section[data-testid="stSidebar"] * {
        color: #eef4f8 !important;
    }

    /* Grup menu navigasi - tanpa jarak, dipisah garis tipis antar item */
    section[data-testid="stSidebar"] div[role="radiogroup"] {
        gap: 0 !important;
    }

    /* Item menu: flat, melebar penuh dari tepi kiri ke tepi kanan sidebar */
    section[data-testid="stSidebar"] label[data-baseweb="radio"] {
        background: transparent !important;
        border-radius: 0 !important;
        margin: 0 -16px !important;
        width: calc(100% + 32px) !important;
        padding: 13px 16px 13px 21px !important;
        border-top: none !important;
        border-right: none !important;
        border-bottom: 1px solid rgba(255, 255, 255, 0.08) !important;
        border-left: 3px solid transparent !important;
        transition: background 0.15s ease-in-out, border-color 0.15s ease-in-out !important;
    }

    /* Hover menu */
    section[data-testid="stSidebar"] label[data-baseweb="radio"]:hover {
        background: rgba(255, 255, 255, 0.06) !important;
    }

    /* Radio bulatan dirapikan */
    section[data-testid="stSidebar"] label[data-baseweb="radio"] > div:first-child {
        display: none !important;
    }

    /* Teks menu */
    section[data-testid="stSidebar"] label[data-baseweb="radio"] p {
        color: rgba(255, 255, 255, 0.94) !important;
        font-size: 13px !important;
        font-weight: 600 !important;
        letter-spacing: 0.6px !important;
        text-transform: uppercase !important;
    }

    /* Menu aktif: latar penuh dari kiri ke kanan + garis aksen di kiri */
    section[data-testid="stSidebar"] label[data-baseweb="radio"]:has(input:checked) {
        background: linear-gradient(90deg, rgba(47,128,237,0.24), rgba(47,128,237,0.03)) !important;
        border-left: 3px solid #56ccf2 !important;
        border-bottom: 1px solid rgba(255, 255, 255, 0.08) !important;
    }

    /* Teks menu aktif */
    section[data-testid="stSidebar"] label[data-baseweb="radio"]:has(input:checked) p {
        color: #ffffff !important;
        font-weight: 700 !important;
    }

    /* Garis pemisah */
    section[data-testid="stSidebar"] hr {
        border-color: rgba(255, 255, 255, 0.16) !important;
        margin: 18px 0 !important;
    }

    /* Caption footer */
    section[data-testid="stSidebar"] .stCaption,
    section[data-testid="stSidebar"] caption {
        color: rgba(238, 244, 248, 0.72) !important;
    }        
            
    """, unsafe_allow_html=True)

# =====================================================================
# =====================================================================
# DATABASE USER (untuk login)
# ---------------------------------------------------------------------
# Pada implementasi nyata sebaiknya disimpan di database / hashing.
# Di sini kita gunakan dictionary sederhana untuk demo.
# =====================================================================
# =====================================================================
USERS = {
    "TeamIT": {
        "password": "Ubicilembususu",
        "nama": "TeamIT",
        "role": "TEGAR RAMDHANI HERYN SYAPUTRA, S.Kom"
    },
    "pimpinan": {
        "password": "Susufullcream",
        "nama": "Pimpinan",
        "role": "FIRDAUS ARDISTYA HUTAMA, S.STP., M. AP"
    },
}


# =====================================================================
# =====================================================================
# DEFINISI VARIABLE DARI KOMPONEN KPI JAM KERJA + 7 ASPEK ASN
# =====================================================================
# =====================================================================
KPI_KOMPONEN = {
    "I":   {"label": "Ijin", "tarif": 0},
    "TL":  {"label": "Terlambat", "tarif": 35000},
    "PC":  {"label": "Pulang Cepat", "tarif": 35000},
    "TAD": {"label": "Tidak Absen Datang", "tarif": 35000},
    "TAP": {"label": "Tidak Absen Pulang", "tarif": 35000},
    "A":   {"label": "Alpha", "tarif": 35000},
    "SN":  {"label": "Tidak Ikut Senam", "tarif": 35000},
    "TSN": {"label": "Terlambat Senam", "tarif": 35000},
    "AP":  {"label": "Tidak Ikut Apel", "tarif": 35000},
}

POTONGAN_MAKSIMAL_HARIAN = 135000
POTONGAN_MAKSIMAL_BULANAN = 135000 * 22  # asumsi 22 hari kerja


ASPEK_7_PATH = DATABASE_DIR / "performance" / "kpi_7aspek_final.csv"

ASPEK_7_COLS = [
    "Orientasi_Pelayanan",
    "Akuntabel",
    "Kompeten",
    "Harmonis",
    "Loyal",
    "Adaptif",
    "Kolaboratif",
]



# =====================================================================
# KONFIGURASI FILE MASTER PEGAWAI
# =====================================================================

KOLOM_MINIMAL_PEGAWAI = [
    "NO",
    "NAMA",
    "NIP BARU",
    "NIK",

    "JK",
    "AGAMA",
    "STATUS PNS/CPNS",

    "JENJANG",

    "NAMA JABATAN",
    "UNIT KERJA",

    "TGL LAHIR",
    "USIA",

    "TMT MASUK KERJA",

    "TMT PANGKAT",
    # "GOL. RUANG" SENGAJA TIDAK didaftarkan di sini. Kolom ini sudah
    # digabung/direname menjadi "GOLONGAN" oleh normalisasi_nama_kolom_pegawai().
    # Jika didaftarkan di sini, normalisasi_data_pegawai() akan membuatnya
    # ulang sebagai kolom baru berisi "-" untuk SEMUA baris - itulah
    # penyebab kolom GOL. RUANG selalu tampak kosong walau GOLONGAN terisi.

    "TMT PENSIUN",
    "TAHUN PENSIUN",
]

def hapus_kolom_duplikat(df):
    """
    Menggabungkan kolom-kolom yang memiliki nama identik menjadi satu kolom.

    Duplikasi nama kolom dapat muncul setelah proses rename (mis. dua kolom
    sumber berbeda dipetakan ke nama target yang sama). Pandas sendiri tidak
    melarang DataFrame memiliki nama kolom kembar, tetapi proses hilir
    (st.dataframe -> pyarrow.Table.from_pandas) akan menolaknya dengan
    'ValueError: Duplicate column names found'. Fungsi ini menjadi lapisan
    pengaman terakhir: untuk setiap nama kolom yang kembar, nilai non-kosong
    pertama yang ditemukan pada baris tersebut dipertahankan, kolom
    selebihnya dibuang.
    """
    if not df.columns.duplicated().any():
        return df

    df = df.copy()
    hasil = pd.DataFrame(index=df.index)
    sudah_diproses = set()

    for kolom in df.columns:
        if kolom in sudah_diproses:
            continue
        sudah_diproses.add(kolom)

        subset = df.loc[:, df.columns == kolom]
        if subset.shape[1] == 1:
            hasil[kolom] = subset.iloc[:, 0]
            continue

        gabungan = subset.iloc[:, 0]
        for i in range(1, subset.shape[1]):
            kosong = gabungan.isna() | (
                gabungan.astype(str).str.strip().isin(["", "-", "nan", "None"])
            )
            gabungan = gabungan.where(~kosong, subset.iloc[:, i])
        hasil[kolom] = gabungan

    return hasil


def normalisasi_nama_kolom_pegawai(df):
    """
    Menyamakan variasi nama kolom dari Data_Master.xlsx / pegawai_kelola.csv
    agar cocok dengan KOLOM_MINIMAL_PEGAWAI, apapun format aslinya.
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # "GOL. RUANG" ditangani terpisah (sebelum rename_map umum) karena
    # berpotensi bentrok dengan kolom "GOLONGAN" yang sudah ada pada file
    # sumber. Jika keduanya hadir sekaligus, keduanya digabung: nilai
    # "GOLONGAN" yang sudah terisi dipertahankan, baris yang kosong diisi
    # dari "GOL. RUANG". Ini mencegah munculnya dua kolom "GOLONGAN".
    if "GOL. RUANG" in df.columns:
        if "GOLONGAN" in df.columns:
            gol_kosong = df["GOLONGAN"].isna() | (
                df["GOLONGAN"].astype(str).str.strip().isin(["", "-", "nan", "None"])
            )
            df.loc[gol_kosong, "GOLONGAN"] = df.loc[gol_kosong, "GOL. RUANG"]
            df = df.drop(columns=["GOL. RUANG"])
        else:
            df = df.rename(columns={"GOL. RUANG": "GOLONGAN"})

    rename_map = {
        "Nama Pegawai": "NAMA",
        "Nama": "NAMA",
        "NIP": "NIP BARU",
        "Jenis Kelamin": "JK",
        "Tanggal Lahir": "TGL LAHIR",
        "Status": "STATUS PNS/CPNS",
        "Jenjang Pendidikan": "JENJANG",
        "Jabatan": "NAMA JABATAN",
        "Unit Kerja": "UNIT KERJA",
        "TMT Kerja": "TMT MASUK KERJA",
        "TMT Pangkat": "TMT PANGKAT",
    }

    df = df.rename(columns=rename_map)

    # Lapisan pengaman umum: kolom lain (mis. "Nama" & "Nama Pegawai" yang
    # hadir bersamaan) juga dapat memicu duplikasi setelah rename_map di atas.
    df = hapus_kolom_duplikat(df)

    return df

# =====================================================================
# =====================================================================
# FUNGSI KATEGORI KPI
#====================================================================
# =====================================================================
def kategori_jam_kerja(nilai):
    if nilai < 75:
        return "Dibawah Ekspektasi"
    elif nilai <= 90:
        return "Sesuai Ekspektasi"
    else:
        return "Diatas Ekspektasi"


def kategori_7_aspek(nilai):
    if nilai < 75:
        return "Dibawah Ekspektasi"
    elif nilai <= 90:
        return "Sesuai Ekspektasi"
    else:
        return "Diatas Ekspektasi"


# =====================================================================
# =====================================================================
# FUNGSI PELINDUNG DATA: NIK & KOLOM TANGGAL
# ---------------------------------------------------------------------
# Ditambahkan untuk memperbaiki 2 bug:
# 1. NIK berubah menjadi format tanggal (mis. "1971-10-02 00:00:00")
#    karena sel di sumber data (Excel/CSV) terbaca/tersimpan sebagai
#    tipe Tanggal, bukan teks/angka.
# 2. TGL LAHIR (dan turunannya, USIA) hilang menjadi "-" karena
#    pd.to_datetime() sebelumnya dipanggil tanpa dayfirst=True dan
#    tanpa format cadangan, sehingga baris dengan format teks yang
#    sedikit berbeda langsung gagal parse.
# =====================================================================
def bersihkan_nik(value):
    """
    Memastikan NIK selalu berupa teks digit murni.

    PENTING: jika NIK di sumber data sudah terlanjur rusak (kolomnya
    terbaca sebagai objek Tanggal/Timestamp), angka NIK yang asli
    sudah tidak bisa direkonstruksi ulang dari kode ini -- datanya
    sudah tertimpa oleh nilai tanggal. Karena itu fungsi ini TIDAK
    menebak angka NIK, melainkan menandainya dengan jelas supaya
    bug langsung terlihat oleh admin dan bisa dicek ulang ke sumber
    data, alih-alih diam-diam menampilkan tanggal yang menyesatkan.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "-"

    # Kasus 1: nilai asli berupa objek tanggal/waktu Python/pandas
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return "NIK TIDAK VALID (terbaca sebagai tanggal)"

    text = str(value).strip()

    if text in ["", "-", "nan", "NaT", "None"]:
        return "-"

    # Kasus 2: NIK sudah tersimpan sebagai teks tanggal, mis. hasil
    # str(Timestamp) => "1971-10-02 00:00:00" atau "1971-10-02"
    if re.match(r"^\d{4}-\d{2}-\d{2}(\s\d{2}:\d{2}:\d{2})?$", text):
        return "NIK TIDAK VALID (terbaca sebagai tanggal)"

    # Kasus 3: NIK terbaca sebagai float / notasi ilmiah,
    # mis. "35300210711824.0" atau "3.53002107118e+13"
    try:
        if re.match(r"^-?\d+(\.\d+)?[eE][+-]?\d+$", text) or text.endswith(".0"):
            text = str(int(float(text)))
    except (ValueError, OverflowError):
        pass

    return text.strip()


def parse_kolom_tanggal(series):
    """
    Mem-parsing satu kolom tanggal dengan beberapa kemungkinan format
    sekaligus (bukan hanya satu format tunggal), supaya tanggal yang
    sebenarnya valid tidak ikut hilang menjadi NaT/"-" hanya karena
    variasi format teks antar baris (mis. campuran "05-09-1993" dan
    "1993-09-05" dalam satu kolom).
    """
    if series is None or len(series) == 0:
        return series

    # Format utama: DD-MM-YYYY (dayfirst) -- format standar dashboard ini.
    hasil = pd.to_datetime(series, errors="coerce", dayfirst=True)

    # Untuk baris yang masih gagal, coba beberapa format eksplisit lain
    # sebelum benar-benar menyerah dan menandainya "-".
    if hasil.isna().any():
        gagal = hasil.isna()
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
            if not gagal.any():
                break
            coba = pd.to_datetime(series[gagal], format=fmt, errors="coerce")
            hasil.loc[gagal] = hasil.loc[gagal].where(hasil.loc[gagal].notna(), coba)
            gagal = hasil.isna()

    return hasil


# =====================================================================
# =====================================================================
# FUNGSI: Memuat data pegawai dari EXCEL
# =====================================================================
# =====================================================================
@st.cache_data
def load_data(path: Path = DATA_MASTER_PATH) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(
            path,
            dtype={
                "NIP BARU": str,
                "NIK": str
            }
        )    
        df["NIK"] = df["NIK"].apply(bersihkan_nik)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()

    df.columns = [str(c).strip() for c in df.columns]
    df = normalisasi_nama_kolom_pegawai(df)
    df = hapus_kolom_duplikat(df)

    kolom_tanggal = [
        "TGL LAHIR",
        "TMT MASUK KERJA",
        "TMT PENSIUN",
        "TMT PANGKAT"
    ]

    for kol in kolom_tanggal:
        if kol in df.columns:
            df[kol] = (
                parse_kolom_tanggal(df[kol])
                .dt.strftime("%d-%m-%Y")
                .fillna("-")
            )

    df = df.fillna("-")   

    return df

def parse_usia(value):
    """Parse usia dari teks seperti '55 Tahun' menjadi angka."""
    if pd.isna(value):
        return np.nan
    if isinstance(value, str):
        match = re.search(r"(\d+)", value)
        return float(match.group(1)) if match else np.nan
    if isinstance(value, (int, float)):
        return float(value)
    return np.nan


def parse_tanggal_pegawai(value):
    """
    Mengubah teks tanggal dari CSV menjadi format tanggal Python.
    Digunakan untuk membaca tanggal lahir, TMT masuk kerja, dan TMT pensiun.
    """
    if pd.isna(value) or str(value).strip() in ["", "-", "--", "nan"]:
        return pd.NaT

    # Jika format CSV kamu hari/bulan/tahun, gunakan dayfirst=True
    return pd.to_datetime(value, errors="coerce", dayfirst=True)


def hitung_usia_real_time(tanggal_lahir):
    """
    Menghitung usia pegawai secara real-time
    dari tanggal lahir sampai tanggal hari ini.
    """
    tanggal_lahir = parse_tanggal_pegawai(tanggal_lahir)

    if pd.isna(tanggal_lahir):
        return "-"

    hari_ini = date.today()
    selisih = relativedelta(hari_ini, tanggal_lahir.date())

    return f"{selisih.years} Tahun"


def ekstrak_tanggal_lahir(tanggal_lahir):
    """
    Memecah tanggal lahir menjadi tanggal, bulan, dan tahun lahir.
    """
    tgl = parse_tanggal_pegawai(tanggal_lahir)

    if pd.isna(tgl):
        return "-", "-", "-"

    bulan_indo = {
        1: "Januari",
        2: "Februari",
        3: "Maret",
        4: "April",
        5: "Mei",
        6: "Juni",
        7: "Juli",
        8: "Agustus",
        9: "September",
        10: "Oktober",
        11: "November",
        12: "Desember",
    }

    return tgl.day, bulan_indo[tgl.month], tgl.year


def hitung_durasi_tanggal(tanggal_awal, tanggal_akhir=None):
    """
    Menghitung durasi dari tanggal_awal sampai tanggal_akhir.
    Output: X Tahun Y Bulan Z Hari.
    Digunakan untuk:
    1. Lama masa kerja
    2. Durasi menuju pensiun
    """
    if tanggal_akhir is None:
        tanggal_akhir = date.today()

    tanggal_awal = parse_tanggal_pegawai(tanggal_awal)

    if pd.isna(tanggal_awal):
        return "-"

    tanggal_awal = tanggal_awal.date()

    # Jika tanggal awal lebih besar dari tanggal akhir, berarti waktunya sudah lewat
    if tanggal_awal > tanggal_akhir:
        return "0 Tahun 0 Bulan 0 Hari"

    selisih = relativedelta(tanggal_akhir, tanggal_awal)

    return f"{selisih.years} Tahun {selisih.months} Bulan {selisih.days} Hari"


def hitung_menu1_durasi(df):
    """
    Menambahkan kolom real-time:
    1. Lama Masa Kerja: dihitung dari TMT MASUK KERJA sampai hari ini.
    2. Durasi Menuju Pensiun: dihitung dari hari ini sampai TMT PENSIUN.
    """
    data = df.copy()
    hari_ini = date.today()

    if "TGL LAHIR" in data.columns:
        data["USIA"] = data["TGL LAHIR"].apply(hitung_usia_real_time)
    else:
        data["USIA"] = "-"

    if "TMT MASUK KERJA" in data.columns:
        data["LAMA MASA KERJA"] = data["TMT MASUK KERJA"].apply(
            lambda x: hitung_durasi_tanggal(x, hari_ini)
        )
    else:
        data["LAMA MASA KERJA"] = "-"

    if "TMT PENSIUN" in data.columns:
        data["DURASI MENUJU PENSIUN"] = data["TMT PENSIUN"].apply(
            lambda x: hitung_durasi_tanggal(hari_ini, parse_tanggal_pegawai(x).date())
            if not pd.isna(parse_tanggal_pegawai(x))
            else "-"
        )
    else:
        data["DURASI MENUJU PENSIUN"] = "-"

    return data


def hitung_masa_pangkat_golongan(df):
    """
    Menghitung masa pangkat dan masa golongan secara real-time
    berdasarkan TMT Pangkat.
    """

    data = df.copy()
    hari_ini = pd.Timestamp.today().normalize()

    # ==========================
    # MASA PANGKAT
    # ==========================
    if "TMT PANGKAT" in data.columns:

        data["TMT_PANGKAT_DATE"] = data["TMT PANGKAT"].apply(
            parse_tanggal_pegawai
        )

        data["MASA_PANGKAT"] = data["TMT_PANGKAT_DATE"].apply(
            lambda x:
            f"{relativedelta(hari_ini, x).years} Tahun "
            f"{relativedelta(hari_ini, x).months} Bulan"
            if pd.notna(x)
            else "-"
        )

        data["MASA_PANGKAT_TAHUN"] = data["TMT_PANGKAT_DATE"].apply(
            lambda x:
            relativedelta(hari_ini, x).years
            if pd.notna(x)
            else np.nan
        )

    else:
        data["MASA_PANGKAT"] = "-"
        data["MASA_PANGKAT_TAHUN"] = np.nan

    return data

    

# =====================================================================
# =====================================================================
# FUNGSI CLEANING & PREPROCESSING KEHADIRAN REAL
# =====================================================================
# =====================================================================
STATUS_MAP = {
    "": "Hadir",
    "LIB": "Libur",
    "A1": "Alpha",
    "I1": "Ijin",
    "C1": "Cuti",
    "S1": "Sakit",
    "DL1": "Dinas Luar",
    "TAD1": "Tidak Absen Datang",
    "TAP1": "Tidak Absen Pulang",
    "TSN1": "Terlambat Senam",
    "SN1": "Tidak Senam",
    "AP1": "Tidak Apel",
}

POTONGAN_MAP = {
    "A1": 35000,
    "I1": 0,
    "C1": 0,
    "S1": 0,
    "DL1": 0,
    "TAD1": 35000,
    "TAP1": 35000,
    "TSN1": 35000,
    "SN1": 35000,
    "AP1": 35000,
}

POTONGAN_TERLAMBAT_PER_10MENIT = 35000

_RE_TL = re.compile(r"TL(\d+)M")
_RE_PC = re.compile(r"PC(\d+)M")

BULAN_ID = {
    1: "Januari",
    2: "Februari",
    3: "Maret",
    4: "April",
    5: "Mei",
    6: "Juni",
    7: "Juli",
    8: "Agustus",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Desember",
}
NAMA_KE_NOMOR = {v.upper(): k for k, v in BULAN_ID.items()}


def bersihkan_nip(v):
    """Menghapus karakter selain angka pada NIP."""
    return re.sub(r"[^0-9]", "", "" if v is None else str(v))


def bulan_dari_nama_file(filename):
    """
    Deteksi bulan dari nama file.
    Contoh:
    SKP KEHADIRAN JANUARI.xlsx -> 1
    """
    base = os.path.basename(filename).upper()

    for nama, nomor in NAMA_KE_NOMOR.items():
        if nama in base:
            return nomor

    return None


def parse_cell(cell):
    """
    Parsing isi cell harian.
    Contoh:
    - kosong / NaN -> Hadir
    - LIB -> Libur
    - A1 -> Alpha
    - TAD1 -> Tidak Absen Datang
    - TAP1 -> Tidak Absen Pulang
    - TL66M -> Terlambat 66 menit
    """

    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return "Hadir", 0, 0

    s = str(cell).strip().upper()

    if not s:
        return "Hadir", 0, 0

    tokens = s.split()

    menit_terlambat = sum(
        int(m.group(1))
        for tk in tokens
        for m in [_RE_TL.fullmatch(tk)]
        if m
    )

    menit_pulang_cepat = sum(
        int(m.group(1))
        for tk in tokens
        for m in [_RE_PC.fullmatch(tk)]
        if m
    )

    prioritas = [
        "LIB",
        "C1",
        "I1",
        "S1",
        "DL1",
        "A1",
        "TAD1",
        "TAP1",
        "TSN1",
        "AP1",
        "SN1",
    ]

    status = next(
        (STATUS_MAP.get(c, c) for c in prioritas if c in tokens),
        None,
    )

    if status is None:
        if menit_terlambat > 0:
            status = "Terlambat"
        elif menit_pulang_cepat > 0:
            status = "Pulang Cepat"
        else:
            status = "Hadir"

    return status, menit_terlambat, menit_pulang_cepat


def hitung_potongan(kode_asli, menit_terlambat=0, menit_pulang_cepat=0):
    if kode_asli is None or (isinstance(kode_asli, float) and pd.isna(kode_asli)):
        return 0

    s = str(kode_asli).strip().upper()

    if not s:
        return 0

    tokens = s.split()
    total_potongan = 0

    for kode, nominal in POTONGAN_MAP.items():
        if kode in tokens:
            total_potongan += nominal

    if menit_terlambat > 0:
        total_potongan += (menit_terlambat // 10) * POTONGAN_TERLAMBAT_PER_10MENIT

    if menit_pulang_cepat > 0:
        total_potongan += (menit_pulang_cepat // 10) * 35000

    return min(total_potongan, POTONGAN_MAKSIMAL_HARIAN)

def read_raw_file(source, filename, header=None):
    """
    Membaca file upload atau file lokal.
    Mendukung .xlsx, .xls, dan .csv.
    """

    filename_lower = filename.lower()

    if hasattr(source, "getvalue"):
        data = BytesIO(source.getvalue())
    else:
        data = source

    if filename_lower.endswith(".csv"):
        return pd.read_csv(data, header=header, dtype=object)

    return pd.read_excel(data, sheet_name=0, header=header, dtype=object)


def bersihkan_file_kehadiran(source, filename):
    """
    Membersihkan 1 file raw kehadiran menjadi format harian:
    nip, nama, opd, tanggal, status_hadir, menit_terlambat, potong_gaji, kode_asli
    """

    raw = read_raw_file(source, filename, header=None)

    header_row = next(
        (
            i
            for i in range(min(10, len(raw)))
            if str(raw.iat[i, 0]).strip().lower() == "no"
        ),
        None,
    )

    if header_row is None:
        raise ValueError(f"Header 'No' tidak ditemukan pada file: {filename}")

    df = read_raw_file(source, filename, header=header_row)

    required_cols = ["NIP", "Nama Pegawai", "OPD"]

    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Kolom wajib '{col}' tidak ditemukan pada file: {filename}")

    df = df.dropna(subset=["NIP", "Nama Pegawai"])

    df["NIP"] = df["NIP"].map(bersihkan_nip)
    df["Nama"] = df["Nama Pegawai"].astype(str).str.strip()
    df["OPD"] = df["OPD"].astype(str).str.strip()

    bulan_file = bulan_dari_nama_file(filename)

    if bulan_file is not None:
        df["Bulan_Clean"] = bulan_file
    elif "Bulan" in df.columns:
        df["Bulan_Clean"] = pd.to_numeric(df["Bulan"], errors="coerce")
    else:
        raise ValueError(
            f"Bulan tidak bisa dideteksi dari nama file atau kolom Bulan pada file: {filename}"
        )

    if "Tahun" in df.columns:
        df["Tahun_Clean"] = pd.to_numeric(df["Tahun"], errors="coerce")
    else:
        df["Tahun_Clean"] = datetime.now().year

    day_cols = [
        c
        for c in df.columns
        if str(c).zfill(2) in {f"{i:02d}" for i in range(1, 32)}
    ]

    if not day_cols:
        raise ValueError(f"Kolom tanggal 01-31 tidak ditemukan pada file: {filename}")

    long_df = df.melt(
        id_vars=[
            "NIP",
            "Nama",
            "OPD",
            "Bulan_Clean",
            "Tahun_Clean",
        ],
        value_vars=day_cols,
        var_name="hari",
        value_name="kode",
    )

    long_df["hari"] = pd.to_numeric(long_df["hari"], errors="coerce")
    long_df = long_df.dropna(subset=["hari", "Bulan_Clean", "Tahun_Clean"])

    def buat_tanggal(row):
        try:
            return date(
                int(row["Tahun_Clean"]),
                int(row["Bulan_Clean"]),
                int(row["hari"]),
            )
        except ValueError:
            return pd.NaT

    long_df["tanggal"] = long_df.apply(buat_tanggal, axis=1)
    long_df = long_df.dropna(subset=["tanggal"])

    parsed = long_df["kode"].map(parse_cell)

    long_df["status_hadir"] = [p[0] for p in parsed]
    long_df["menit_terlambat"] = [p[1] for p in parsed]
    long_df["menit_pulang_cepat"] = [p[2] for p in parsed]

    long_df["potong_gaji"] = long_df.apply(
        lambda row: hitung_potongan(
            row["kode"],
            row["menit_terlambat"],
            row["menit_pulang_cepat"],
        ),
        axis=1,
    )

    harian = (
        long_df[
            [
                "NIP",
                "Nama",
                "OPD",
                "tanggal",
                "status_hadir",
                "menit_terlambat",
                "potong_gaji",
                "kode",
            ]
        ]
        .rename(
            columns={
                "NIP": "nip",
                "Nama": "nama",
                "OPD": "opd",
                "kode": "kode_asli",
            }
        )
        .drop_duplicates(subset=["nip", "tanggal"], keep="last")
        .sort_values(["tanggal", "nip"])
        .reset_index(drop=True)
    )

    return harian


def proses_upload_kehadiran(uploaded_files, output_path=KEHADIRAN_PATH):
    """
    Memproses file upload dan MENAMBAHKAN ke data lama.

    Alur:
    1. Baca data lama dari kehadiran_final.csv jika sudah ada.
    2. Proses file raw baru yang di-upload.
    3. Gabungkan data lama + data baru.
    4. Jika ada nip + tanggal yang sama, pakai data terbaru.
    5. Simpan kembali ke kehadiran_final.csv.
    """

    all_data = []

    # 1. Baca data lama jika file kehadiran_final.csv sudah ada
    if output_path.exists():
        old_df = pd.read_csv(output_path, dtype={"nip": str})

        if "tanggal" in old_df.columns:
            old_df["tanggal"] = pd.to_datetime(
                old_df["tanggal"],
                errors="coerce"
            )

        all_data.append(old_df)

    # 2. Proses file baru yang di-upload
    for uploaded_file in uploaded_files:
        df_harian = bersihkan_file_kehadiran(uploaded_file, uploaded_file.name)

        df_harian["tanggal"] = pd.to_datetime(
            df_harian["tanggal"],
            errors="coerce"
        )

        all_data.append(df_harian)

    if not all_data:
        return pd.DataFrame()

    # 3. Gabungkan data lama + data baru
    final_df = pd.concat(all_data, ignore_index=True)

    # 4. Hapus duplikat.
    # Jika bulan/tanggal yang sama di-upload ulang,
    # data upload terbaru akan menggantikan data lama.
    final_df = (
        final_df
        .drop_duplicates(subset=["nip", "tanggal"], keep="last")
        .sort_values(["tanggal", "nip"])
        .reset_index(drop=True)
    )

    # 5. Simpan ulang sebagai database final
    final_df.to_csv(output_path, index=False, encoding="utf-8-sig")

    return final_df


@st.cache_data
def load_kehadiran_final(path: Path = KEHADIRAN_PATH):
    """
    Membaca database KPI Kehadiran.
    Jika file belum ada atau masih kosong,
    mengembalikan DataFrame kosong.
    """

    if not path.exists():
        return pd.DataFrame()

    try:
        df = pd.read_csv(path, dtype={"nip": str})
    except (pd.errors.EmptyDataError, FileNotFoundError):
        return pd.DataFrame()

    if "tanggal" in df.columns:
        df["tanggal"] = pd.to_datetime(df["tanggal"], errors="coerce")

    return df


def kode_contains(series, kode):
    return series.fillna("").astype(str).str.upper().str.split().apply(lambda x: kode in x)

# Potongan maksimal bulanan mengikuti konstanta di dashboard
def ubah_kehadiran_ke_kpi(df_kehadiran):
    """
    Mengubah data harian kehadiran_final.csv menjadi format KPI Menu Jam Kerja.
    Potongan maksimal dihitung berdasarkan jumlah hari kerja aktual dalam periode data.
    """

    if df_kehadiran.empty:
        return pd.DataFrame()

    df = df_kehadiran.copy()

    df["tanggal"] = pd.to_datetime(df["tanggal"], errors="coerce")
    df = df.dropna(subset=["tanggal"])

    df["kode_asli"] = df["kode_asli"].fillna("").astype(str).str.upper()
    df["potong_gaji"] = pd.to_numeric(df["potong_gaji"], errors="coerce").fillna(0)
    df["menit_terlambat"] = pd.to_numeric(df["menit_terlambat"], errors="coerce").fillna(0)

    
    # Tanggal merah tetap dianggap masuk jika jatuh pada Senin-Jumat.
    # Kode LIB tetap boleh ada, tetapi tidak mengurangi hari kerja.
    df_kerja = df[df["tanggal"].dt.weekday < 5].copy()


    if df_kerja.empty:
        return pd.DataFrame()

    # Hitung komponen harian
    df_kerja["K_I"] = (
        kode_contains(df_kerja["kode_asli"], "I1")
        | kode_contains(df_kerja["kode_asli"], "C1")
        | kode_contains(df_kerja["kode_asli"], "S1")
        | kode_contains(df_kerja["kode_asli"], "DL1")
    ).astype(int)

    df_kerja["K_TL"] = (df_kerja["menit_terlambat"] // 10).astype(int)
    df_kerja["K_PC"] = (
        df_kerja["kode_asli"]
        .str.extract(r"PC(\d+)M", expand=False)
        .fillna(0)
        .astype(int)
        // 10
    )

    df_kerja["K_TAD"] = kode_contains(df_kerja["kode_asli"], "TAD1").astype(int)
    df_kerja["K_TAP"] = kode_contains(df_kerja["kode_asli"], "TAP1").astype(int)
    df_kerja["K_A"] = kode_contains(df_kerja["kode_asli"], "A1").astype(int)
    df_kerja["K_SN"] = kode_contains(df_kerja["kode_asli"], "SN1").astype(int)
    df_kerja["K_TSN"] = kode_contains(df_kerja["kode_asli"], "TSN1").astype(int)
    df_kerja["K_AP"] = kode_contains(df_kerja["kode_asli"], "AP1").astype(int)

    agg_cols = [
        "K_I",
        "K_TL",
        "K_PC",
        "K_TAD",
        "K_TAP",
        "K_A",
        "K_SN",
        "K_TSN",
        "K_AP",
    ]

    kpi = (
        df_kerja.groupby(["nip", "nama"], as_index=False)
        .agg(
            OPD=("opd", "first"),
            Tanggal_Awal=("tanggal", "min"),
            Tanggal_Akhir=("tanggal", "max"),
            Hari_Data=("tanggal", "nunique"),
            Potongan_Total=("potong_gaji", "sum"),
            **{col: (col, "sum") for col in agg_cols},
        )
    )

    kpi = kpi.rename(columns={"nama": "NAMA"})

    # Hari kerja aktual per pegawai dalam periode data
    kpi["Potongan_Maksimal_Periode"] = (
        kpi["Hari_Data"] * POTONGAN_MAKSIMAL_HARIAN
    )

    kpi["Skor_Jam_Kerja"] = (
        100 - (
            kpi["Potongan_Total"]
            / kpi["Potongan_Maksimal_Periode"].replace(0, np.nan)
            * 100
        )
    ).fillna(100).clip(lower=0, upper=100).round(1)

    kpi["Kategori_Jam_Kerja"] = kpi["Skor_Jam_Kerja"].apply(kategori_jam_kerja)

    kpi["Jam_Kerja_Aktual"] = np.nan
    kpi["Jam_Kerja_Target"] = np.nan
    kpi["Persen_Jam_Kerja"] = kpi["Skor_Jam_Kerja"]

    return kpi

# =====================================================================
# =====================================================================
# FUNGSI: untuk KPI 7 ASPEK ASN
# =====================================================================
# =====================================================================
##------fungsi untuk mendeteksi periode data 7 aspek berdasarkan nama file yang diupload------
def periode_dari_nama_file_7aspek(filename):
    """
    Fungsi ini digunakan untuk mendeteksi periode data 7 Aspek ASN
    berdasarkan nama file yang diupload.

    Contoh nama file:
    - 7 ASPEK JANUARI 2026.xlsx  -> 2026-01
    - 7 ASPEK FEBRUARI 2026.csv  -> 2026-02
    - DATA 7 ASPEK APRIL.xlsx    -> 2026-04

    Jika nama file tidak mengandung tahun, maka tahun otomatis
    memakai tahun saat aplikasi dijalankan.
    """

    # Ubah nama file menjadi huruf besar agar pencarian bulan lebih mudah
    base = os.path.basename(filename).upper()

    # Default tahun memakai tahun saat ini
    tahun = datetime.now().year

    # Cari apakah nama file mengandung tahun, misalnya 2026
    tahun_match = re.search(r"(20\d{2})", base)

    # Jika ditemukan tahun di nama file, gunakan tahun tersebut
    if tahun_match:
        tahun = int(tahun_match.group(1))

    # Cek apakah nama bulan Indonesia ada di nama file
    for nama_bulan, nomor_bulan in NAMA_KE_NOMOR.items():
        if nama_bulan in base:
            return f"{tahun}-{nomor_bulan:02d}"

    # Jika nama bulan tidak ditemukan, gunakan bulan saat ini sebagai cadangan
    return datetime.now().strftime("%Y-%m")

##------fungsi untuk normalisasi nama kolom 7 aspek agar konsisten meskipun format file berbeda-beda------
def normalisasi_nama_kolom_7aspek(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {
        "Nama": "NAMA",
        "nama": "NAMA",
        "Nip": "NIP",
        "nip": "NIP",
        "NIP BARU": "NIP",
        "Rata 7 Aspek": "Rata_7_Aspek",
        "Rata-rata": "Rata_7_Aspek",
    }

    df = df.rename(columns=rename_map)
    df = hapus_kolom_duplikat(df)
    return df

##------fungsi untuk membaca file 7 aspek, membersihkan data, dan menambahkan kolom perhitungan------
def baca_file_7aspek(source, filename):
    filename_lower = filename.lower()

    if hasattr(source, "getvalue"):
        data = BytesIO(source.getvalue())
    else:
        data = source

    if filename_lower.endswith(".csv"):
        df = pd.read_csv(data, dtype=str)
    else:
        df = pd.read_excel(data, dtype=str)

    df = normalisasi_nama_kolom_7aspek(df)

    if "NAMA" not in df.columns:
        raise ValueError("File wajib memiliki kolom NAMA.")

    for col in ASPEK_7_COLS:
        if col not in df.columns:
            raise ValueError(f"Kolom wajib '{col}' tidak ditemukan.")

    if "NIP" not in df.columns:
        df["NIP"] = "-"

    if "Periode" not in df.columns:
        df["Periode"] = periode_dari_nama_file_7aspek(filename)

    df["NAMA"] = df["NAMA"].astype(str).str.strip().str.upper()
    df["NIP"] = df["NIP"].astype(str).str.strip()

    for col in ASPEK_7_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).clip(0, 100)

    df["Rata_7_Aspek"] = df[ASPEK_7_COLS].mean(axis=1).round(1)
    df["Kategori_7_Aspek"] = df["Rata_7_Aspek"].apply(kategori_7_aspek)
    df["Waktu_Update"] = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    kolom_final = [
        "NIP",
        "NAMA",
        "Periode",
        *ASPEK_7_COLS,
        "Rata_7_Aspek",
        "Kategori_7_Aspek",
        "Waktu_Update",
    ]

    return df[kolom_final]

##------fungsi untuk memproses upload file 7 aspek, menggabungkan dengan data lama, dan menyimpan hasil akhir------
def proses_upload_7aspek(uploaded_files, output_path=ASPEK_7_PATH):
    all_data = []

    if os.path.exists(output_path):
        old_df = pd.read_csv(output_path, dtype=str).fillna("-")
        old_df = normalisasi_nama_kolom_7aspek(old_df)

        for col in ASPEK_7_COLS + ["Rata_7_Aspek"]:
            if col in old_df.columns:
                old_df[col] = pd.to_numeric(old_df[col], errors="coerce").fillna(0)

        all_data.append(old_df)

    for uploaded_file in uploaded_files:
        df_baru = baca_file_7aspek(uploaded_file, uploaded_file.name)
        all_data.append(df_baru)

    if not all_data:
        return pd.DataFrame()

    final_df = pd.concat(all_data, ignore_index=True)

    final_df["NAMA"] = final_df["NAMA"].astype(str).str.strip().str.upper()
    final_df["NIP"] = final_df["NIP"].astype(str).str.strip()
    final_df["Periode"] = final_df["Periode"].astype(str).str.strip()

    final_df = (
        final_df
        .drop_duplicates(subset=["NIP", "NAMA", "Periode"], keep="last")
        .sort_values(["Periode", "NAMA"])
        .reset_index(drop=True)
    )

    final_df.to_csv(output_path, index=False, encoding="utf-8-sig")

    return final_df

##------fungsi untuk memuat data 7 aspek dari file akhir------
@st.cache_data
def load_7aspek_final(path: Path = ASPEK_7_PATH):
    """
    Membaca database KPI 7 Aspek ASN.
    Jika file belum ada atau masih kosong,
    mengembalikan DataFrame kosong.
    """

    if not path.exists():
        return pd.DataFrame()

    try:
        df = pd.read_csv(path, dtype=str).fillna("-")
    except (pd.errors.EmptyDataError, FileNotFoundError):
        return pd.DataFrame()

    df = normalisasi_nama_kolom_7aspek(df)

    for col in ASPEK_7_COLS + ["Rata_7_Aspek"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    if "Kategori_7_Aspek" not in df.columns:
        df["Kategori_7_Aspek"] = df["Rata_7_Aspek"].apply(kategori_7_aspek)

    return df

# =====================================================================
# =====================================================================
# FUNGSI RAPORT PEGAWAI
# =====================================================================
# =====================================================================
def ambil_biodata_pegawai(df_pegawai, nama_pegawai):
    """
    Mengambil biodata pegawai dari data master.

    Perbaikan:
    1. Pencarian nama lebih fleksibel.
    2. Usia dihitung ulang secara real-time.
    3. Kolom tanggal diformat dd-mm-yyyy.
    4. NIK dan NIP dipastikan berupa string.
    """

    if df_pegawai.empty:
        return {}

    data = df_pegawai.copy()

    # ==========================
    # Normalisasi nama
    # ==========================
    data["NAMA_KEY"] = (
        data["NAMA"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    nama_dicari = str(nama_pegawai).upper().strip()

    hasil = data[data["NAMA_KEY"] == nama_dicari]

    # Jika tidak ketemu, gunakan pencarian contains
    if hasil.empty:
        hasil = data[
            data["NAMA_KEY"].str.contains(
                nama_dicari,
                case=False,
                na=False
            )
        ]

    if hasil.empty:
        return {}

    biodata = hasil.iloc[0].to_dict()

    # ======================================================
    # Hitung ulang usia secara realtime
    # ======================================================
    biodata["USIA"] = hitung_usia_real_time(
        biodata.get("TGL LAHIR")
    )

    # ======================================================
    # Format semua kolom tanggal
    # ======================================================
    tanggal_cols = [
        "TGL LAHIR",
        "TMT MASUK KERJA",
        "TMT PANGKAT",
        "TMT PENSIUN",
    ]

    for col in tanggal_cols:

        if col in biodata:

            try:
                tgl = pd.to_datetime(
                    biodata[col],
                    errors="coerce",
                    dayfirst=True
                )

                biodata[col] = (
                    tgl.strftime("%d-%m-%Y")
                    if pd.notna(tgl)
                    else "-"
                )

            except:
                biodata[col] = "-"

    # ======================================================
    # Pastikan NIK berupa teks
    # ======================================================
    if "NIK" in biodata:

        biodata["NIK"] = (
            str(biodata["NIK"])
            .replace(".0", "")
            .strip()
        )

    # ======================================================
    # Pastikan NIP berupa teks
    # ======================================================
    if "NIP BARU" in biodata:

        biodata["NIP BARU"] = (
            str(biodata["NIP BARU"])
            .replace(".0", "")
            .strip()
        )

    return biodata


def bersihkan_nama_matching(nama):
    """
    Membersihkan nama agar pencocokan antar data lebih mudah.

    Fungsi ini menghapus:
    - tanda baca
    - gelar yang menyebabkan nama berbeda antar file
    - spasi berlebih

    Tujuannya agar nama di DATA_PEGAWAI.csv dan kehadiran_final.csv
    tetap bisa cocok meskipun formatnya sedikit berbeda.
    """

    if nama is None:
        return ""

    nama = str(nama).upper()

    # Hapus tanda baca dan karakter selain huruf/angka/spasi
    nama = re.sub(r"[^A-Z0-9\s]", " ", nama)

    # Hapus gelar umum yang sering membuat nama tidak cocok
    daftar_gelar = [
        "S STP", "S SOS", "S KOM", "S IP", "S H", "S E", "S AP",
        "M AP", "M SI", "M M", "M HUM", "DR", "DRA", "DRS",
        "A MD", "A MD T", "A MD KOM"
    ]

    for gelar in daftar_gelar:
        nama = re.sub(rf"\b{gelar}\b", " ", nama)

    # Rapikan spasi ganda
    nama = re.sub(r"\s+", " ", nama).strip()

    return nama


def ambil_kpi_jam_pegawai(kpi_jam, nama_pegawai):
    """
    Mengambil ringkasan KPI Jam Kerja milik pegawai.
    memakai pencocokan nama yang lebih fleksibel, 
    sehingga nama dengan gelar atau tanda baca berbeda tetap bisa ditemukan.
    """

    if kpi_jam.empty:
        return {}

    data = kpi_jam.copy()

    # Buat nama bersih untuk data KPI jam kerja
    data["NAMA_MATCH"] = data["NAMA"].apply(bersihkan_nama_matching)

    # Buat nama bersih untuk pegawai yang dipilih
    nama_dicari = bersihkan_nama_matching(nama_pegawai)

    # 1. Coba cocokkan nama yang sudah dibersihkan secara persis
    hasil = data[data["NAMA_MATCH"] == nama_dicari]

    # 2. Jika belum ketemu, coba cari dengan metode contains dua arah
    if hasil.empty:
        hasil = data[
            data["NAMA_MATCH"].apply(
                lambda x: nama_dicari in x or x in nama_dicari
            )
        ]

    if hasil.empty:
        return {}

    return hasil.iloc[0].to_dict()

def ambil_kpi_7aspek_pegawai(kpi_7aspek, nama_pegawai):
    """
    Mengambil ringkasan KPI 7 Aspek BerAKHLAK milik pegawai.

    Jika pegawai memiliki beberapa bulan data,
    fungsi ini menghitung rata-rata seluruh periode.
    """

    if kpi_7aspek.empty:
        return {}

    data = kpi_7aspek.copy()
    data["NAMA_MATCH"] = data["NAMA"].apply(bersihkan_nama_matching)
    nama_dicari = bersihkan_nama_matching(nama_pegawai)

    hasil = data[data["NAMA_MATCH"] == nama_dicari]

    if hasil.empty:
        hasil = data[
            data["NAMA_MATCH"].apply(
                lambda x: nama_dicari in x or x in nama_dicari
        )
    ]

    if hasil.empty:
        return {}

    rata_aspek = hasil[ASPEK_7_COLS].mean().round(1)
    rata_total = round(rata_aspek.mean(), 1)

    periode_list = sorted(hasil["Periode"].astype(str).unique().tolist())

    if len(periode_list) == 1:
        periode = periode_list[0]
    else:
        periode = f"{periode_list[0]} s.d. {periode_list[-1]}"

    return {
        "Periode": periode,
        "Rata_7_Aspek": rata_total,
        "Kategori_7_Aspek": kategori_7_aspek(rata_total),
        **{col: rata_aspek[col] for col in ASPEK_7_COLS}
    }


def buat_pdf_raport_pegawai(biodata, kpi_jam, kpi_7aspek):
    """
    Membuat file PDF raport pegawai 1 halaman.

    Isi PDF:
    1. Judul raport
    2. Biodata pegawai
    3. Tabel KPI Kehadiran/Jam Kerja
    4. Tabel KPI 7 Aspek BerAKHLAK
    """

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
    )

    styles = getSampleStyleSheet()
    elements = []

    #----------HEADER RAPORT DENGAN LOGO----------#
    logo_path = LOGO_PATH

    if logo_path.exists():
        logo = Image(str(logo_path), width=3.2 * cm, height=2.2 * cm)
    else:
        logo = Paragraph("", styles["Normal"])

    header_text = Paragraph(
        """
        <para align="center">
            <b>RAPORT PEGAWAI</b><br/>
            <b>BIRO UMUM SEKRETARIAT DAERAH</b><br/>
            <b>PROVINSI JAWA TIMUR</b>
        </para>
        """,
        styles["Title"]
    )

    header_table = Table(
        [[logo, header_text]],
        colWidths=[4 * cm, 14 * cm]
    )

    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, 0), "CENTER"),
        ("ALIGN", (1, 0), (1, 0), "CENTER"),
        ("BOX", (0, 0), (-1, -1), 0, colors.white),
    ]))

    elements.append(header_table)
    elements.append(Spacer(1, 0.3 * cm))

    #----------TABEL BIODATA PEGAWAI-------------#
    biodata_table = [
        ["KETERANGAN", "ISI"],
        ["Nama", biodata.get("NAMA", "-")],
        ["NIP", biodata.get("NIP BARU", "-")],
        ["NIK", biodata.get("NIK", "-")],
        ["Jenis Kelamin", biodata.get("JK", "-")],
        ["Agama", biodata.get("AGAMA", "-")],
        ["Status Pegawai", biodata.get("STATUS PNS/CPNS", "-")],
        ["Jenjang", biodata.get("JENJANG", "-")],
        ["Jabatan", biodata.get("NAMA JABATAN", "-")],
        ["Unit Kerja", biodata.get("UNIT KERJA", "-")],
        ["Usia", biodata.get("USIA", "-")],
        ["TMT Masuk Kerja", biodata.get("TMT MASUK KERJA", "-")],
        ["TMT Pensiun", biodata.get("TMT PENSIUN", "-")],
    ]

    table_bio = Table(biodata_table, colWidths=[5 * cm, 13 * cm])
    table_bio.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),

    # Header abu-abu
    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),

    # Kolom kiri isi biodata
    ("BACKGROUND", (0, 1), (0, -1), colors.whitesmoke),

    # HEADER TENGAH
    ("ALIGN", (0, 0), (-1, 0), "CENTER"),

    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),

    ("FONTSIZE", (0, 0), (-1, -1), 8),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
]))


    #----------TABEL KPI KEHADIRAN/JAM KERJA-------------#
    elements.append(table_bio)
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(Paragraph("<b>RANGKUMAN KPI KEHADIRAN / JAM KERJA</b>", styles["Heading3"]))

    jam_table = [
        ["INDIKATOR", "NILAI"],
        ["Skor Jam Kerja", str(kpi_jam.get("Skor_Jam_Kerja", "-"))],
        ["Status", str(kpi_jam.get("Kategori_Jam_Kerja", "-"))],
        ["Jumlah Hari Kerja", str(kpi_jam.get("Hari_Data", "-"))],
        ["Total Potongan", f"Rp {int(kpi_jam.get('Potongan_Total', 0)):,}".replace(",", ".") if kpi_jam else "-"],
    ]

    table_jam = Table(jam_table, colWidths=[8 * cm, 10 * cm])
    table_jam.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),

    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),

    # HEADER TENGAH
    ("ALIGN", (0, 0), (-1, 0), "CENTER"),

    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),

    ("FONTSIZE", (0, 0), (-1, -1), 8),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))


    #----------TABEL KPI 7 ASPEK BERAKHLAK-------------#
    elements.append(table_jam)
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(Paragraph("<b>RANGKUMAN KPI 7 ASPEK BERAKHLAK</b>", styles["Heading3"]))

    aspek_table = [
        ["ASPEK", "NILAI"],
        ["Orientasi Pelayanan", str(kpi_7aspek.get("Orientasi_Pelayanan", "-"))],
        ["Akuntabel", str(kpi_7aspek.get("Akuntabel", "-"))],
        ["Kompeten", str(kpi_7aspek.get("Kompeten", "-"))],
        ["Harmonis", str(kpi_7aspek.get("Harmonis", "-"))],
        ["Loyal", str(kpi_7aspek.get("Loyal", "-"))],
        ["Adaptif", str(kpi_7aspek.get("Adaptif", "-"))],
        ["Kolaboratif", str(kpi_7aspek.get("Kolaboratif", "-"))],
        ["Rata-rata 7 Aspek", str(kpi_7aspek.get("Rata_7_Aspek", "-"))],
        ["Status", str(kpi_7aspek.get("Kategori_7_Aspek", "-"))],
        ["Periode", str(kpi_7aspek.get("Periode", "-"))],
    ]

    table_aspek = Table(aspek_table, colWidths=[8 * cm, 10 * cm])
    table_aspek.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),

    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),

    # HEADER TENGAH
    ("ALIGN", (0, 0), (-1, 0), "CENTER"),

    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),

    ("FONTSIZE", (0, 0), (-1, -1), 8),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))


    #----------SUSUNAN AKHIR RAPORT PEGAWAI-------------#
    elements.append(table_aspek)
    elements.append(Spacer(1, 0.25 * cm))

    tanggal_cetak = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    footer = Paragraph(
        f"Dicetak otomatis melalui Dashboard Kepegawaian pada {tanggal_cetak}.",
        styles["Normal"]
    )

    #----------FOOTER RAPORT-------------#
    elements.append(footer)
    doc.build(elements)

    buffer.seek(0)
    return buffer

# =====================================================================
# =====================================================================
# FUNGSI: Arsip otomatis pegawai pensiun
# ---------------------------------------------------------------------
# Fungsi ini akan:
# 1. Mengecek pegawai yang TMT pensiunnya sudah lewat
# 2. Memindahkan pegawai tersebut ke arsip_pegawai.csv
# 3. Menghapusnya dari data aktif
#
# Fungsi ini dijalankan otomatis saat aplikasi dibuka.
# =====================================================================
# =====================================================================
def arsipkan_pegawai_pensiun_otomatis(df):
    """
    Memindahkan pegawai yang telah memasuki masa pensiun
    ke database Arsip Pegawai secara otomatis.

    Alur:
    1. Cek apakah dataframe kosong.
    2. Konversi kolom TMT PENSIUN menjadi datetime.
    3. Tentukan pegawai yang sudah pensiun.
    4. Tambahkan metadata arsip.
    5. Gabungkan ke database arsip.
    6. Hapus duplikat berdasarkan NIP.
    7. Simpan kembali database arsip.
    8. Kembalikan hanya pegawai aktif.
    """

    # ==========================================================
    # Data kosong
    # ==========================================================
    if df.empty:
        return df

    data = df.copy()

    if "TMT PENSIUN" not in data.columns:
        return data

    # ==========================================================
    # Hari ini (datetime agar tipe sama)
    # ==========================================================
    hari_ini = pd.Timestamp.today().normalize()

    # ==========================================================
    # Konversi tanggal pensiun
    # ==========================================================
    data["TMT_PENSIUN_DATE"] = (
        data["TMT PENSIUN"]
        .apply(parse_tanggal_pegawai)
    )

    # ==========================================================
    # Pegawai yang sudah pensiun
    # ==========================================================
    pegawai_pensiun = data[
        data["TMT_PENSIUN_DATE"].notna()
        &
        (data["TMT_PENSIUN_DATE"] <= hari_ini)
    ].copy()

    # Tidak ada pegawai pensiun
    if pegawai_pensiun.empty:
        return data.drop(columns=["TMT_PENSIUN_DATE"])

    # ==========================================================
    # Metadata Arsip
    # ==========================================================
    pegawai_pensiun["STATUS ARSIP"] = "PENSIUN"
    pegawai_pensiun["TANGGAL ARSIP"] = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    pegawai_pensiun["KETERANGAN"] = (
        "Pegawai otomatis diarsipkan karena pensiun"
    )

    pegawai_pensiun = pegawai_pensiun.drop(columns=["TMT_PENSIUN_DATE"])

    # ==========================================================
    # Membaca database arsip
    # ==========================================================
    if PEGAWAI_ARSIP_PATH.exists():

        try:
            df_arsip = pd.read_csv(
                PEGAWAI_ARSIP_PATH,
                dtype=str
            ).fillna("-")

        except pd.errors.EmptyDataError:
            df_arsip = pd.DataFrame()

    else:
        df_arsip = pd.DataFrame()

    # ==========================================================
    # Gabungkan arsip lama + baru
    # ==========================================================
    df_arsip = pd.concat(
        [df_arsip, pegawai_pensiun],
        ignore_index=True
    )

    # ==========================================================
    # Hapus duplikat
    # ==========================================================
    if "NIP BARU" in df_arsip.columns:

        df_arsip = df_arsip.drop_duplicates(
            subset=["NIP BARU"],
            keep="last"
        )

    # ==========================================================
    # Simpan database arsip
    # ==========================================================
    df_arsip.to_csv(
        PEGAWAI_ARSIP_PATH,
        index=False,
        encoding="utf-8-sig"
    )

    # Refresh cache Streamlit
    load_arsip_pegawai.clear()

    # ==========================================================
    # Sisakan pegawai aktif
    # ==========================================================
    data_aktif = data[
        data["TMT_PENSIUN_DATE"].isna()
        |
        (data["TMT_PENSIUN_DATE"] > hari_ini)
    ].copy()

    data_aktif = data_aktif.drop(columns=["TMT_PENSIUN_DATE"])

    return data_aktif


# =====================================================================
# FUNGSI: Load data arsip pegawai
# ---------------------------------------------------------------------
# Membaca seluruh pegawai yang sudah:
# - pensiun
# - pindah biro
# - resign
# - non aktif
# =====================================================================
@st.cache_data
def load_arsip_pegawai(path: Path = PEGAWAI_ARSIP_PATH):
    """
    Membaca database Arsip Pegawai.
    Jika file belum ada atau masih kosong,
    mengembalikan DataFrame kosong.
    """

    if not path.exists():
        return pd.DataFrame()

    try:
        df = pd.read_csv(path, dtype=str).fillna("-")
    except (pd.errors.EmptyDataError, FileNotFoundError):
        return pd.DataFrame()

    return df

# =====================================================================
# =====================================================================
# FUNGSI BANTU: Header judul halaman
# =====================================================================
# =====================================================================
def page_header(title: str, subtitle: str = ""):
    col1, col2 = st.columns([2, 6])

    with col1:
        st.image(LOGO_PATH, width=160)

    with col2:
        st.markdown(f"""
            <div class="page-title">
                <h1>{title}</h1>
                <p>{subtitle}</p>
            </div>
        """, unsafe_allow_html=True)

def kpi_card(col, title, value, color="blue", font_size=26):
    col.markdown(f"""
        <div class='kpi-card'>
            <h3>{title}</h3>
            <h1 style="font-size:{font_size}px; line-height:1.2; word-break:break-word;">
                {value}
            </h1>
        </div>
    """, unsafe_allow_html=True)



# =====================================================================
# =====================================================================
# FUNGSI: HISTORY ACTIVITY (log kegiatan pengguna)
# ---------------------------------------------------------------------
# Log kegiatan bersifat ringkas (masuk, keluar, download, input pegawai,
# keluarkan pegawai, update pegawai) dan disimpan di MEMORI SERVER
# (bukan session per-browser, dan BUKAN file/CSV/database di penyimpanan
# lokal laptop). Dengan st.cache_resource, satu list yang sama dipakai
# bersama oleh SELURUH pengguna aplikasi ini selama proses Streamlit
# masih berjalan, sehingga histori 1 bulan tetap terlihat di website
# meskipun berpindah menu/browser/user lain yang login.
#
# Batasannya: karena tidak ditulis ke disk sama sekali, histori akan
# ikut kosong kembali jika server/aplikasi Streamlit di-restart penuh
# (mis. redeploy). Ini adalah konsekuensi wajar dari "tanpa penyimpanan
# lokal" - kalau perlu histori yang tahan restart server, itu baru
# membutuhkan penyimpanan permanen (file/database).
# =====================================================================
# =====================================================================
@st.cache_resource
def _log_aktivitas_store() -> list:
    """Wadah in-memory yang dipakai bersama seluruh pengguna & sesi."""
    return []


def catat_aktivitas(kegiatan: str):
    """
    Mencatat satu baris kegiatan pengguna ke penyimpanan in-memory
    server (bukan session_state, bukan file). Otomatis membuang
    baris yang sudah lebih tua dari 30 hari agar memori tidak terus
    membengkak.
    """
    log_store = _log_aktivitas_store()

    log_store.append({
        "WAKTU": datetime.now(),
        "USERNAME": st.session_state.get("username", "-"),
        "NAMA": st.session_state.get("nama", "-"),
        "KEGIATAN": kegiatan,
    })

    batas_awal = datetime.now() - relativedelta(months=1)
    log_store[:] = [b for b in log_store if b["WAKTU"] >= batas_awal]


def load_log_aktivitas() -> pd.DataFrame:
    """
    Memuat log aktivitas dari penyimpanan in-memory server (dipakai
    bersama seluruh pengguna), berisi kegiatan selama kurang lebih
    1 bulan terakhir.
    """
    kolom = ["WAKTU", "USERNAME", "NAMA", "KEGIATAN"]
    data = _log_aktivitas_store()

    if not data:
        return pd.DataFrame(columns=kolom + ["WAKTU_DATE"])

    df = pd.DataFrame(data)

    for kol in kolom:
        if kol not in df.columns:
            df[kol] = "-"

    df["WAKTU_DATE"] = pd.to_datetime(df["WAKTU"], errors="coerce")
    df["WAKTU"] = df["WAKTU_DATE"].dt.strftime("%d-%m-%Y %H:%M:%S")

    return df



# =====================================================================
# =====================================================================
# HALAMAN 1: LOGIN
# =====================================================================
# =====================================================================
def page_login():
    """Tampilan halaman login profesional dengan username & password."""

    # =========================================================
    # CSS KHUSUS HALAMAN LOGIN
    # (diberi scope ke elemen form login saja, tidak memengaruhi
    # tampilan tombol/input di menu-menu lain)
    # =========================================================
    st.markdown("""
        <style>
        div[data-testid="stForm"] {
            background: #ffffff;
            border-radius: 18px;
            padding: 40px 44px 30px 44px;
            box-shadow: 0 16px 40px rgba(30,60,114,0.12);
            border: 1px solid rgba(30,60,114,0.08);
        }

        div[data-testid="stForm"] div[data-testid="stTextInput"] input {
            border-radius: 8px !important;
            border: 1px solid #d7dee6 !important;
            padding: 10px 12px !important;
        }

        div[data-testid="stForm"] div[data-testid="stTextInput"] input:focus {
            border: 1px solid #2a5298 !important;
            box-shadow: 0 0 0 2px rgba(42,82,152,0.15) !important;
        }

        div[data-testid="stFormSubmitButton"] button {
            background: linear-gradient(90deg, #1e3c72, #2a5298);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 12px 0;
            font-weight: 700;
            font-size: 14px;
            letter-spacing: 0.8px;
            margin-top: 6px;
            transition: all 0.2s ease-in-out;
        }

        div[data-testid="stFormSubmitButton"] button:hover {
            background: linear-gradient(90deg, #2a5298, #1e3c72);
            box-shadow: 0 8px 20px rgba(30,60,114,0.30);
            transform: translateY(-1px);
        }

        /* ---- Hilangkan anchor link (#) otomatis di sebelah heading ---- */
        h1 a, h2 a, h3 a, h4 a, h5 a, h6 a,
        [data-testid="stHeadingWithActionElements"] a {
            display: none !important;
        }
        [data-testid="stHeadingWithActionElements"] {
            justify-content: center !important;
        }

        /* ---- Hilangkan tombol zoom in/out (fullscreen) pada logo ---- */
        button[title="View fullscreen"],
        [data-testid="StyledFullScreenButton"] {
            display: none !important;
            visibility: hidden !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # Layout center menggunakan kolom kosong di kiri-kanan
    col1, col2, col3 = st.columns([1, 1.3, 1])

    with col2:
        # ---- Logo & identitas instansi ----
        logo1, logo2, logo3 = st.columns([1, 1, 1])
        with logo2:
            st.image(LOGO_PATH, width=150)

        st.markdown("""
            <div style="width:100%; text-align:center; margin:0 auto; padding:14px 0 32px 0;">
                <h1 style="
                    color:#1e3c72;
                    font-size:38px;
                    font-weight:800;
                    letter-spacing:0.6px;
                    text-transform:uppercase;
                    margin:0 auto 8px auto;
                    text-align:center;
                    width:100%;
                ">
                    Dashboard Kepegawaian
                </h1>
                <p style="
                    color:#3d4b5c;
                    font-size:15px;
                    font-weight:600;
                    letter-spacing:0.4px;
                    margin:0 auto;
                    text-align:center;
                    width:100%;
                ">
                    BIRO UMUM SEKRETARIAT DAERAH PROVINSI JAWA TIMUR
                </p>
            </div>
        """, unsafe_allow_html=True)

        # Form input login
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Masukkan username")
            password = st.text_input("Password", type="password", placeholder="Masukkan password")
            submitted = st.form_submit_button("MASUK", use_container_width=True)

            # Validasi login
            if submitted:
                if username in USERS and USERS[username]["password"] == password:
                    # Simpan status login ke session_state
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = username
                    st.session_state["nama"] = USERS[username]["nama"]
                    st.session_state["role"] = USERS[username]["role"]
                    st.session_state["login_time"] = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    catat_aktivitas("Masuk ke sistem")
                    st.success("Login berhasil. Memuat dashboard...")
                    st.rerun()
                else:
                    st.error("Username atau password salah.")

        # ---- Footer ----
        st.markdown("""
            <div style="text-align:center; padding-top:24px;">
                <p style="color:#9aa4b0; font-size:11px; margin:0;">
                    © 2026 Sekretariat Daerah Provinsi Jawa Timur — Dashboard Kepegawaian
                </p>
            </div>
        """, unsafe_allow_html=True)


# =====================================================================
# =====================================================================
# MENU 1: INFORMASI PEGAWAI (dengan pencarian)
# =====================================================================
# =====================================================================
def filter_pegawai_belum_pensiun(df):
    """
    Menghapus pegawai dari tampilan Menu 1 apabila durasi menuju pensiun
    sudah habis / 0 Tahun 0 Bulan 0 Hari.
    """

    # Jika data kosong, langsung kembalikan dataframe kosong
    if df.empty:
        return df

    # Copy data agar dataframe asli tidak berubah
    data = df.copy()

    # Jika kolom TMT PENSIUN tidak ada, data tidak difilter
    if "TMT PENSIUN" not in data.columns:
        return data

    # Ambil tanggal hari ini
    hari_ini = pd.Timestamp.today().normalize()

    # Ubah kolom TMT PENSIUN menjadi format tanggal agar bisa dibandingkan
    data["TMT_PENSIUN_DATE"] = data["TMT PENSIUN"].apply(parse_tanggal_pegawai)

    # Pegawai aktif adalah pegawai yang:
    # 1. TMT PENSIUN tidak kosong/valid, dan tanggal pensiun masih lebih besar dari hari ini
    # 2. Jika TMT PENSIUN kosong/tidak valid, tetap ditampilkan agar tidak salah hapus
    data = data[
        data["TMT_PENSIUN_DATE"].isna()
        | (data["TMT_PENSIUN_DATE"] > hari_ini)
    ]

    # Hapus kolom bantu agar tidak ikut tampil di tabel
    data = data.drop(columns=["TMT_PENSIUN_DATE"])

    return data


def menu_informasi_pegawai(df: pd.DataFrame):
    page_header("Informasi Pegawai",
                "Daftar lengkap pegawai dengan fitur pencarian & filter")

    if df.empty:
        st.warning(
            f"Data Master Pegawai tidak ditemukan.\n"
            f"Pastikan file berada di:\n{DATA_MASTER_PATH}"
        )
        return
    
    # ---------- Bagian Pencarian & Filter ----------
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        # Input pencarian berdasarkan nama / NIP
        keyword = st.text_input("🔍 Cari Nama / NIP / NIK", placeholder="Ketik kata kunci...")
    with col2:
        # Filter Unit Kerja
        unit_opt = ["Semua"] + sorted(df["UNIT KERJA"].dropna().unique().tolist()) if "UNIT KERJA" in df.columns else ["Semua"]
        f_unit = st.selectbox("Unit Kerja", unit_opt)
    with col3:
        # Filter Status Pegawai
        sp_opt = ["Semua"] + sorted(df["STATUS PNS/CPNS"].dropna().unique().tolist()) if "STATUS PNS/CPNS" in df.columns else ["Semua"]
        f_status = st.selectbox("Status Pegawai", sp_opt)

    # ---------- Terapkan filter ----------
    # Tambahkan kolom lama masa kerja dan durasi menuju pensiun secara real-time
    data = hitung_menu1_durasi(df)

    # Otomatis keluarkan pegawai dari Menu 1 jika sudah pensiun
    # Pegawai dianggap pensiun jika TMT PENSIUN sudah <= tanggal hari ini
    data = filter_pegawai_belum_pensiun(data)

    # Simpan jumlah pegawai aktif untuk informasi tampilan
    total_pegawai_aktif = len(data)

    if keyword:
        kw = keyword.lower()

        # Cari di kolom NAMA, NIP BARU, NIK
        mask = (
            data["NAMA"].str.lower().str.contains(kw, na=False)
            | data.get("NIP BARU", pd.Series([""]*len(data))).str.lower().str.contains(kw, na=False)
            | data.get("NIK", pd.Series([""]*len(data))).str.lower().str.contains(kw, na=False)
        )

        data = data[mask]

    if f_unit != "Semua" and "UNIT KERJA" in data.columns:
        data = data[data["UNIT KERJA"] == f_unit]

    if f_status != "Semua" and "STATUS PNS/CPNS" in data.columns:
        data = data[data["STATUS PNS/CPNS"] == f_status]

    st.info(
        f"Menampilkan **{len(data)}** dari **{total_pegawai_aktif}** pegawai aktif. "
        "Pegawai dengan durasi pensiun 0  otomatis tidak ditampilkan."
    )

    
    data = hitung_masa_pangkat_golongan(data)

    # Nomor urut (kolom NO) dibuat ulang 1..N mengikuti hasil pencarian/filter
    # yang sedang tampil di layar, supaya selalu berurutan rapi -- bukan
    # nomor asli dari data mentah yang bisa bolong/loncat setelah difilter.
    data = data.reset_index(drop=True)
    data["NO"] = range(1, len(data) + 1)

    kolom_tampil = [c for c in
        ["NO", "NAMA", "NIP BARU", "NIK", "TGL LAHIR", "USIA", "JK", "AGAMA", "STATUS PNS/CPNS",
        "JENJANG", "GOLONGAN", "NAMA JABATAN", "UNIT KERJA",
        "TMT MASUK KERJA", "LAMA MASA KERJA",
        "TMT PANGKAT", "MASA_PANGKAT",
        "TMT PENSIUN", "DURASI MENUJU PENSIUN"]
    if c in data.columns]
    st.dataframe(data[kolom_tampil], use_container_width=True, height=450, hide_index=True)

    # =========================================================
    # CETAK / DOWNLOAD DATA (mengikuti hasil pencarian & filter di atas)
    # =========================================================
    tabel_unduh = data[kolom_tampil].copy().reset_index(drop=True)
    csv_informasi_pegawai = tabel_unduh.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        "🖨️ Cetak / Download CSV",
        csv_informasi_pegawai,
        file_name=f"informasi_pegawai_{date.today().strftime('%Y-%m-%d')}.csv",
        mime="text/csv",
        use_container_width=True,
        on_click=lambda: catat_aktivitas("Download Informasi Pegawai")
    )


# =====================================================================
# MENU 2: DASHBOARD INFORMASI PEGAWAI PROFESIONAL
# =====================================================================
def kpi_card_dashboard(col, title, value):
    col.markdown(f"""
        <div style="
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            border-radius: 14px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            color: white;
            height: 148px;
            padding: 14px 10px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            text-align: center;
        ">
            <div style="
                font-size: 12px;
                font-weight: 600;
                letter-spacing: 0.2px;
                line-height: 1.35;
                opacity: 0.9;
                margin-bottom: 10px;
            ">
                {title}
            </div>
            <div style="
                font-size: 32px;
                font-weight: 800;
                line-height: 1;
            ">
                {value}
            </div>
        </div>
    """, unsafe_allow_html=True)

def menu_dashboard(df: pd.DataFrame):
    """
    Dashboard Informasi Pegawai untuk pimpinan.

    Alur cerita dashboard ini (dari atas ke bawah) dirancang supaya mudah
    dipahami oleh kalangan non-teknis:
      1. Ringkasan Utama         -> angka-angka kunci sekilas pandang
      2. Profil Usia, Pendidikan, dan Agama Pegawai -> identitas dasar pegawai
      3. Jenjang Karier Pegawai  -> sebaran golongan dan jabatan
      4. Sebaran Unit Kerja dan Status Kepegawaian -> penempatan pegawai
      5. Masa Kerja dan Masa Jabatan Pegawai -> lama mengabdi & lama di golongan saat ini
      6. Proyeksi dan Perencanaan Pensiun -> ditutup dengan daftar pegawai
         yang akan pensiun < 1 tahun.
    """

    page_header(
        "Dashboard Informasi Pegawai",
        "Executive summary data kepegawaian Biro Umum"
    )

    if df.empty:
        st.warning("Data pegawai belum tersedia.")
        return

    # =========================================================
    # PERSIAPAN DATA
    # =========================================================

    # Copy data agar dataframe asli tidak berubah
    data = df.copy()

    # Hitung seluruh informasi real-time
    data = hitung_menu1_durasi(data)
    data = hitung_masa_pangkat_golongan(data)

    # =========================================================
    # KONVERSI KOLOM TANGGAL
    # =========================================================
    data["TGL_LAHIR_DATE"] = (
        data["TGL LAHIR"].apply(parse_tanggal_pegawai)
        if "TGL LAHIR" in data.columns else pd.NaT
    )

    data["TMT_MASUK_DATE"] = (
        data["TMT MASUK KERJA"].apply(parse_tanggal_pegawai)
        if "TMT MASUK KERJA" in data.columns else pd.NaT
    )

    data["TMT_PENSIUN_DATE"] = (
        data["TMT PENSIUN"].apply(parse_tanggal_pegawai)
        if "TMT PENSIUN" in data.columns else pd.NaT
    )

    # (Opsional)
    data["TMT_PANGKAT_DATE"] = (
        data["TMT PANGKAT"].apply(parse_tanggal_pegawai)
        if "TMT PANGKAT" in data.columns else pd.NaT
    )

    # =========================================================
    # HITUNG USIA
    # =========================================================
    data["USIA_NUM"] = (
        data["USIA"].apply(parse_usia)
        if "USIA" in data.columns
        else np.nan
    )

    # =========================================================
    # HITUNG MASA KERJA
    # =========================================================
    hari_ini = pd.Timestamp.today().normalize()

    data["MASA_KERJA_TAHUN"] = data["TMT_MASUK_DATE"].apply(
        lambda x: relativedelta(hari_ini, x).years
        if pd.notna(x) and x <= hari_ini
        else np.nan
    )

    # =========================================================
    # HITUNG USIA SAAT MASUK KERJA
    # =========================================================
    def _usia_saat_masuk(row):
        lahir = row["TGL_LAHIR_DATE"]
        masuk = row["TMT_MASUK_DATE"]
        if pd.notna(lahir) and pd.notna(masuk) and masuk >= lahir:
            return relativedelta(masuk, lahir).years
        return np.nan

    data["USIA_SAAT_MASUK"] = data.apply(_usia_saat_masuk, axis=1)

    # =========================================================
    # HITUNG SISA PENSIUN
    # =========================================================
    data["SISA_PENSIUN_HARI"] = data["TMT_PENSIUN_DATE"].apply(
        lambda x: (x - hari_ini).days
        if pd.notna(x)
        else np.nan
    )

    # =========================================================
    # FILTER PEGAWAI AKTIF
    # =========================================================
    data = data[
        data["TMT_PENSIUN_DATE"].isna()
        | (data["TMT_PENSIUN_DATE"] > hari_ini)
    ]

    # =========================================================
    # FILTER / SLICER DASHBOARD
    # =========================================================
    st.markdown("### Filter Data Pegawai")

    f1, f2, f3, f4, f5 = st.columns(5)

    with f1:
        status_opsi = ["Semua"]
        if "STATUS PNS/CPNS" in data.columns:
            status_opsi += sorted(data["STATUS PNS/CPNS"].dropna().astype(str).unique().tolist())

        filter_status = st.selectbox(
            "Status Kepegawaian",
            status_opsi,
            key="dash_filter_status"
        )

    with f2:
        unit_opsi = ["Semua"]
        if "UNIT KERJA" in data.columns:
            unit_opsi += sorted(data["UNIT KERJA"].dropna().astype(str).unique().tolist())

        filter_unit = st.selectbox(
            "Unit Kerja",
            unit_opsi,
            key="dash_filter_unit"
        )

    with f3:
        jenjang_opsi = ["Semua"]
        if "JENJANG" in data.columns:
            jenjang_opsi += sorted(data["JENJANG"].dropna().astype(str).unique().tolist())

        filter_jenjang = st.selectbox(
            "Pendidikan Terakhir",
            jenjang_opsi,
            key="dash_filter_jenjang"
        )

    with f4:
        jk_opsi = ["Semua"]
        if "JK" in data.columns:
            jk_opsi += sorted(data["JK"].dropna().astype(str).unique().tolist())

        filter_jk = st.selectbox(
            "Jenis Kelamin",
            jk_opsi,
            key="dash_filter_jk"
        )

    with f5:
        golongan_opsi = ["Semua"]
        if "GOLONGAN" in data.columns:
            golongan_opsi += sorted(
                data["GOLONGAN"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

        filter_golongan = st.selectbox(
            "Golongan",
            golongan_opsi,
            key="dash_filter_golongan"
        )

    # Terapkan filter status
    if filter_status != "Semua" and "STATUS PNS/CPNS" in data.columns:
        data = data[data["STATUS PNS/CPNS"] == filter_status]

    # Terapkan filter unit kerja
    if filter_unit != "Semua" and "UNIT KERJA" in data.columns:
        data = data[data["UNIT KERJA"] == filter_unit]

    # Terapkan filter pendidikan
    if filter_jenjang != "Semua" and "JENJANG" in data.columns:
        data = data[data["JENJANG"] == filter_jenjang]

    # Terapkan filter jenis kelamin (filter tetap ada, walau KPI/chart gender dihapus)
    if filter_jk != "Semua" and "JK" in data.columns:
        data = data[data["JK"] == filter_jk]

    # Terapkan filter golongan
    if filter_golongan != "Semua" and "GOLONGAN" in data.columns:
        data = data[data["GOLONGAN"] == filter_golongan]

    if data.empty:
        st.warning("Tidak ada data pegawai yang sesuai dengan filter.")
        return

    # =========================================================
    # 1. RINGKASAN UTAMA
    # =========================================================
    st.markdown("---")
    st.markdown("### Ringkasan Utama")
    st.caption(
        "Angka-angka inti kondisi kepegawaian saat ini, sesuai filter yang dipilih di atas."
    )

    total_pegawai = len(data)

    rata_usia = (
        round(data["USIA_NUM"].mean(), 1)
        if not data["USIA_NUM"].dropna().empty
        else "-"
    )

    rata_usia_masuk = (
        round(data["USIA_SAAT_MASUK"].mean(), 1)
        if not data["USIA_SAAT_MASUK"].dropna().empty
        else "-"
    )

    rata_masa_kerja = (
        round(data["MASA_KERJA_TAHUN"].mean(), 1)
        if not data["MASA_KERJA_TAHUN"].dropna().empty
        else "-"
    )

    pensiun_1_tahun = (
        data["SISA_PENSIUN_HARI"].between(0, 365, inclusive="both").sum()
        if "SISA_PENSIUN_HARI" in data.columns
        else 0
    )

    eselon_count = (
        data["NAMA JABATAN"].astype(str).str.contains("KEPALA", case=False, na=False).sum()
        if "NAMA JABATAN" in data.columns
        else 0
    )

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    kpi_card_dashboard(c1, "Total Pegawai", total_pegawai)
    kpi_card_dashboard(c2, "Jumlah Eselon", eselon_count)
    kpi_card_dashboard(c3, "Rata-rata Usia", rata_usia)
    kpi_card_dashboard(c4, "Rata-rata Usia Saat Masuk Kerja", rata_usia_masuk)
    kpi_card_dashboard(c5, "Rata-rata Masa Kerja (Tahun)", rata_masa_kerja)
    kpi_card_dashboard(c6, "Pegawai Akan Pensiun (1 Tahun)", pensiun_1_tahun)

    # =========================================================
    # 2. PROFIL USIA, PENDIDIKAN, DAN AGAMA PEGAWAI (identitas dasar)
    # =========================================================
    st.markdown("---")
    st.markdown("### Profil Identitas Pegawai")
    st.caption(
        "Gambaran usia, latar belakang pendidikan, dan komposisi agama seluruh pegawai aktif."
    )

    usia_bins = [20, 29, 36, 44, 51, 58, 70]
    usia_labels = [
        "21-28 tahun",
        "29-36 tahun",
        "37-44 tahun",
        "45-51 tahun",
        "52-58 tahun",
        ">58 tahun",
    ]

    data["KELOMPOK_USIA"] = pd.cut(
        data["USIA_NUM"],
        bins=usia_bins,
        labels=usia_labels,
        include_lowest=True
    )

    g1, g2, g3 = st.columns(3)

    with g1:
        usia_df = (
            data["KELOMPOK_USIA"]
            .value_counts()
            .sort_index()
            .reset_index()
        )
        usia_df.columns = ["Rentang Usia", "Jumlah"]

        fig = px.bar(
            usia_df,
            x="Rentang Usia",
            y="Jumlah",
            text="Jumlah",
            title="Distribusi Kelompok Usia",
            color="Jumlah",
            color_continuous_scale="Blues"
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(height=420)

        st.plotly_chart(fig, use_container_width=True)

    with g2:
        if "JENJANG" in data.columns:
            pendidikan_df = data["JENJANG"].value_counts().reset_index()
            pendidikan_df.columns = ["Jenjang Pendidikan", "Jumlah"]

            fig = px.bar(
                pendidikan_df.sort_values("Jumlah"),
                x="Jumlah",
                y="Jenjang Pendidikan",
                orientation="h",
                text="Jumlah",
                title="Distribusi Pendidikan Terakhir",
                color="Jumlah",
                color_continuous_scale="Blues"
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(height=420)

            st.plotly_chart(fig, use_container_width=True)

    with g3:
        if "AGAMA" in data.columns:
            agama_df = data["AGAMA"].value_counts().reset_index()
            agama_df.columns = ["Agama", "Jumlah"]

            fig = px.pie(
                agama_df,
                names="Agama",
                values="Jumlah",
                title="Distribusi Agama Pegawai",
                hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig.update_layout(height=420)

            st.plotly_chart(fig, use_container_width=True)

    # =========================================================
    # 3. JENJANG KARIER PEGAWAI
    # =========================================================
    st.markdown("---")
    st.markdown("### Jenjang Karier Pegawai")
    st.caption(
        "Sebaran golongan/pangkat serta jabatan dengan jumlah pegawai terbanyak."
    )

    g4, g5 = st.columns(2)

    with g4:
        if "GOLONGAN" in data.columns:
            # Bersihkan nilai kosong/'-'/'nan' menjadi label eksplisit "Tanpa Golongan",
            # supaya batang kosong pada grafik tidak membingungkan pembaca -- ini
            # mewakili pegawai yang memang tidak memiliki golongan/pangkat (misalnya
            # PPPK Paruh Waktu atau status kepegawaian non-golongan lainnya).
            gol_bersih = data["GOLONGAN"].astype(str).str.strip()
            gol_bersih = gol_bersih.replace(
                {"": "Tanpa Golongan", "nan": "Tanpa Golongan",
                 "None": "Tanpa Golongan", "-": "Tanpa Golongan"}
            )

            golongan_df = gol_bersih.value_counts().reset_index()
            golongan_df.columns = ["Golongan", "Jumlah"]

            fig = px.bar(
                golongan_df.sort_values("Jumlah", ascending=False),
                x="Golongan",
                y="Jumlah",
                text="Jumlah",
                title="Distribusi Golongan / Pangkat",
                color="Jumlah",
                color_continuous_scale="Purples"
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(height=430, xaxis_title="Golongan")

            st.plotly_chart(fig, use_container_width=True)

    with g5:
        if "NAMA JABATAN" in data.columns:
            jabatan_df = data["NAMA JABATAN"].value_counts().head(10).reset_index()
            jabatan_df.columns = ["Jabatan", "Jumlah"]

            fig = px.bar(
                jabatan_df.sort_values("Jumlah"),
                x="Jumlah",
                y="Jabatan",
                orientation="h",
                text="Jumlah",
                title="10 Jabatan dengan Pegawai Terbanyak",
                color="Jumlah",
                color_continuous_scale="Teal"
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(height=430)

            st.plotly_chart(fig, use_container_width=True)

    # =========================================================
    # 4. SEBARAN UNIT KERJA DAN STATUS KEPEGAWAIAN
    # =========================================================
    st.markdown("---")
    st.markdown("### Sebaran Unit Kerja dan Status Kepegawaian")
    st.caption(
        "Unit kerja dengan jumlah pegawai terbanyak dan komposisi status kepegawaian."
    )

    g6, g7 = st.columns(2)

    with g6:
        if "UNIT KERJA" in data.columns:
            unit_df = data["UNIT KERJA"].value_counts().head(10).reset_index()
            unit_df.columns = ["Unit Kerja", "Jumlah"]

            fig = px.bar(
                unit_df.sort_values("Jumlah"),
                x="Jumlah",
                y="Unit Kerja",
                orientation="h",
                text="Jumlah",
                title="10 Unit Kerja dengan Pegawai Terbanyak",
                color="Jumlah",
                color_continuous_scale="Blues"
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(height=430)

            st.plotly_chart(fig, use_container_width=True)

    with g7:
        if "STATUS PNS/CPNS" in data.columns:
            status_df = data["STATUS PNS/CPNS"].value_counts().reset_index()
            status_df.columns = ["Status Kepegawaian", "Jumlah"]

            fig = px.pie(
                status_df,
                names="Status Kepegawaian",
                values="Jumlah",
                title="Komposisi Status Kepegawaian",
                hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            fig.update_layout(height=430)

            st.plotly_chart(fig, use_container_width=True)

    # =========================================================
    # 5. MASA KERJA DAN MASA JABATAN PEGAWAI
    # =========================================================
    st.markdown("---")
    st.markdown("### Masa Kerja dan Masa Jabatan Pegawai")
    st.caption(
        "Lama masa kerja pegawai secara keseluruhan, serta lama pegawai menduduki "
        "golongan/pangkat saat ini -- untuk melihat berapa pegawai yang sudah lama "
        "berada pada golongan/jabatan yang sama."
    )
 
    masa_bins = [0, 10, 20, 30, 40, 50, 60]
    masa_labels = [
        "0-10 tahun",
        "11-20 tahun",
        "21-30 tahun",
        "31-40 tahun",
        "41-50 tahun",
        "51-60 tahun",
    ]
 
    data["KELOMPOK_MASA_KERJA"] = pd.cut(
        data["MASA_KERJA_TAHUN"],
        bins=masa_bins,
        labels=masa_labels,
        include_lowest=True
    )
 
    STEP_MASA_PANGKAT = 2
 
    max_masa_pangkat = data["MASA_PANGKAT_TAHUN"].max()
 
    if pd.notna(max_masa_pangkat):
        batas_atas = int(
            np.ceil((max_masa_pangkat + 1) / STEP_MASA_PANGKAT) * STEP_MASA_PANGKAT
        )
        batas_atas = max(batas_atas, STEP_MASA_PANGKAT)
 
        masa_pangkat_bins = list(range(0, batas_atas + STEP_MASA_PANGKAT, STEP_MASA_PANGKAT))
        masa_pangkat_labels = [
            f"{masa_pangkat_bins[i]}-{masa_pangkat_bins[i + 1] - 1} tahun"
            for i in range(len(masa_pangkat_bins) - 1)
        ]
    else:
        masa_pangkat_bins = [0, STEP_MASA_PANGKAT]
        masa_pangkat_labels = [f"0-{STEP_MASA_PANGKAT - 1} tahun"]
 
    data["KELOMPOK_MASA_PANGKAT"] = pd.cut(
        data["MASA_PANGKAT_TAHUN"],
        bins=masa_pangkat_bins,
        labels=masa_pangkat_labels,
        include_lowest=True
    )
 
    g8, g9 = st.columns(2)
 
    with g8:
        masa_df = (
            data["KELOMPOK_MASA_KERJA"]
            .value_counts()
            .sort_index()
            .reset_index()
        )
        masa_df.columns = ["Rentang Masa Kerja", "Jumlah"]
 
        fig = px.bar(
            masa_df,
            x="Rentang Masa Kerja",
            y="Jumlah",
            text="Jumlah",
            title="Distribusi Masa Kerja Pegawai",
            color="Jumlah",
            color_continuous_scale="Purples"
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(height=420)
 
        st.plotly_chart(fig, use_container_width=True)
 
    with g9:
        masa_pangkat_df = (
            data["KELOMPOK_MASA_PANGKAT"]
            .value_counts()
            .sort_index()
            .reset_index()
        )
        masa_pangkat_df.columns = ["Rentang Masa Golongan/Pangkat", "Jumlah"]
 
        fig = px.bar(
            masa_pangkat_df,
            x="Rentang Masa Golongan/Pangkat",
            y="Jumlah",
            text="Jumlah",
            title="Distribusi Masa Golongan/Pangkat Saat Ini",
            color="Jumlah",
            color_continuous_scale="Oranges"
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(height=420)
 
        st.plotly_chart(fig, use_container_width=True)


    # =========================================================
    # 6. PROYEKSI DAN PERENCANAAN PENSIUN
    # =========================================================
    st.markdown("---")
    st.markdown("### Proyeksi dan Perencanaan Pensiun")
    st.caption(
        "Proyeksi jumlah pegawai pensiun per tahun, ditutup dengan daftar nama pegawai "
        "yang akan pensiun dalam 1 tahun ke depan untuk kebutuhan perencanaan suksesi."
    )

    data["TAHUN_PENSIUN_REAL"] = data["TMT_PENSIUN_DATE"].dt.year

    pensiun_df = (
        data.dropna(subset=["TAHUN_PENSIUN_REAL"])
        .groupby("TAHUN_PENSIUN_REAL")
        .size()
        .reset_index(name="Jumlah")
        .sort_values("TAHUN_PENSIUN_REAL")
    )

    if not pensiun_df.empty:
        fig = px.bar(
            pensiun_df,
            x="TAHUN_PENSIUN_REAL",
            y="Jumlah",
            text="Jumlah",
            title="Proyeksi Jumlah Pegawai Pensiun per Tahun",
            color="Jumlah",
            color_continuous_scale="Reds"
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(
            xaxis_title="Tahun Pensiun",
            yaxis_title="Jumlah Pegawai",
            height=400
        )
        st.plotly_chart(fig, use_container_width=True)

    # ---- Daftar pegawai yang akan pensiun < 1 tahun (paling bawah, full width) ----
    st.markdown("#### Daftar Pegawai yang Akan Pensiun dalam 1 Tahun")

    cols_pensiun = [
        "NAMA",
        "UNIT KERJA",
        "NAMA JABATAN",
        "GOLONGAN",
        "TMT PENSIUN",
        "DURASI MENUJU PENSIUN",
    ]
    cols_pensiun = [c for c in cols_pensiun if c in data.columns]

    pensiun_dekat = data[
        data["SISA_PENSIUN_HARI"].between(0, 365, inclusive="both")
    ].sort_values("SISA_PENSIUN_HARI").reset_index(drop=True)

    if pensiun_dekat.empty:
        st.success("Tidak ada pegawai yang pensiun dalam 1 tahun.")
    else:
        tabel_pensiun = pensiun_dekat[cols_pensiun].copy()
        tabel_pensiun.insert(0, "NO", range(1, len(tabel_pensiun) + 1))

        st.dataframe(
            tabel_pensiun,
            use_container_width=True,
            height=380,
            hide_index=True
        )


# =====================================================================
# =====================================================================
# MENU 3: KELOLA PEGAWAI
# ---------------------------------------------------------------------
# Berisi 3 fitur utama:
# 1. Update Pegawai
# 2. Input Pegawai
# 3. Hapus Karyawan
#
# Semua perubahan disimpan ke pegawai_kelola.csv sehingga Menu 1 otomatis
# membaca data terbaru setelah st.rerun().
# =====================================================================
# =====================================================================
def bersihkan_teks(value):
    """Membersihkan input teks agar konsisten."""
    if value is None:
        return "-"
    value = str(value).strip()
    return value if value else "-"


def format_tanggal_csv(value):
    """Mengubah tanggal Python menjadi format teks DD-MM-YYYY untuk CSV."""
    if value is None:
        return "-"
    if isinstance(value, date):
        return value.strftime("%d-%m-%Y")

    parsed = parse_tanggal_pegawai(value)
    if pd.isna(parsed):
        return "-"

    return parsed.strftime("%d-%m-%Y")


def default_date_input(value, default_value=None):
    """
    Mengubah nilai dari CSV menjadi date untuk st.date_input.
    Jika gagal, gunakan default_value.
    """
    parsed = parse_tanggal_pegawai(value)

    if not pd.isna(parsed):
        return parsed.date()

    if default_value is not None:
        return default_value

    return date.today()


def hitung_usia_dari_tanggal_lahir(tanggal_lahir):
    """Menghitung usia dari tanggal lahir."""
    if tanggal_lahir is None:
        return "-"

    hari_ini = pd.Timestamp.today().normalize()
    usia = (
        hari_ini.year
        - tanggal_lahir.year
        - ((hari_ini.month, hari_ini.day) < (tanggal_lahir.month, tanggal_lahir.day))
    )

    return f"{usia} Tahun"


def opsi_unik(df, kolom, default_list=None):
    """
    Mengambil opsi unik dari kolom dataframe.
    Jika kolom kosong/tidak tersedia, gunakan default_list.
    """
    if default_list is None:
        default_list = []

    if df.empty or kolom not in df.columns:
        return default_list

    data = (
        df[kolom]
        .dropna()
        .astype(str)
        .str.strip()
    )

    data = data[~data.isin(["", "-", "nan", "None"])]

    hasil = sorted(data.unique().tolist())

    return hasil if hasil else default_list


def normalisasi_data_pegawai(df):
    """
    Menyamakan struktur kolom pegawai agar aman untuk input, update, dan hapus.
    """
    if df is None or df.empty:
        df = pd.DataFrame(columns=KOLOM_MINIMAL_PEGAWAI)
    else:
        df = df.copy()

    df.columns = [str(c).strip() for c in df.columns]

    for col in KOLOM_MINIMAL_PEGAWAI:
        if col not in df.columns:
            df[col] = "-"

    df = df.fillna("-")

    # ------------------------------------------------------------
    # Hitung ulang USIA dan TAHUN PENSIUN untuk SELURUH baris.
    # ------------------------------------------------------------
    # BUG SEBELUMNYA: kedua kolom ini hanya pernah dihitung satu-per-satu
    # di form Input Pegawai dan Update Pegawai. Untuk data yang masuk lewat
    # jalur lain (load awal dari Data_Master.xlsx),
    # tidak ada kode yang menghitungnya sama sekali - akibatnya kolom
    # USIA dan TAHUN PENSIUN selalu "-" untuk seluruh/hampir seluruh
    # pegawai. Dihitung ulang di sini (fungsi yang dipanggil di semua
    # jalur simpan/baca data) supaya selalu konsisten dan tidak kosong.
    if "TGL LAHIR" in df.columns:
        df["USIA"] = df["TGL LAHIR"].apply(hitung_usia_real_time)

    if "TMT PENSIUN" in df.columns:
        def _hitung_tahun_pensiun(nilai):
            tgl = parse_tanggal_pegawai(nilai)
            return str(tgl.year) if pd.notna(tgl) else "-"

        df["TAHUN PENSIUN"] = df["TMT PENSIUN"].apply(_hitung_tahun_pensiun)

    # Bersihkan kolom penting
    if "NAMA" in df.columns:
        df["NAMA"] = df["NAMA"].astype(str).str.strip()

    if "NIP BARU" in df.columns:
        df["NIP BARU"] = df["NIP BARU"].astype(str).str.strip()

    if "NO" in df.columns:
        df["NO"] = range(1, len(df) + 1)

    return df


def load_data_pegawai_kelola(seed_df, path=PEGAWAI_MASTER_PATH):
    if os.path.exists(path):
        try:
            df = pd.read_csv(path, dtype=str).fillna("-")
            df = normalisasi_nama_kolom_pegawai(df)
            df = normalisasi_data_pegawai(df)

            # Bersihkan NIK: menandai NIK yang sudah terlanjur tersimpan
            # sebagai teks tanggal (bug lama), bukan menampilkannya diam-diam.
            if "NIK" in df.columns:
                df["NIK"] = df["NIK"].apply(bersihkan_nik)

            # Hilangkan jam pada seluruh kolom tanggal
            kolom_tanggal = [
                "TGL LAHIR",
                "TMT MASUK KERJA",
                "TMT PANGKAT",
                "TMT PENSIUN",
            ]

            for kol in kolom_tanggal:
                if kol in df.columns:
                    df[kol] = (
                        parse_kolom_tanggal(df[kol])
                        .dt.strftime("%d-%m-%Y")
                        .fillna("-")
                    )

            return df

        except pd.errors.EmptyDataError:
            pass

    seed_df = normalisasi_data_pegawai(seed_df)

    if "NIK" in seed_df.columns:
        seed_df["NIK"] = seed_df["NIK"].apply(bersihkan_nik)

    kolom_tanggal = [
        "TGL LAHIR",
        "TMT MASUK KERJA",
        "TMT PANGKAT",
        "TMT PENSIUN",
    ]

    for kol in kolom_tanggal:
        if kol in seed_df.columns:
            seed_df[kol] = (
                parse_kolom_tanggal(seed_df[kol])
                .dt.strftime("%d-%m-%Y")
                .fillna("-")
            )

    return seed_df

def simpan_data_pegawai_kelola(df, path=PEGAWAI_MASTER_PATH):
    """
    Menyimpan dataframe pegawai ke database operasional.
    """
    df = normalisasi_data_pegawai(df)

    # Bersihkan NIK sebelum disimpan, supaya nilai yang sudah rusak
    # (mis. tertimpa objek tanggal) tidak ikut tersimpan permanen ke CSV.
    if "NIK" in df.columns:
        df["NIK"] = df["NIK"].apply(bersihkan_nik)

    # Urutkan ulang nomor agar selalu rapi
    if "NO" in df.columns:
        df["NO"] = range(1, len(df) + 1)

    # quoting=QUOTE_ALL: mencegah Excel/aplikasi spreadsheet lain
    # mendeteksi otomatis kolom NIK/NIP sebagai angka besar (notasi
    # ilmiah) atau kolom tanggal sebagai Date saat file CSV ini dibuka
    # ulang secara manual, dan mencegah pergeseran kolom akibat koma
    # yang tidak ter-escape pada kolom teks bebas (mis. alamat).
    df.to_csv(
        path,
        index=False,
        encoding="utf-8-sig",
        quoting=csv.QUOTE_ALL,
    )

    # Bersihkan cache agar Menu 1, Dashboard, dan KPI membaca data terbaru
    st.cache_data.clear()

    return df


def arsipkan_pegawai(row_pegawai, status, alasan, tanggal_arsip, catatan="-"):
    """
    Memindahkan pegawai ke database Arsip Pegawai.
    Digunakan untuk pegawai yang dihapus, pensiun, atau status arsip lainnya.
    """

    row_arsip = row_pegawai.copy()

    row_arsip["STATUS ARSIP"] = status
    row_arsip["TANGGAL ARSIP"] = format_tanggal_csv(tanggal_arsip)
    row_arsip["KETERANGAN"] = alasan
    row_arsip["CATATAN"] = bersihkan_teks(catatan)
    row_arsip["DIPROSES OLEH"] = st.session_state.get("nama", "-")
    row_arsip["WAKTU PROSES"] = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    if PEGAWAI_ARSIP_PATH.exists():
        df_arsip = pd.read_csv(PEGAWAI_ARSIP_PATH, dtype=str).fillna("-")
    else:
        df_arsip = pd.DataFrame()

    df_arsip = pd.concat(
        [df_arsip, pd.DataFrame([row_arsip])],
        ignore_index=True
    )

    # Hindari duplikasi berdasarkan NIP
    if "NIP BARU" in df_arsip.columns:
        df_arsip = df_arsip.drop_duplicates(
            subset=["NIP BARU"],
            keep="last"
        )

    df_arsip.to_csv(
        PEGAWAI_ARSIP_PATH,
        index=False,
        encoding="utf-8-sig"
    )

def validasi_nip_unik(df, nip_baru, index_diabaikan=None):
    """
    Validasi agar NIP tidak ganda.
    index_diabaikan digunakan saat update pegawai agar NIP milik dirinya sendiri
    tidak dianggap duplikat.
    """
    if "NIP BARU" not in df.columns:
        return True

    nip_baru = str(nip_baru).strip()

    data = df.copy()

    if index_diabaikan is not None and index_diabaikan in data.index:
        data = data.drop(index=index_diabaikan)

    return nip_baru not in data["NIP BARU"].astype(str).str.strip().values


def buat_display_pegawai(row):
    """Format tampilan pegawai pada selectbox."""
    nama = row.get("NAMA", "-")
    nip = row.get("NIP BARU", "-")
    jabatan = row.get("NAMA JABATAN", "-")

    return f"{nama} | NIP: {nip} | {jabatan}"


def form_input_pegawai(df):
    """
    Fitur 1: Input pegawai baru.
    """
    st.subheader("Input Pegawai Baru")
    st.caption("Gunakan fitur ini untuk menambahkan pegawai baru ke database master.")

    list_agama = opsi_unik(df, "AGAMA", ["Islam", "Kristen", "Katolik", "Hindu", "Buddha", "Konghucu"])
    list_status = opsi_unik(df, "STATUS PNS/CPNS", ["PNS", "CPNS", "PPPK", "Non ASN"])
    list_jenjang = opsi_unik(df, "JENJANG", ["SMA", "D3", "S1", "S2", "S3"])
    list_jabatan = opsi_unik(df, "NAMA JABATAN", ["-"])
    list_unit = opsi_unik(df, "UNIT KERJA", ["BIRO UMUM"])

    with st.form("form_input_pegawai_baru", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            nama = st.text_input("Nama Pegawai *")
            nip = st.text_input("NIP Baru *")
            nik = st.text_input("NIK")
            jk = st.selectbox("Jenis Kelamin *", ["L", "P"])
            agama = st.selectbox("Agama", list_agama)
            status = st.selectbox("Status PNS/CPNS", list_status)

        with col2:
            jenjang = st.selectbox("Jenjang Pendidikan", list_jenjang)
            jabatan = st.selectbox("Nama Jabatan", list_jabatan)
            unit = st.selectbox("Unit Kerja", list_unit)
            tanggal_lahir = st.date_input(
                "Tanggal Lahir",
                value=date(1990, 1, 1),
                format="DD/MM/YYYY"
            )
            tmt_masuk_kerja = st.date_input(
                "TMT Masuk Kerja",
                value=date.today(),
                format="DD/MM/YYYY"
            )
            tmt_pensiun = st.date_input(
                "TMT Pensiun",
                value=date(2050, 1, 1),
                format="DD/MM/YYYY"
            )

        simpan = st.form_submit_button("Simpan Pegawai Baru", use_container_width=True)

    if simpan:
        nama = bersihkan_teks(nama)
        nip = bersihkan_teks(nip)

        if nama == "-" or nip == "-":
            st.error("Nama Pegawai dan NIP Baru wajib diisi.")
            return

        if not validasi_nip_unik(df, nip):
            st.error("NIP Baru sudah terdaftar. Gunakan NIP yang berbeda.")
            return

        df_baru = df.copy()

        row_baru = {col: "-" for col in df_baru.columns}

        row_baru.update({
            "NO": len(df_baru) + 1,
            "NAMA": nama.upper(),
            "NIP BARU": nip,
            "NIK": bersihkan_teks(nik),
            "JK": jk,
            "AGAMA": agama,
            "STATUS PNS/CPNS": status,
            "JENJANG": jenjang,
            "NAMA JABATAN": jabatan,
            "UNIT KERJA": unit,
            "TGL LAHIR": format_tanggal_csv(tanggal_lahir),
            "USIA": hitung_usia_dari_tanggal_lahir(tanggal_lahir),
            "TMT MASUK KERJA": format_tanggal_csv(tmt_masuk_kerja),
            "TMT PENSIUN": format_tanggal_csv(tmt_pensiun),
            "TAHUN PENSIUN": str(tmt_pensiun.year),
        })

        df_baru = pd.concat(
            [df_baru, pd.DataFrame([row_baru])],
            ignore_index=True
        )

        simpan_data_pegawai_kelola(df_baru)
        catat_aktivitas("Input Pegawai")

        st.success(f"Pegawai {nama.upper()} berhasil ditambahkan dan otomatis muncul di Menu 1.")
        st.rerun()


def cari_pegawai_dataframe(df, keyword):
    """
    Mencari pegawai berdasarkan Nama, NIP, NIK, Jabatan, atau Unit Kerja.
    """
    if df.empty:
        return df

    keyword = str(keyword).strip().lower()

    if not keyword:
        return pd.DataFrame()

    kolom_pencarian = [
        "NAMA",
        "NIP BARU",
        "NIK",
        "NAMA JABATAN",
        "UNIT KERJA",
        "STATUS PNS/CPNS",
        "JENJANG",
    ]

    mask = pd.Series(False, index=df.index)

    for kolom in kolom_pencarian:
        if kolom in df.columns:
            mask = mask | df[kolom].astype(str).str.lower().str.contains(keyword, na=False)

    return df[mask]

def form_update_pegawai(df):
    """
    Fitur 2: Update data pegawai yang sudah ada.
    """
    st.subheader("Update Pegawai")
    st.caption("Gunakan fitur ini untuk memperbarui informasi pegawai yang sudah ada.")

    if df.empty:
        st.warning("Data pegawai masih kosong.")
        return

    df = df.copy()

    df = df.copy()

    df = df.copy()

    keyword_update = st.text_input(
        "Cari pegawai yang akan diupdate",
        placeholder="Ketik Nama / NIP / NIK / Jabatan / Unit Kerja",
        key="keyword_update_pegawai"
    )

    hasil_cari = cari_pegawai_dataframe(df, keyword_update)

    if keyword_update and hasil_cari.empty:
        st.warning("Pegawai tidak ditemukan. Coba gunakan kata kunci lain.")
        return

    if not keyword_update:
        st.info("Silakan ketik nama, NIP, NIK, jabatan, atau unit kerja pegawai terlebih dahulu.")
        return

    st.caption(f"Ditemukan {len(hasil_cari)} pegawai.")

    pilihan = {
        buat_display_pegawai(row): idx
        for idx, row in hasil_cari.iterrows()
    }

    pegawai_dipilih = st.selectbox(
        "Pilih dari hasil pencarian",
        list(pilihan.keys()),
        key="pilih_update_pegawai"
    )

    index_pegawai = pilihan[pegawai_dipilih]
    row = df.loc[index_pegawai]

    list_agama = opsi_unik(df, "AGAMA", ["Islam", "Kristen", "Katolik", "Hindu", "Buddha", "Konghucu"])
    list_status = opsi_unik(df, "STATUS PNS/CPNS", ["PNS", "CPNS", "PPPK", "Non ASN"])
    list_jenjang = opsi_unik(df, "JENJANG", ["SMA", "D3", "S1", "S2", "S3"])
    list_jabatan = opsi_unik(df, "NAMA JABATAN", ["-"])
    list_unit = opsi_unik(df, "UNIT KERJA", ["BIRO UMUM"])

    def index_opsi(list_opsi, nilai):
        nilai = str(nilai).strip()
        return list_opsi.index(nilai) if nilai in list_opsi else 0

    with st.form("form_update_pegawai"):
        col1, col2 = st.columns(2)

        with col1:
            nama = st.text_input("Nama Pegawai *", value=row.get("NAMA", "-"))
            nip = st.text_input("NIP Baru *", value=row.get("NIP BARU", "-"))
            nik = st.text_input("NIK", value=row.get("NIK", "-"))

            jk_lama = row.get("JK", "L")
            jk = st.selectbox(
                "Jenis Kelamin *",
                ["L", "P"],
                index=0 if jk_lama == "L" else 1
            )

            agama = st.selectbox(
                "Agama",
                list_agama,
                index=index_opsi(list_agama, row.get("AGAMA", "-"))
            )

            status = st.selectbox(
                "Status PNS/CPNS",
                list_status,
                index=index_opsi(list_status, row.get("STATUS PNS/CPNS", "-"))
            )

        with col2:
            jenjang = st.selectbox(
                "Jenjang Pendidikan",
                list_jenjang,
                index=index_opsi(list_jenjang, row.get("JENJANG", "-"))
            )

            jabatan = st.selectbox(
                "Nama Jabatan",
                list_jabatan,
                index=index_opsi(list_jabatan, row.get("NAMA JABATAN", "-"))
            )

            unit = st.selectbox(
                "Unit Kerja",
                list_unit,
                index=index_opsi(list_unit, row.get("UNIT KERJA", "-"))
            )

            tanggal_lahir = st.date_input(
                "Tanggal Lahir",
                value=default_date_input(row.get("TGL LAHIR", "-"), date(1990, 1, 1)),
                format="DD/MM/YYYY"
            )

            tmt_masuk_kerja = st.date_input(
                "TMT Masuk Kerja",
                value=default_date_input(row.get("TMT MASUK KERJA", "-"), date.today()),
                format="DD/MM/YYYY"
            )

            tmt_pensiun = st.date_input(
                "TMT Pensiun",
                value=default_date_input(row.get("TMT PENSIUN", "-"), date(2050, 1, 1)),
                format="DD/MM/YYYY"
            )

        update = st.form_submit_button("Update Data Pegawai", use_container_width=True)

    if update:
        nama = bersihkan_teks(nama)
        nip = bersihkan_teks(nip)

        if nama == "-" or nip == "-":
            st.error("Nama Pegawai dan NIP Baru wajib diisi.")
            return

        if not validasi_nip_unik(df, nip, index_diabaikan=index_pegawai):
            st.error("NIP Baru sudah digunakan oleh pegawai lain.")
            return

        df.loc[index_pegawai, "NAMA"] = nama.upper()
        df.loc[index_pegawai, "NIP BARU"] = nip
        df.loc[index_pegawai, "NIK"] = bersihkan_teks(nik)
        df.loc[index_pegawai, "JK"] = jk
        df.loc[index_pegawai, "AGAMA"] = agama
        df.loc[index_pegawai, "STATUS PNS/CPNS"] = status
        df.loc[index_pegawai, "JENJANG"] = jenjang
        df.loc[index_pegawai, "NAMA JABATAN"] = jabatan
        df.loc[index_pegawai, "UNIT KERJA"] = unit
        df.loc[index_pegawai, "TGL LAHIR"] = format_tanggal_csv(tanggal_lahir)
        df.loc[index_pegawai, "USIA"] = hitung_usia_dari_tanggal_lahir(tanggal_lahir)
        df.loc[index_pegawai, "TMT MASUK KERJA"] = format_tanggal_csv(tmt_masuk_kerja)
        df.loc[index_pegawai, "TMT PENSIUN"] = format_tanggal_csv(tmt_pensiun)
        df.loc[index_pegawai, "TAHUN PENSIUN"] = str(tmt_pensiun.year)

        simpan_data_pegawai_kelola(df)
        catat_aktivitas("Update Pegawai")

        st.success(f"Data pegawai {nama.upper()} berhasil diperbarui dan otomatis berubah di Menu 1.")
        st.rerun()


def form_hapus_pegawai(df):
    """
    Fitur 3: Hapus pegawai dari data aktif.
    Data pegawai dipindahkan ke Arsip Pegawai.
    """
    st.subheader("Hapus Pegawai")
    st.caption(
        "Gunakan fitur ini untuk pegawai yang resign sebelum pensiun atau berpindah biro "
        "sehingga tidak lagi tampil di Menu 1."
    )

    if df.empty:
        st.warning("Data pegawai masih kosong.")
        return

    df = df.copy()

    keyword_hapus = st.text_input(
        "Cari pegawai yang akan dihapus",
        placeholder="Ketik Nama / NIP / NIK / Jabatan / Unit Kerja",
        key="keyword_hapus_pegawai"
    )

    hasil_cari = cari_pegawai_dataframe(df, keyword_hapus)

    if keyword_hapus and hasil_cari.empty:
        st.warning("Pegawai tidak ditemukan. Coba gunakan kata kunci lain.")
        return

    if not keyword_hapus:
        st.info("Silakan ketik nama, NIP, NIK, jabatan, atau unit kerja pegawai terlebih dahulu.")
        return

    st.caption(f"Ditemukan {len(hasil_cari)} pegawai.")

    pilihan = {
        buat_display_pegawai(row): idx
        for idx, row in hasil_cari.iterrows()
    }

    pegawai_dipilih = st.selectbox(
        "Pilih dari hasil pencarian",
        list(pilihan.keys()),
        key="hapus_pegawai_selectbox"
    )

    index_pegawai = pilihan[pegawai_dipilih]
    row = df.loc[index_pegawai]

    st.warning(
        f"Pegawai yang dipilih: **{row.get('NAMA', '-')}** "
        f"dengan NIP **{row.get('NIP BARU', '-')}**."
    )

    with st.form("form_hapus_pegawai"):
        alasan = st.selectbox(
            "Alasan penghapusan",
            [
                "Resign sebelum tenggat pensiun",
                "Pindah biro / bukan di Biro Umum",
                "Pensiun",
                "Alasan lainnya",
            ]
        )

        tanggal_hapus = st.date_input(
            "Tanggal efektif",
            value=date.today(),
            format="DD/MM/YYYY"
        )

        catatan = st.text_area(
            "Catatan tambahan",
            placeholder="Contoh: pindah ke biro lain per tanggal ...",
        )

        konfirmasi = st.checkbox(
            "Saya yakin ingin menghapus pegawai ini dari data aktif."
        )

        hapus = st.form_submit_button("Hapus dari Data Aktif", use_container_width=True)

    if hapus:
        if not konfirmasi:
            st.error("Centang konfirmasi terlebih dahulu sebelum menghapus pegawai.")
            return

        arsipkan_pegawai(
            row_pegawai=row.to_dict(),
            status="DIHAPUS",
            alasan=alasan,
            tanggal_arsip=tanggal_hapus,
            catatan=catatan
        )

        df = df.drop(index=index_pegawai).reset_index(drop=True)

        simpan_data_pegawai_kelola(df)
        catat_aktivitas("Keluarkan Pegawai")

        st.success(
            f"Pegawai {row.get('NAMA', '-')} berhasil dihapus dari data aktif "
            "dan dipindahkan ke Arsip Pegawai."
        )
        st.rerun()


def menu_kelola_pegawai(df: pd.DataFrame, df_seed: pd.DataFrame = None):
    """
    Halaman utama Menu 3: Kelola Pegawai.
    """
    page_header(
        "Kelola Pegawai",
        "Update, input, dan hapus data pegawai aktif Biro Umum"
    )

    st.info(
        "Semua perubahan pada menu ini akan disimpan ke database operasional pegawai. "
        "Menu 1 akan otomatis menampilkan data terbaru setelah proses berhasil."
    )

    df = normalisasi_data_pegawai(df)

    tab_update, tab_input, tab_hapus = st.tabs([
        "Update Pegawai",
        "Input Pegawai",
        "Hapus Pegawai",
    ])

    with tab_update:
        form_update_pegawai(df)

    with tab_input:
        form_input_pegawai(df)

    with tab_hapus:
        form_hapus_pegawai(df)




# =====================================================================
# =====================================================================
# MENU 4: KPI JAM KERJA  (MENGIKUTI SPESIFIKASI RESMI)
# ---------------------------------------------------------------------
# Komponen KPI Jam Kerja sesuai gambar yang Anda kirim:
#   [I]   Ijin               - Gabungan semua ijin (tidak dipotong)
#   [TL]  Terlambat          - Rp 35.000 / kejadian
#   [PC]  Pulang Cepat       - Rp 35.000 / kejadian
#   [TAD] Tidak Absen Datang - Rp 35.000 / kejadian
#   [TAP] Tidak Absen Pulang - Rp 35.000 / kejadian
#   [A]   Alpha              - Rp 35.000 / kejadian
#   [SN]  Tidak Ikut Senam   - Rp 35.000 / kejadian
#   [TSN] Terlambat Senam    - Rp 35.000 / kejadian
#   [AP]  Tidak Ikut Apel    - Rp 35.000 / kejadian
#   [TOT] Total Potongan / Hari = Rp 135.000 (cap)
#   [LIB] Libur               - tidak dihitung
#
# Skor Jam Kerja = 100 - (Total Potongan / Potongan Maksimal Bulanan) * 100
# Status Tercapai bila Skor >= 80.
# =====================================================================
# =====================================================================
def menu_kpi_jam_kerja(kpi: pd.DataFrame):
    page_header(
        "KPI Jam Kerja",
        "Pencapaian disiplin jam kerja berbasis data kehadiran real"
    )

    st.info(
        "Upload file kehadiran bulanan. "
        "Jika periode yang sama diupload kembali, data periode tersebut akan diganti dengan data terbaru."
    )

    uploaded_files = st.file_uploader(
        "Upload file raw kehadiran",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="Contoh: SKP KEHADIRAN JANUARI.xlsx, SKP KEHADIRAN FEBRUARI.xlsx, dan seterusnya.",
    )

    col_upload1, col_upload2 = st.columns([1, 3])

    with col_upload1:
        proses = st.button("Proses Upload", use_container_width=True)

    if proses:
        if not uploaded_files:
            st.warning("Silakan upload minimal 1 file raw kehadiran terlebih dahulu.")
            st.stop()

        try:
            with st.spinner("Memproses file kehadiran..."):
                hasil = proses_upload_kehadiran(
                    uploaded_files,
                    output_path=KEHADIRAN_PATH
                )
                st.cache_data.clear()

            st.success(
                f"Cleaning & preprocessing selesai.\n"
                f"Total data harian: {len(hasil):,} baris."
            )


        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
            st.stop()

    # =========================================================
    # LOAD DATA KEHADIRAN FINAL
    # =========================================================
    df_kehadiran = load_kehadiran_final()

    if df_kehadiran.empty:
        st.warning(
            "Database kehadiran belum tersedia.\n"
            "Silakan upload minimal satu file kehadiran terlebih dahulu."
        )
        return

    df_kehadiran = df_kehadiran.copy()
    df_kehadiran["tanggal"] = pd.to_datetime(df_kehadiran["tanggal"], errors="coerce")
    df_kehadiran = df_kehadiran.dropna(subset=["tanggal"])

    if df_kehadiran.empty:
        st.warning("Kolom tanggal pada kehadiran_final.csv tidak valid.")
        return

    # =========================================================
    # TABEL REFERENSI TARIF
    # =========================================================
    with st.expander("Spesifikasi Komponen & Tarif Potongan", expanded=False):
        ref = pd.DataFrame([
            {
                "Kode": k,
                "Komponen": v["label"],
                "Tarif (Rp)": f"{v['tarif']:,}".replace(",", "."),
            }
            for k, v in KPI_KOMPONEN.items()
        ])

        st.table(ref)

        st.caption(
            f"[TOT] Total potongan maksimal per hari: "
            f"Rp {POTONGAN_MAKSIMAL_HARIAN:,}".replace(",", ".")
        )

        st.caption(
            "Catatan: TL dihitung berdasarkan kelipatan 10 menit keterlambatan. "
            "Sabtu dan Minggu tidak dihitung sebagai hari kerja. "
            "Tanggal merah tetap dianggap masuk jika jatuh pada Senin-Jumat."
        )

    # =========================================================
    # FILTER PERIODE DATA KEHADIRAN
    # =========================================================
    st.markdown("---")
    st.subheader("Filter Periode Data")

    data_kehadiran = df_kehadiran.copy()
    data_kehadiran["tahun"] = data_kehadiran["tanggal"].dt.year
    data_kehadiran["bulan"] = data_kehadiran["tanggal"].dt.month
    data_kehadiran["bulan_tahun"] = (
        data_kehadiran["bulan"].map(BULAN_ID)
        + " "
        + data_kehadiran["tahun"].astype(str)
    )

    mode_filter = st.radio(
        "Mode tampilan data:",
        ["Semua Data", "Pilih Bulan", "Pilih Rentang Tanggal"],
        horizontal=True,
        key="mode_filter_jam_kerja"
    )

    filtered_kehadiran = data_kehadiran.copy()

    if mode_filter == "Pilih Bulan":
        bulan_opsi = (
            data_kehadiran[["tahun", "bulan", "bulan_tahun"]]
            .drop_duplicates()
            .sort_values(["tahun", "bulan"])
        )

        bulan_pilih = st.selectbox(
            "Pilih bulan",
            bulan_opsi["bulan_tahun"].tolist(),
            key="pilih_bulan_jam_kerja"
        )

        filtered_kehadiran = data_kehadiran[
            data_kehadiran["bulan_tahun"] == bulan_pilih
        ]

    elif mode_filter == "Pilih Rentang Tanggal":
        min_date = data_kehadiran["tanggal"].min()
        max_date = data_kehadiran["tanggal"].max()

        periode = st.date_input(
            "Pilih rentang tanggal",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date(),
            key="periode_jam_kerja"
        )

        if isinstance(periode, tuple) and len(periode) == 2:
            start_date, end_date = periode

            filtered_kehadiran = data_kehadiran[
                (data_kehadiran["tanggal"] >= pd.to_datetime(start_date))
                & (data_kehadiran["tanggal"] <= pd.to_datetime(end_date))
            ]

    # Hitung ulang KPI berdasarkan periode yang dipilih
    kpi = ubah_kehadiran_ke_kpi(filtered_kehadiran)

    if kpi.empty:
        st.warning("Data KPI tidak tersedia untuk periode yang dipilih.")
        st.stop()

    # =========================================================
    # KPI RINGKAS
    # =========================================================
    st.markdown("---")
    total = len(kpi)

    tercapai = kpi["Kategori_Jam_Kerja"].isin(
        ["Sesuai Ekspektasi", "Diatas Ekspektasi"]
    ).sum()

    tidak = total - tercapai
    avg_skor = round(kpi["Skor_Jam_Kerja"].mean(), 1)
    total_pot = int(kpi["Potongan_Total"].sum())

    c1, c2, c3, c4, c5 = st.columns(5)

    kpi_card(c1, "Total Pegawai", total, font_size=24)
    kpi_card(c2, "Tercapai", tercapai, "green", font_size=24)
    kpi_card(c3, "Tidak Tercapai", tidak, "red", font_size=24)
    kpi_card(c4, "Rata² Skor", f"{avg_skor}", font_size=24)
    kpi_card(
        c5,
        "Total Potongan",
        f"Rp {total_pot:,}".replace(",", "."),
        "orange",
        font_size=18
    )

    # =========================================================
    # TOTAL KEJADIAN PER KOMPONEN
    # =========================================================
    st.markdown("---")
    st.subheader("Total Kejadian Pelanggaran per Komponen")

    rekap = []

    for kode, info in KPI_KOMPONEN.items():
        col = f"K_{kode}"

        if col not in kpi.columns:
            continue

        jml = int(kpi[col].sum())
        pot = jml * info["tarif"]

        rekap.append({
            "Kode": kode,
            "Komponen": info["label"],
            "Total Kejadian": jml,
            "Total Potongan": pot,
        })

    rekap_df = pd.DataFrame(rekap)

    plot_df = rekap_df[
        rekap_df["Komponen"].isin(
            [v["label"] for k, v in KPI_KOMPONEN.items() if v["tarif"] > 0]
        )
    ]

    if plot_df.empty:
        st.info("Belum ada data pelanggaran berpotongan untuk ditampilkan.")
    else:
        fig = px.bar(
            plot_df.sort_values("Total Kejadian"),
            x="Total Kejadian",
            y="Komponen",
            orientation="h",
            text="Total Kejadian",
            color="Total Kejadian",
            color_continuous_scale="Reds",
            title="Frekuensi Kejadian per Komponen Pelanggaran",
        )

        fig.update_traces(textposition="outside")
        fig.update_layout(height=520)

        st.plotly_chart(fig, use_container_width=True)

    view_rekap = rekap_df.copy()
    view_rekap["Total Potongan"] = view_rekap["Total Potongan"].apply(
        lambda x: f"Rp {x:,}".replace(",", ".")
    )

    st.dataframe(
        view_rekap,
        use_container_width=True,
        height=360,
        hide_index=True
    )

    # =========================================================
    # DISTRIBUSI STATUS & TOP POTONGAN
    # =========================================================
    st.markdown("---")

    g3, g4 = st.columns([1, 2])

    with g3:
        pie_df = kpi["Kategori_Jam_Kerja"].value_counts().reset_index()
        pie_df.columns = ["Status", "Jumlah"]

        fig = px.pie(
            pie_df,
            names="Status",
            values="Jumlah",
            title="Status KPI Jam Kerja",
            hole=0.45,
            color="Status",
            color_discrete_map={
                "Diatas Ekspektasi": "#27ae60",
                "Sesuai Ekspektasi": "#f1c40f",
                "Dibawah Ekspektasi": "#e74c3c",
            },
        )

        st.plotly_chart(fig, use_container_width=True)

    with g4:
        top = kpi.nlargest(15, "Potongan_Total")[
            ["NAMA", "Potongan_Total", "Skor_Jam_Kerja"]
        ]

        fig = px.bar(
            top,
            x="Potongan_Total",
            y="NAMA",
            orientation="h",
            color="Potongan_Total",
            color_continuous_scale="Reds",
            text="Potongan_Total",
            title="15 Pegawai dengan Potongan Tertinggi",
        )

        fig.update_traces(texttemplate="Rp %{text:,.0f}", textposition="outside")
        fig.update_layout(
            yaxis={"categoryorder": "total ascending"},
            height=520
        )

        st.plotly_chart(fig, use_container_width=True)

    #---------DETAIL PER PEGAWAI---------#
    st.markdown("---")
    st.subheader("Detail KPI Jam Kerja per Pegawai")

    nama = st.selectbox(
        "Pilih pegawai",
        kpi["NAMA"].tolist(),
        key="sel_jam_detail"
    )

    row = kpi[kpi["NAMA"] == nama].iloc[0]

    p1, p2, p3, p4 = st.columns(4)

    kpi_card(p1, "Skor Jam Kerja", f"{row['Skor_Jam_Kerja']}")
    kpi_card(
        p2,
        "Total Potongan",
        f"Rp {int(row['Potongan_Total']):,}".replace(",", "."),
        "orange",
    )
    kpi_card(p3, "Status", row["Kategori_Jam_Kerja"])

    total_pelanggaran = sum(
        int(row[f"K_{k}"])
        for k, v in KPI_KOMPONEN.items()
        if v["tarif"] > 0 and f"K_{k}" in kpi.columns
    )

    kpi_card(p4, "Total Kejadian", total_pelanggaran)

    rinc = []

    for kode, info in KPI_KOMPONEN.items():
        col = f"K_{kode}"

        if col not in kpi.columns:
            continue

        rinc.append({
            "Komponen": f"[{kode}] {info['label']}",
            "Kejadian": int(row[col]),
            "Tarif (Rp)": info["tarif"],
            "Potongan (Rp)": int(row[col]) * info["tarif"],
        })

    rinc_df = pd.DataFrame(rinc)

    fig = px.bar(
        rinc_df,
        x="Komponen",
        y="Kejadian",
        color="Potongan (Rp)",
        color_continuous_scale="Reds",
        text="Kejadian",
        title=f"Rincian Komponen - {nama}",
    )

    fig.update_traces(textposition="outside")
    fig.update_layout(xaxis_tickangle=-30, height=520)

    st.plotly_chart(fig, use_container_width=True)

    view_rinc = rinc_df.copy()
    view_rinc["Tarif (Rp)"] = view_rinc["Tarif (Rp)"].apply(
        lambda x: f"Rp {x:,}".replace(",", ".")
    )
    view_rinc["Potongan (Rp)"] = view_rinc["Potongan (Rp)"].apply(
        lambda x: f"Rp {x:,}".replace(",", ".")
    )

    st.dataframe(
        view_rinc,
        use_container_width=True,
        hide_index=True
    )

    #--------REKAP SELURUH PEGAWAI---------#
    st.markdown("---")
    st.subheader("Rekap KPI Jam Kerja Seluruh Pegawai")

    cari = st.text_input(
        "🔍 Cari nama",
        key="cari_jam_rekap_kpi"
    )

    show = kpi.copy()

    if cari:
        show = show[
            show["NAMA"].str.lower().str.contains(cari.lower(), na=False)
        ]

    cols_show = (
        ["NAMA", "OPD", "Hari_Data", "Potongan_Maksimal_Periode"]
        + [f"K_{k}" for k in KPI_KOMPONEN if f"K_{k}" in kpi.columns]
        + ["Potongan_Total", "Skor_Jam_Kerja", "Kategori_Jam_Kerja"]
    )

    rename_map = {f"K_{k}": k for k in KPI_KOMPONEN}
    rename_map.update({
        "Potongan_Maksimal_Periode": "Maksimal Potongan Periode",
        "Potongan_Total": "Potongan (Rp)",
        "Skor_Jam_Kerja": "Skor",
        "Kategori_Jam_Kerja": "Status",
        "Hari_Data": "Jumlah Hari Kerja",
    })

    st.dataframe(
        show[cols_show].rename(columns=rename_map),
        use_container_width=True,
        height=420,
    )

    csv_kpi = (
        show[cols_show]
        .rename(columns=rename_map)
        .to_csv(index=False)
        .encode("utf-8-sig")
    )

    st.download_button(
        "⬇️ Download Rekap KPI Jam Kerja",
        csv_kpi,
        file_name="rekap_kpi_jam_kerja.csv",
        mime="text/csv",
        on_click=lambda: catat_aktivitas("Download Rekap KPI Jam Kerja")
    )

# =====================================================================
# FUNGSI: Menyaring data KPI agar hanya menampilkan pegawai aktif
# ---------------------------------------------------------------------
# Data 7 Aspek/KPI diupload terpisah (CSV/XLSX) dan tidak otomatis
# tersinkron dengan status pegawai (pensiun/dihapus/diarsipkan).
# Fungsi ini mencocokkan NAMA di data KPI dengan NAMA pegawai yang
# masih aktif (memakai bersihkan_nama_matching, sama seperti pencocokan
# nama yang sudah dipakai di fitur lain) lalu membuang baris pegawai
# yang sudah tidak aktif.
# =====================================================================
def filter_kpi_pegawai_aktif(kpi_df, pegawai_aktif_df):
    if kpi_df.empty or pegawai_aktif_df.empty:
        return kpi_df

    if "NAMA" not in kpi_df.columns or "NAMA" not in pegawai_aktif_df.columns:
        return kpi_df

    nama_aktif = set(
        pegawai_aktif_df["NAMA"].astype(str).apply(bersihkan_nama_matching)
    )

    data = kpi_df.copy()
    data["_NAMA_MATCH"] = data["NAMA"].astype(str).apply(bersihkan_nama_matching)
    data = data[data["_NAMA_MATCH"].isin(nama_aktif)]

    return data.drop(columns=["_NAMA_MATCH"])


# =====================================================================
# =====================================================================
# MENU 6: KPI 7 ASPEK ASN 
# =====================================================================
# =====================================================================
def menu_kpi_7aspek(kpi: pd.DataFrame, pegawai_aktif: pd.DataFrame):
    page_header(
        "KPI 7 Aspek ASN",
        "Penilaian 7 aspek ASN berbasis data real CSV/XLSX yang bisa terus diupdate"
    )

    st.info(
        "Upload file penilaian 7 Aspek ASN di sini. "
        "Data baru akan ditambahkan ke data lama. "
        "Jika NIP + NAMA + Periode sama, data terbaru akan menggantikan data lama."
    )

    uploaded_files = st.file_uploader(
        "Upload file 7 Aspek ASN",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        help="Kolom wajib: NAMA, Orientasi_Pelayanan, Akuntabel, Kompeten, Harmonis, Loyal, Adaptif, Kolaboratif. Kolom opsional: NIP, Periode."
    )

    if st.button("Proses Upload 7 Aspek", use_container_width=True):
        if not uploaded_files:
            st.warning("Silakan upload minimal 1 file terlebih dahulu.")
            st.stop()

        try:
            with st.spinner("Memproses data 7 Aspek ASN..."):
                hasil = proses_upload_7aspek(uploaded_files, ASPEK_7_PATH)
                st.cache_data.clear()

            st.success(
                f"Data KPI 7 Aspek berhasil diperbarui.\n"
                f"Total data: {len(hasil):,} baris."
            )

        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
            st.stop()

    kpi = load_7aspek_final(ASPEK_7_PATH)

    # Saring hanya pegawai yang masih aktif (bukan pensiun/dihapus/diarsipkan)
    kpi = filter_kpi_pegawai_aktif(kpi, pegawai_aktif)

    if kpi.empty:
        st.warning(
            "Database KPI 7 Aspek belum tersedia.\n"
            "Silakan upload minimal satu file terlebih dahulu."
        )
        return


    #------FILTER PERIODE DATA 7 ASPEK--------#
    st.markdown("---")
    st.subheader("Filter Periode Data")

    data_7 = kpi.copy()

    data_7["tanggal_periode"] = pd.to_datetime(
        data_7["Periode"].astype(str) + "-01",
        errors="coerce"
    )

    data_7 = data_7.dropna(subset=["tanggal_periode"])

    if data_7.empty:
        st.warning("Kolom Periode tidak valid. Gunakan format seperti 2026-01.")
        st.stop()

    data_7["tahun"] = data_7["tanggal_periode"].dt.year
    data_7["bulan"] = data_7["tanggal_periode"].dt.month
    data_7["bulan_tahun"] = (
        data_7["bulan"].map(BULAN_ID)
        + " "
        + data_7["tahun"].astype(str)
    )

    mode_filter = st.radio(
        "Mode tampilan data:",
        ["Semua Data", "Pilih Bulan"],
        horizontal=True,
        key="mode_filter_7aspek"
    )

    filtered_7 = data_7.copy()

    if mode_filter == "Pilih Bulan":
        bulan_opsi = (
            data_7[["tahun", "bulan", "bulan_tahun"]]
            .drop_duplicates()
            .sort_values(["tahun", "bulan"])
        )

        bulan_pilih = st.selectbox(
            "Pilih bulan",
            bulan_opsi["bulan_tahun"].tolist(),
            key="pilih_bulan_7aspek"
        )

        filtered_7 = data_7[
            data_7["bulan_tahun"] == bulan_pilih
        ]

    kpi = filtered_7.copy()

    if kpi.empty:
        st.warning("Data KPI 7 Aspek tidak tersedia untuk bulan yang dipilih.")
        st.stop()

    # Jika mode "Semua Data", satu pegawai bisa muncul berkali-kali
    # karena ada data Januari, Februari, Maret, April.
    # Maka total pegawai harus dihitung berdasarkan NIP unik, bukan jumlah baris.
    total = kpi["NIP"].nunique() if "NIP" in kpi.columns else kpi["NAMA"].nunique()

    # Ambil data terakhir per pegawai agar status tidak dihitung dobel.
    # Contoh: 1 pegawai punya data Jan-Apr, yang dipakai untuk ringkasan adalah data periode terakhir.
    kpi_ringkas = (
        kpi.sort_values("Periode")
        .drop_duplicates(subset=["NIP"], keep="last")
        if "NIP" in kpi.columns
        else kpi.sort_values("Periode").drop_duplicates(subset=["NAMA"], keep="last")
    )

    tercapai = kpi_ringkas["Kategori_7_Aspek"].isin(
        ["Sesuai Ekspektasi", "Diatas Ekspektasi"]
    ).sum()

    tidak = (kpi_ringkas["Kategori_7_Aspek"] == "Dibawah Ekspektasi").sum()

    rata2 = round(kpi_ringkas["Rata_7_Aspek"].mean(), 1)

    c1, c2, c3, c4 = st.columns(4)
    kpi_card(c1, "Total Pegawai", total, font_size=24)
    kpi_card(c2, "Tercapai", tercapai, font_size=24)
    kpi_card(c3, "Dibawah Ekspektasi", tidak, font_size=24)
    kpi_card(c4, "Rata-rata Skor", rata2, font_size=24)

    st.markdown("---")

    rata_aspek = kpi[ASPEK_7_COLS].mean().round(1)

    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=rata_aspek.values,
        theta=[a.replace("_", " ") for a in ASPEK_7_COLS],
        fill="toself",
        name="Rata-rata Skor",
        line=dict(color="#1e3c72")
    ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        title="Rata-rata Skor per Aspek ASN",
        showlegend=False
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    g1, g2 = st.columns([1, 2])

    with g1:
        pie_df = kpi["Kategori_7_Aspek"].value_counts().reset_index()
        pie_df.columns = ["Status", "Jumlah"]

        fig = px.pie(
            pie_df,
            names="Status",
            values="Jumlah",
            title="Status KPI 7 Aspek ASN",
            hole=0.45,
            color="Status",
            color_discrete_map={
                "Diatas Ekspektasi": "#27ae60",
                "Sesuai Ekspektasi": "#f1c40f",
                "Dibawah Ekspektasi": "#e74c3c",
            },
        )
        fig.update_layout(height=480)
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        # Pakai kpi_ringkas (1 baris per pegawai) agar tidak ada
        # pegawai yang tampil >1x per bar (penyebab bar 2 warna/tumpuk)
        bottom = kpi_ringkas.nsmallest(15, "Rata_7_Aspek")[["NAMA", "Periode", "Rata_7_Aspek", "Kategori_7_Aspek"]]

        fig = px.bar(
            bottom,
            x="Rata_7_Aspek",
            y="NAMA",
            orientation="h",
            color="Rata_7_Aspek",
            color_continuous_scale="Reds",
            text="Rata_7_Aspek",
            title="15 Pegawai dengan Skor 7 Aspek Terendah",
        )
        fig.update_layout(
            yaxis={"categoryorder": "total ascending"},
            height=480
        )
        st.plotly_chart(fig, use_container_width=True)

    
    #------DETAIL 7 ASPEK PER PEGAWAI--------#
    st.markdown("---")
    st.subheader("Detail 7 Aspek per Pegawai")

    # Selectbox untuk memilih pegawai yang ingin dilihat detailnya
    nama = st.selectbox(
        "Pilih pegawai",
        kpi["NAMA"].dropna().astype(str).sort_values().unique().tolist(),
        key="sel_7aspek"
    )

    # Ambil seluruh data pegawai terpilih.
    # Jika mode Semua Data aktif, pegawai bisa punya beberapa baris periode.
    data_pegawai = kpi[kpi["NAMA"] == nama].copy()

    # Pilihan mode detail:
    # 1. Per Bulan = melihat nilai pegawai pada bulan tertentu
    # 2. Rata-rata Keseluruhan = menghitung rata-rata seluruh bulan yang tersedia
    mode_detail = st.radio(
        "Mode detail pegawai:",
        ["Per Bulan", "Rata-rata Keseluruhan"],
        horizontal=True,
        key="mode_detail_7aspek"
    )

    if mode_detail == "Per Bulan":
        # Jika data pegawai lebih dari satu periode,
        # user bisa memilih periode detail yang ingin dilihat.
        if len(data_pegawai) > 1:
            periode_detail = st.selectbox(
                "Pilih periode detail",
                sorted(data_pegawai["Periode"].astype(str).unique().tolist()),
                key="periode_detail_7aspek"
            )

            # Ambil satu baris data sesuai periode yang dipilih
            row = data_pegawai[
                data_pegawai["Periode"].astype(str) == periode_detail
            ].iloc[0]
        else:
            # Jika hanya ada satu periode, langsung ambil baris pertama
            row = data_pegawai.iloc[0]

        # Buat dataframe detail aspek untuk grafik batang
        df_one = pd.DataFrame({
            "Aspek": [a.replace("_", " ") for a in ASPEK_7_COLS],
            "Skor": [row[a] for a in ASPEK_7_COLS]
        })

        # Nilai ringkasan untuk KPI card
        nilai_rata = row["Rata_7_Aspek"]
        status_nilai = row["Kategori_7_Aspek"]
        label_periode = row.get("Periode", "-")
        judul_grafik = f"Skor 7 Aspek ASN - {nama} - {label_periode}"

    else:
        # Mode Rata-rata Keseluruhan:
        # Menghitung rata-rata setiap aspek dari seluruh periode milik pegawai.
        rata_per_aspek = data_pegawai[ASPEK_7_COLS].mean().round(1)

        # Hitung rata-rata total dari seluruh aspek dan seluruh periode
        nilai_rata = round(rata_per_aspek.mean(), 1)

        # Tentukan kategori berdasarkan rata-rata keseluruhan
        status_nilai = kategori_7_aspek(nilai_rata)

        # Buat label periode, contoh: 2026-01 s.d. 2026-04
        periode_unik = sorted(data_pegawai["Periode"].astype(str).unique().tolist())

        if len(periode_unik) == 1:
            label_periode = periode_unik[0]
        else:
            label_periode = f"{periode_unik[0]} s.d. {periode_unik[-1]}"

        # Dataframe untuk grafik rata-rata keseluruhan per aspek
        df_one = pd.DataFrame({
            "Aspek": [a.replace("_", " ") for a in ASPEK_7_COLS],
            "Skor": rata_per_aspek.values
        })

        judul_grafik = f"Rata-rata Keseluruhan 7 Aspek ASN - {nama}"

    # KPI card detail pegawai
    p1, p2, p3 = st.columns(3)
    kpi_card(p1, "Rata-rata 7 Aspek", nilai_rata)
    kpi_card(p2, "Status", status_nilai)
    kpi_card(p3, "Periode", label_periode)

    # Urutkan dari skor terbesar ke terkecil
    df_one = df_one.sort_values("Skor", ascending=False)

    # Grafik batang skor per aspek
    fig = px.bar(
        df_one,
        x="Aspek",
        y="Skor",
        color="Skor",
        color_continuous_scale="Blues",
        title=judul_grafik
    )
    fig.update_xaxes(categoryorder="array", categoryarray=df_one["Aspek"].tolist())

    # Batas skor dibuat 0-100 agar konsisten
    fig.update_yaxes(range=[0, 100])
    fig.update_layout(xaxis_tickangle=-30, height=520)

    st.plotly_chart(fig, use_container_width=True)

    # Tabel detail skor aspek
    st.dataframe(df_one, use_container_width=True, hide_index=True)

    # Jika mode rata-rata keseluruhan dipilih,
    # tampilkan juga riwayat nilai bulanan pegawai.
    if mode_detail == "Rata-rata Keseluruhan":
        st.markdown("#### Riwayat Nilai Bulanan Pegawai")

        riwayat_cols = [
            "Periode",
            *ASPEK_7_COLS,
            "Rata_7_Aspek",
            "Kategori_7_Aspek",
        ]

        riwayat_cols = [c for c in riwayat_cols if c in data_pegawai.columns]

        st.dataframe(
            data_pegawai[riwayat_cols].sort_values("Periode"),
            use_container_width=True,
            hide_index=True
        )

    st.markdown("---")
    st.subheader("Rekap KPI 7 Aspek Seluruh Pegawai")

    cari = st.text_input("🔍 Cari nama", key="cari_7aspek_rekap")

    show = kpi.copy()

    if cari:
        show = show[
            show["NAMA"].astype(str).str.lower().str.contains(cari.lower(), na=False)
        ]

    cols_show = [
        "NIP",
        "NAMA",
        "Periode",
        *ASPEK_7_COLS,
        "Rata_7_Aspek",
        "Kategori_7_Aspek",
        "Waktu_Update",
    ]

    cols_show = [c for c in cols_show if c in show.columns]

    st.dataframe(
        show[cols_show],
        use_container_width=True,
        height=420
    )

    csv_kpi = show[cols_show].to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        "⬇️ Download Rekap KPI 7 Aspek",
        csv_kpi,
        file_name="rekap_kpi_7aspek.csv",
        mime="text/csv",
        on_click=lambda: catat_aktivitas("Download Rekap KPI 7 Aspek")
    )


# =====================================================================
# =====================================================================
# MENU 7: RAPORT PEGAWAI
# =====================================================================
# =====================================================================
def menu_raport_pegawai(df_pegawai, kpi_jam_real, kpi_7aspek):
    """
    Menu Raport Pegawai.

    Menu ini menampilkan:
    1. Biodata pegawai dari Menu Informasi Pegawai
    2. Ringkasan KPI Kehadiran/Jam Kerja
    3. Ringkasan KPI 7 Aspek BerAKHLAK
    4. Tombol cetak/download raport dalam bentuk PDF
    """

    page_header(
        "Raport Pegawai",
        "Biodata pegawai dan rangkuman KPI Kehadiran serta KPI 7 Aspek BerAKHLAK"
    )

    if df_pegawai.empty:
        st.warning("Data pegawai belum tersedia.")
        return

    # Membuat daftar nama pegawai dari database pegawai.
    # Selectbox Streamlit bisa diketik langsung
    daftar_nama = sorted(
        df_pegawai["NAMA"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    # Saat diklik, user bisa langsung mengetik nama,
    # lalu Streamlit akan menampilkan nama-nama yang mirip.
    nama = st.selectbox(
        "🔍 Cari / pilih nama pegawai",
        daftar_nama,
        key="pilih_pegawai_raport"
    )

    biodata = ambil_biodata_pegawai(df_pegawai, nama)
    kpi_jam = ambil_kpi_jam_pegawai(kpi_jam_real, nama)
    kpi_aspek = ambil_kpi_7aspek_pegawai(kpi_7aspek, nama)

    st.markdown("---")
    st.subheader("Biodata Pegawai")

    biodata_view = pd.DataFrame([
        ["Nama", biodata.get("NAMA", "-")],
        ["NIP", biodata.get("NIP BARU", "-")],
        ["NIK", biodata.get("NIK", "-")],
        ["Jenis Kelamin", biodata.get("JK", "-")],
        ["Agama", biodata.get("AGAMA", "-")],
        ["Status Pegawai", biodata.get("STATUS PNS/CPNS", "-")],
        ["Jenjang", biodata.get("JENJANG", "-")],
        ["Jabatan", biodata.get("NAMA JABATAN", "-")],
        ["Unit Kerja", biodata.get("UNIT KERJA", "-")],
        ["Usia", biodata.get("USIA", "-")],
        ["TMT Masuk Kerja", biodata.get("TMT MASUK KERJA", "-")],
        ["TMT Pensiun", biodata.get("TMT PENSIUN", "-")],
    ], columns=["Keterangan", "Isi"])

    st.dataframe(biodata_view, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Rangkuman KPI Kehadiran / Jam Kerja")

    kpi_jam_view = pd.DataFrame([
        ["Skor Jam Kerja", kpi_jam.get("Skor_Jam_Kerja", "-")],
        ["Status Jam Kerja", kpi_jam.get("Kategori_Jam_Kerja", "-")],
        ["Jumlah Hari Kerja", kpi_jam.get("Hari_Data", "-")],
        ["Total Potongan", f"Rp {int(kpi_jam.get('Potongan_Total', 0)):,}".replace(",", ".") if kpi_jam else "-"],
    ], columns=["Indikator", "Nilai"])

    st.dataframe(kpi_jam_view, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Rangkuman KPI 7 Aspek BerAKHLAK")

    kpi_aspek_view = pd.DataFrame([
        ["Periode", kpi_aspek.get("Periode", "-")],
        ["Orientasi Pelayanan", kpi_aspek.get("Orientasi_Pelayanan", "-")],
        ["Akuntabel", kpi_aspek.get("Akuntabel", "-")],
        ["Kompeten", kpi_aspek.get("Kompeten", "-")],
        ["Harmonis", kpi_aspek.get("Harmonis", "-")],
        ["Loyal", kpi_aspek.get("Loyal", "-")],
        ["Adaptif", kpi_aspek.get("Adaptif", "-")],
        ["Kolaboratif", kpi_aspek.get("Kolaboratif", "-")],
        ["Rata-rata 7 Aspek", kpi_aspek.get("Rata_7_Aspek", "-")],
        ["Status 7 Aspek", kpi_aspek.get("Kategori_7_Aspek", "-")],
    ], columns=["Indikator", "Nilai"])

    st.dataframe(kpi_aspek_view, use_container_width=True, hide_index=True)

    st.markdown("---")

    pdf_buffer = buat_pdf_raport_pegawai(
        biodata=biodata,
        kpi_jam=kpi_jam,
        kpi_7aspek=kpi_aspek
    )

    # Membersihkan karakter yang kurang aman untuk nama file
    nama_file_bersih = re.sub(r'[\\/*?:"<>|]', "", nama)

    # Format nama file PDF raport
    nama_file = f"RAPORT {nama_file_bersih}.pdf"

    st.download_button(
        "⬇️ Cetak / Download Raport Pegawai PDF",
        data=pdf_buffer,
        file_name=nama_file,
        mime="application/pdf",
        use_container_width=True,
        on_click=lambda: catat_aktivitas("Download Raport Pegawai")
    )


# =====================================================================
# =====================================================================
# MENU 8 : ARSIP PEGAWAI
# ---------------------------------------------------------------------
# Menampilkan seluruh pegawai yang:
# - pensiun
# - pindah biro
# - resign
# - non aktif
# =====================================================================
# =====================================================================
def menu_arsip_pegawai(df_arsip):
    """
    Menu khusus arsip pegawai.
    """

    page_header(
        "ARSIP PEGAWAI",
        "Data pegawai pensiun, pindah biro, resign, dan non aktif"
    )

    if df_arsip.empty:
        st.warning("Belum ada data arsip pegawai.")
        return

    st.info(
        "Menu ini menampilkan seluruh pegawai yang sudah tidak aktif "
        "di lingkungan Biro Umum."
    )

    # =========================================================
    # FILTER DATA
    # =========================================================
    col1, col2, col3 = st.columns(3)

    with col1:
        keyword = st.text_input(
            "🔍 Cari Nama / NIP",
            placeholder="Ketik nama atau NIP pegawai"
        )

    with col2:
        status_opsi = ["Semua"]

        if "STATUS ARSIP" in df_arsip.columns:
            status_opsi += sorted(
                df_arsip["STATUS ARSIP"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

        filter_status = st.selectbox(
            "Filter Status Arsip",
            status_opsi
        )

    with col3:
        unit_opsi = ["Semua"]

        if "UNIT KERJA" in df_arsip.columns:
            unit_opsi += sorted(
                df_arsip["UNIT KERJA"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

        filter_unit = st.selectbox(
            "Filter Unit Kerja",
            unit_opsi
        )

    # =========================================================
    # FILTER DATAFRAME
    # =========================================================
    data = df_arsip.copy()

    if keyword:
        kw = keyword.lower()

        mask = (
            data["NAMA"].astype(str).str.lower().str.contains(kw, na=False)
            | data["NIP BARU"].astype(str).str.lower().str.contains(kw, na=False)
        )

        data = data[mask]

    if filter_status != "Semua":
        data = data[data["STATUS ARSIP"] == filter_status]

    if filter_unit != "Semua":
        data = data[data["UNIT KERJA"] == filter_unit]

   
    #----------KPI RINGKAS---------#
    st.markdown("---")

    total_arsip = len(data)

    total_pensiun = (
        (data["STATUS ARSIP"] == "PENSIUN").sum()
        if "STATUS ARSIP" in data.columns else 0
    )

    total_nonaktif = (
        data["STATUS ARSIP"]
        .isin(["DIHAPUS", "PINDAH BIRO"])
        .sum()
        if "STATUS ARSIP" in data.columns else 0
    )
    c1, c2, c3 = st.columns(3)

    kpi_card(c1, "Total Arsip", total_arsip)
    kpi_card(c2, "Pegawai Pensiun", total_pensiun)
    kpi_card(c3, "Pindah / Dihapus", total_nonaktif)

    #---------TABEL ARSIP---------#
    st.markdown("---")
    st.subheader("Tabel Arsip Pegawai")

    kolom_tampil = [
        "NO",
        "NAMA",
        "NIP BARU",
        "JK",
        "STATUS PNS/CPNS",
        "NAMA JABATAN",
        "UNIT KERJA",
        "TMT MASUK KERJA",
        "TMT PENSIUN",
        "STATUS ARSIP",
        "TANGGAL ARSIP",
        "KETERANGAN",
    ]

    kolom_tampil = [c for c in kolom_tampil if c in data.columns]

    st.dataframe(
        data[kolom_tampil],
        use_container_width=True,
        height=500
    )

    # =========================================================
    # DOWNLOAD EXCEL / CSV
    # =========================================================
    csv = data.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        "⬇️ Download Arsip Pegawai",
        csv,
        file_name="arsip_pegawai.csv",
        mime="text/csv",
        use_container_width=True,
        on_click=lambda: catat_aktivitas("Download Arsip Pegawai")
    )


# =====================================================================
# =====================================================================
# MENU: HISTORY ACTIVITY
# =====================================================================
# =====================================================================
def menu_history_aktivitas():
    page_header(
        "History Activity",
        "Log kegiatan seluruh pengguna pada sistem selama 1 bulan terakhir"
    )

    log = load_log_aktivitas()

    if log.empty:
        st.info("Belum ada kegiatan yang tercatat.")
        return

    # Hanya tampilkan 1 bulan (30 hari) terakhir
    batas_awal = pd.Timestamp.today().normalize() - pd.Timedelta(days=30)
    log = log[log["WAKTU_DATE"].isna() | (log["WAKTU_DATE"] >= batas_awal)]

    # Urutkan dari kegiatan paling baru
    log = log.sort_values("WAKTU_DATE", ascending=False, na_position="last")

    # ---------- Info pengguna terakhir mengunjungi website ----------
    if not log.empty:
        aktivitas_terakhir = log.iloc[0]
        st.info(
            f"Pengunjung terakhir: **{aktivitas_terakhir['NAMA']}** "
            f"{aktivitas_terakhir['KEGIATAN']} pada {aktivitas_terakhir['WAKTU']}"
        )

    # ---------- Filter pencarian ----------
    col1, col2 = st.columns([2, 1])
    with col1:
        keyword = st.text_input(
            "🔍 Cari User / Nama / Kegiatan",
            placeholder="Ketik kata kunci..."
        )
    with col2:
        opsi_user = ["Semua"] + sorted(log["NAMA"].dropna().unique().tolist())
        f_user = st.selectbox("Filter User", opsi_user)

    if keyword:
        kw = keyword.lower()
        mask = (
            log["USERNAME"].str.lower().str.contains(kw, na=False)
            | log["NAMA"].str.lower().str.contains(kw, na=False)
            | log["KEGIATAN"].str.lower().str.contains(kw, na=False)
        )
        log = log[mask]

    if f_user != "Semua":
        log = log[log["NAMA"] == f_user]

    st.info(f"Menampilkan **{len(log)}** kegiatan dalam 30 hari terakhir.")

    # ---------- Tabel History Activity ----------
    tabel_log = log[["WAKTU", "USERNAME", "KEGIATAN"]].rename(columns={
        "WAKTU": "Waktu",
        "USERNAME": "User",
        "NAMA": "Nama",
        "KEGIATAN": "Kegiatan",
    }).reset_index(drop=True)
    tabel_log.insert(0, "No", range(1, len(tabel_log) + 1))

    st.dataframe(
        tabel_log,
        use_container_width=True,
        height=500,
        hide_index=True
    )

    # ---------- Download ----------
    csv_log = tabel_log.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Download History Activity",
        csv_log,
        file_name=f"history_activity_{date.today().strftime('%Y-%m-%d')}.csv",
        mime="text/csv",
        use_container_width=True,
        on_click=lambda: catat_aktivitas("Download History Activity")
    )


# =====================================================================
# =====================================================================
# HALAMAN 3: LOGOUT
# =====================================================================
# =====================================================================
def page_logout():
    """Tampilan konfirmasi logout."""
    page_header("Logout", "Anda akan keluar dari sistem")
    st.warning(f"Halo **{st.session_state.get('nama','User')}**, "
               f"Anda login sejak {st.session_state.get('login_time','-')}.")
    st.write("Apakah Anda yakin ingin keluar dari aplikasi?")

    c1, c2, _ = st.columns([1,1,3])
    with c1:
        if st.button("Ya, Logout"):
            catat_aktivitas("Keluar dari sistem")
            # Hapus seluruh session_state untuk reset
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.success("Anda telah berhasil logout. Sampai jumpa! 👋")
            st.rerun()
    with c2:
        if st.button("Batal"):
            # Kembali ke menu utama
            st.session_state["menu"] = "Informasi Pegawai"
            st.rerun()


# =====================================================================
# =====================================================================
# MAIN APP - mengatur navigasi seluruh halaman
# =====================================================================
# =====================================================================
def main():
    # Inisialisasi state login bila belum ada
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if "nama" not in st.session_state:
        st.session_state["nama"] = "User"

    if "role" not in st.session_state:
        st.session_state["role"] = "-"

    if "login_time" not in st.session_state:
        st.session_state["login_time"] = "-"

    # ====== JIKA BELUM LOGIN -> tampilkan halaman login ======
    if not st.session_state["logged_in"]:
        page_login()
        return

    # ====== Setelah login: muat data ======
    # DATA_PEGAWAI.csv digunakan sebagai data awal.
    # Setelah ada perubahan melalui Menu Kelola Pegawai,
    # aplikasi akan memakai pegawai_kelola.csv sebagai database aktif.
    df_seed = load_data(DATA_MASTER_PATH)
    df = load_data_pegawai_kelola(df_seed, PEGAWAI_MASTER_PATH)

    # Menu KPI 7 Aspek real dari hasil upload CSV/XLSX
    kpi_7aspek = load_7aspek_final(ASPEK_7_PATH)
    
    # Menu KPI Jam Kerja real dari database operasional kehadiran
    df_kehadiran = load_kehadiran_final()
    kpi_jam_real = (
        ubah_kehadiran_ke_kpi(df_kehadiran)
        if not df_kehadiran.empty
        else pd.DataFrame()
    )

    df = arsipkan_pegawai_pensiun_otomatis(df)
    # Simpan ulang data aktif setelah pegawai pensiun dipindahkan
    simpan_data_pegawai_kelola(df)
    
    # Load data arsip pegawai
    df_arsip = load_arsip_pegawai()

    # ====== Sidebar: identitas user + menu navigasi ======
    with st.sidebar:
        st.markdown("""
            <div style="
                padding-bottom:18px;
                margin-bottom:20px;
                border-bottom:1px solid rgba(255,255,255,0.14);
            ">
                <div style="
                    font-size:16px;
                    font-weight:800;
                    letter-spacing:0.8px;
                    color:white;
                    line-height:1.35;
                    text-transform:uppercase;
                ">
                    Kepegawaian Biro Umum
                </div>
                <div style="
                    font-size:11px;
                    color:rgba(255,255,255,0.62);
                    margin-top:5px;
                    letter-spacing:0.3px;
                ">
                    Sistem Informasi Aparatur
                </div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div style="
            background:rgba(255,255,255,0.10);
            border:1px solid rgba(255,255,255,0.13);
            padding:14px;
            border-radius:16px;
            margin-bottom:18px;
            box-shadow:0 10px 24px rgba(0,0,0,0.14);
        ">
            <div style="
                font-size:12px;
                color:rgba(255,255,255,0.70);
                margin-bottom:6px;
            ">
                Welcome,
            </div>
            <div style="
                font-size:16px;
                font-weight:800;
                color:white;
            ">
                {st.session_state.get('nama', 'User')}
            </div>
            <div style="
                font-size:12px;
                color:rgba(255,255,255,0.72);
                margin-top:2px;
            ">
                {st.session_state.get('role', '-')}
            </div>
            <div style="
                font-size:10px;
                color:rgba(255,255,255,0.55);
                margin-top:8px;
            ">
                Login: {st.session_state.get('login_time', '-')}
            </div>
        </div>
    """, unsafe_allow_html=True)

        menu = st.radio(
            "Pilih menu:",
            [
                "INFORMASI PEGAWAI",
                "DASHBOARD INFORMASI PEGAWAI",
                "KELOLA PEGAWAI",
                "KPI JAM KERJA",
                "KPI 7 ASPEK ASN",
                "RAPORT PEGAWAI",
                "ARSIP PEGAWAI",
                "HISTORY ACTIVITY",
                "LOGOUT",
            ],
            label_visibility="collapsed"
        )

        st.markdown("""
            <div style="
                font-size:11px;
                color:rgba(255,255,255,0.58);
                line-height:1.5;
                padding:4px 4px 0 4px;
            ">
                © 2026 Sekretariat Daerah<br>
                Dashboard Kepegawaian v1.0
            </div>
        """, unsafe_allow_html=True)

    # ====== Routing menu ======
    if menu == "INFORMASI PEGAWAI":
        menu_informasi_pegawai(df)

    elif menu == "DASHBOARD INFORMASI PEGAWAI":
        menu_dashboard(df)

    elif menu == "KELOLA PEGAWAI":
        menu_kelola_pegawai(df, df_seed)

    elif menu == "KPI JAM KERJA":
        menu_kpi_jam_kerja(kpi_jam_real)

    elif menu == "KPI 7 ASPEK ASN":
        menu_kpi_7aspek(kpi_7aspek, df)

    elif menu == "RAPORT PEGAWAI":
        menu_raport_pegawai(df, kpi_jam_real, kpi_7aspek)

    elif menu == "ARSIP PEGAWAI":
        menu_arsip_pegawai(df_arsip)

    elif menu == "HISTORY ACTIVITY":
        menu_history_aktivitas()

    elif menu == "LOGOUT":
        page_logout()

# Titik masuk program
if __name__ == "__main__":
    main()