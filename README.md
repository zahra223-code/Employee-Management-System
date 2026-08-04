# Dashboard Kepegawaian, Biro Umum Sekretariat Daerah Provinsi Jawa Timur

Aplikasi web internal berbasis Streamlit untuk pengelolaan data kepegawaian, pemantauan kinerja pegawai (KPI Jam Kerja dan KPI 7 Aspek ASN/BerAKHLAK), serta penyusunan laporan dan raport pegawai secara otomatis. Aplikasi ini dikembangkan untuk menggantikan proses rekapitulasi manual berbasis spreadsheet yang selama ini tersebar di berbagai berkas terpisah, dengan menyatukan seluruh alur kerja, mulai dari pencarian data pegawai, pemantauan kedisiplinan jam kerja, penilaian kinerja tujuh aspek ASN, hingga pengarsipan pegawai purna tugas, ke dalam satu dashboard terpusat.

---

## Daftar Isi

- [Ringkasan Proyek](#ringkasan-proyek)
- [Fitur Utama](#fitur-utama)
- [Arsitektur dan Alur Data](#arsitektur-dan-alur-data)
- [Struktur Direktori Proyek](#struktur-direktori-proyek)
- [Teknologi yang Digunakan](#teknologi-yang-digunakan)
- [Instalasi dan Menjalankan Aplikasi](#instalasi-dan-menjalankan-aplikasi)
- [Spesifikasi Format Data](#spesifikasi-format-data)
- [Autentikasi Pengguna](#autentikasi-pengguna)
- [Peta Menu Aplikasi](#peta-menu-aplikasi)
- [Catatan Teknis dan Keputusan Desain](#catatan-teknis-dan-keputusan-desain)
- [Batasan dan Risiko yang Diketahui](#batasan-dan-risiko-yang-diketahui)
- [Roadmap Pengembangan](#roadmap-pengembangan)
- [Lisensi dan Kontak](#lisensi-dan-kontak)

---

## Ringkasan Proyek

Dashboard Kepegawaian merupakan aplikasi *single-page* berbasis Streamlit yang berfungsi sebagai sistem informasi kepegawaian skala internal, tanpa bergantung pada infrastruktur basis data relasional. Seluruh operasi baca dan tulis data dilakukan melalui berkas CSV/XLSX yang dikelola oleh `pandas`, sehingga aplikasi dapat dijalankan pada lingkungan komputasi terbatas tanpa proses instalasi server basis data terpisah. Pendekatan ini merupakan trade-off yang disengaja: kemudahan deployment dan portabilitas diprioritaskan di atas skalabilitas multi-pengguna simultan, pertimbangan ini dijelaskan lebih lanjut pada bagian [Batasan dan Risiko yang Diketahui](#batasan-dan-risiko-yang-diketahui).

---

## Fitur Utama

### Manajemen Data Pegawai

- Pencarian pegawai secara *real-time* berdasarkan nama, NIP, atau NIK.
- Penambahan, perubahan, dan penghapusan (pengarsipan) data pegawai melalui antarmuka form, tanpa memerlukan penyuntingan berkas CSV/Excel secara manual.
- Deteksi dan pengarsipan pensiun otomatis: pegawai yang telah melewati Tanggal Mulai Tugas (TMT) pensiun secara otomatis dipindahkan ke Arsip Pegawai setiap kali aplikasi dimuat, sehingga data aktif senantiasa konsisten dengan status kepegawaian yang sebenarnya.
- Normalisasi otomatis terhadap variasi penamaan kolom pada berkas master (misalnya `GOL. RUANG` versus `GOLONGAN`), guna menjaga konsistensi saat data diimpor dari sumber yang berbeda-beda.

### Kelola Pegawai

- Formulir input terstruktur untuk penambahan data pegawai baru, dengan validasi kolom wajib sebelum data disimpan ke `pegawai_kelola.csv`.
- Mekanisme pembaruan data pegawai (biodata, jabatan, unit kerja, golongan, dan atribut kepegawaian lainnya) tanpa mengubah baris data pegawai lain.
- Fungsi pengarsipan manual, memungkinkan pengguna memindahkan pegawai ke Arsip Pegawai secara langsung (di luar mekanisme deteksi pensiun otomatis), untuk kasus seperti mutasi, resign, atau pemberhentian.
- Setiap perubahan data melalui menu ini tercatat pada Log Aktivitas Pengguna, sehingga jejak perubahan data kepegawaian tetap dapat ditelusuri.

### Dashboard Informasi Pegawai

- Visualisasi interaktif menggunakan Plotly untuk komposisi usia, jenis kelamin, jenjang pendidikan, agama, status kepegawaian, serta sebaran unit kerja dan jabatan.
- Ringkasan kuantitatif jumlah pegawai aktif yang diperbarui secara real-time terhadap seluruh filter yang diterapkan pengguna.

### KPI Jam Kerja

- Unggah data kehadiran mentah (harian) dalam format CSV atau XLSX, yang diproses secara otomatis menjadi skor kedisiplinan per pegawai.
- Perhitungan otomatis untuk sembilan komponen pelanggaran: Ijin, Terlambat, Pulang Cepat, Tidak Absen Datang, Tidak Absen Pulang, Alpha, Tidak Ikut Senam, Terlambat Senam, dan Tidak Ikut Apel, masing-masing dengan tarif potongan tersendiri.
- Filter periode fleksibel: seluruh data, bulan tertentu, atau rentang tanggal kustom.
- Grafik distribusi status kinerja, papan peringkat lima belas pegawai dengan potongan tertinggi, serta rincian data per pegawai.
- Penyaringan otomatis terhadap status keaktifan pegawai, sehingga pegawai yang telah pensiun, dihapus, atau diarsipkan tidak memengaruhi statistik.
- Ekspor rekap ke format CSV.

### KPI 7 Aspek ASN (BerAKHLAK)

- Unggah penilaian tujuh aspek (Orientasi Pelayanan, Akuntabel, Kompeten, Harmonis, Loyal, Adaptif, Kolaboratif) per periode, dengan mekanisme *upsert*: data dengan kombinasi NIP, Nama, dan Periode yang identik akan diperbarui, bukan diduplikasi.
- Ringkasan kartu KPI (Tercapai, Dibawah Ekspektasi, Rata-rata Skor) dihitung dari satu baris data terbaru per pegawai, sehingga terhindar dari bias perhitungan ketika satu pegawai memiliki data dari beberapa periode sekaligus.
- Radar chart rata-rata skor per aspek dan pie chart distribusi status, yang konsisten secara numerik dengan angka pada kartu ringkasan.
- Tampilan detail per pegawai dalam mode "Per Bulan" maupun "Rata-rata Keseluruhan".
- Penyaringan otomatis terhadap status keaktifan pegawai, disertai pesan peringatan eksplisit apabila data hasil unggahan tidak cocok dengan data pegawai (misalnya akibat kesalahan penulisan nama).

### Raport Pegawai

- Menggabungkan biodata, ringkasan KPI Jam Kerja, dan ringkasan KPI 7 Aspek dalam satu tampilan terpadu per pegawai.
- Ekspor raport individual ke format PDF siap cetak, dibangun menggunakan ReportLab.

### Arsip Pegawai

- Menyimpan riwayat pegawai yang telah pensiun, pindah tugas, mengundurkan diri, atau dinonaktifkan, terpisah dari data aktif namun tetap dapat ditelusuri.
- Ekspor data arsip ke format CSV.

### Log Aktivitas Pengguna

- Mencatat jejak aktivitas pengguna selama satu sesi berjalan (login, unggah data, unduh laporan, perubahan data pegawai, logout, dan sebagainya).
- Ekspor log aktivitas ke format CSV.

### Autentikasi dan Konsistensi Antarmuka

- Sistem login berbasis sesi (`st.session_state`) dengan halaman login terpisah.
- Desain antarmuka yang konsisten dengan identitas visual instansi pemerintahan (palet warna biru, tipografi Poppins), termasuk sidebar navigasi yang telah dioptimalkan agar keterbacaan teks terjaga di seluruh jenis komponen input.

---

## Arsitektur dan Alur Data

Aplikasi ini tidak menggunakan sistem manajemen basis data relasional. Seluruh data operasional disimpan sebagai berkas CSV di dalam struktur direktori lokal, dan dibaca-tulis ulang oleh aplikasi melalui `pandas`. Keputusan arsitektural ini didasari pertimbangan bahwa aplikasi perlu dijalankan secara internal tanpa memerlukan instalasi server basis data terpisah, dengan konsekuensi logis berupa keterbatasan pada skenario akses konkuren (lihat [Batasan dan Risiko yang Diketahui](#batasan-dan-risiko-yang-diketahui)).

Alur data secara umum dapat digambarkan sebagai berikut:

```
Data_Master.xlsx (data awal)
        |
        v
pegawai_kelola.csv  <──────────────┐
        |                          │  (perubahan tersimpan otomatis)
        v                          │
  [ Aplikasi Streamlit ] ──────────┘
        |
        ├── kehadiran_final.csv      → diolah menjadi KPI Jam Kerja
        ├── kpi_7aspek_final.csv     → diolah menjadi KPI 7 Aspek ASN
        └── arsip_pegawai.csv        ← pegawai pensiun/nonaktif dipindahkan otomatis
```

Setiap kali aplikasi dimuat, data pegawai aktif disaring ulang terhadap TMT pensiun, dan pegawai yang telah melewati masa tugas dipindahkan secara otomatis ke arsip. Mekanisme ini memastikan bahwa seluruh perhitungan statistik dan KPI pada dashboard senantiasa merepresentasikan populasi pegawai yang benar-benar masih aktif, bukan sekadar data historis yang belum dibersihkan.

---

## Struktur Direktori Proyek

Aplikasi mengasumsikan struktur direktori berikut relatif terhadap root proyek (lihat konfigurasi `BASE_DIR` pada kode sumber):

```
project-root/
├── 01_source_code/
│   └── dashboard_kepegawaian.py        # Entry point aplikasi
├── 03_dataset/
│   └── employee/
│       └── Data_Master.xlsx            # Data pegawai awal (seed)
├── 04_database_operasional/
│   ├── employee/
│   │   ├── pegawai_kelola.csv          # Data pegawai aktif (live)
│   │   └── arsip_pegawai.csv           # Arsip pegawai nonaktif
│   ├── attendance/
│   │   └── kehadiran_final.csv         # Basis data kehadiran harian
│   ├── performance/
│   │   └── kpi_7aspek_final.csv        # Basis data penilaian 7 aspek ASN
│   └── activity/
│       └── log_aktivitas.csv           # Tidak lagi digunakan; log kini in-memory
└── 05_asset/
    └── LOGO.png                        # Logo instansi untuk halaman login/sidebar
```

Folder `04_database_operasional/` berisi data pegawai yang bersifat sensitif dan tidak boleh diunggah ke repositori publik. Gunakan berkas `.gitignore` untuk mengecualikan folder ini secara eksplisit (lihat bagian [Instalasi](#instalasi-dan-menjalankan-aplikasi)).

---

## Teknologi yang Digunakan

| Kebutuhan Fungsional | Pustaka/Framework |
|---|---|
| Framework web dan dashboard | [Streamlit](https://streamlit.io/) |
| Manipulasi data tabular | [pandas](https://pandas.pydata.org/), [numpy](https://numpy.org/) |
| Visualisasi interaktif | [Plotly](https://plotly.com/python/) (`plotly.express`, `plotly.graph_objects`) |
| Pembacaan dan penulisan berkas Excel | [openpyxl](https://openpyxl.readthedocs.io/) (dependensi engine untuk `pandas`) |
| Pembuatan dokumen PDF | [ReportLab](https://www.reportlab.com/) |
| Perhitungan tanggal presisi | [python-dateutil](https://dateutil.readthedocs.io/) |
| Manajemen path lintas platform | `pathlib` (modul bawaan Python) |

---

## Instalasi dan Menjalankan Aplikasi

### 1. Prasyarat Sistem

- Python versi 3.9 atau lebih baru.
- `pip` telah terpasang pada sistem.

### 2. Clone Repositori

```bash
git clone https://github.com/<username>/<nama-repo>.git
cd <nama-repo>
```

### 3. Pembuatan Virtual Environment

Penggunaan virtual environment sangat disarankan untuk mengisolasi dependensi proyek dari instalasi Python sistem.

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

### 4. Instalasi Dependensi

Buat berkas `requirements.txt` (apabila belum tersedia) dengan isi sebagai berikut:

```
streamlit
pandas
numpy
plotly
reportlab
python-dateutil
openpyxl
```

Kemudian jalankan:

```bash
pip install -r requirements.txt
```

### 5. Persiapan Struktur Direktori Data

Pastikan struktur direktori telah sesuai dengan penjelasan pada bagian [Struktur Direktori Proyek](#struktur-direktori-proyek), dengan minimal dua berkas berikut telah tersedia:

- `03_dataset/employee/Data_Master.xlsx`, data pegawai awal.
- `05_asset/LOGO.png`, logo instansi.

Berkas-berkas pada `04_database_operasional/` akan dibuat dan diisi secara otomatis oleh aplikasi seiring penggunaan (unggah data, penambahan/perubahan data, dan sebagainya).

### 6. Menjalankan Aplikasi

```bash
streamlit run 01_source_code/dashboard_kepegawaian.py
```

Aplikasi akan terbuka secara otomatis pada `http://localhost:8501`.

---

## Spesifikasi Format Data

### Data Master Pegawai (`Data_Master.xlsx`)

Kolom minimal yang diharapkan (variasi penamaan kolom dinormalisasi secara otomatis oleh aplikasi):

`NO`, `NAMA`, `NIP BARU`, `NIK`, `JK`, `AGAMA`, `STATUS PNS/CPNS`, `JENJANG`, `NAMA JABATAN`, `UNIT KERJA`, `TGL LAHIR`, `USIA`, `TMT MASUK KERJA`, `TMT PANGKAT`, `GOLONGAN`, `TMT PENSIUN`, `TAHUN PENSIUN`.

### Data Kehadiran (diunggah melalui menu KPI Jam Kerja)

Format CSV atau XLSX berisi data kehadiran harian per pegawai, mencakup kolom seperti `nama`, `nip`, `opd`, `tanggal`, `kode_asli`, `menit_terlambat`, dan `potong_gaji`. Kode kehadiran dipetakan secara otomatis ke sembilan komponen pelanggaran (Ijin, Terlambat, Pulang Cepat, dan seterusnya).

### Data Penilaian 7 Aspek ASN (diunggah melalui menu KPI 7 Aspek ASN)

Kolom wajib: `NAMA`, `Orientasi_Pelayanan`, `Akuntabel`, `Kompeten`, `Harmonis`, `Loyal`, `Adaptif`, `Kolaboratif`.

Kolom opsional: `NIP`, `Periode` (format `YYYY-MM`, contoh: `2026-01`).

---

## Autentikasi Pengguna

Aplikasi menggunakan sistem login berbasis dictionary Python (`USERS`) yang tervalidasi melalui `st.session_state`, sesuai untuk lingkup pengguna internal terbatas dengan kebutuhan keamanan minimal.

**Catatan penting untuk lingkungan produksi.** Implementasi saat ini menyimpan kredensial sebagai teks biasa (*plain text*) di dalam kode sumber, pendekatan ini memadai untuk tahap pengembangan atau demonstrasi, tetapi tidak sesuai untuk deployment produksi tanpa modifikasi. Rekomendasi perbaikan meliputi:

- Memindahkan kredensial ke *environment variable* atau *secrets manager* (misalnya `st.secrets` pada Streamlit Cloud).
- Menerapkan mekanisme *hashing* password (misalnya menggunakan `bcrypt` atau `passlib`).
- Memastikan berkas kode yang memuat kredensial asli tidak disertakan dalam repositori publik.

---

## Peta Menu Aplikasi

| Menu | Fungsi |
|---|---|
| Informasi Pegawai | Pencarian dan tampilan detail data pegawai |
| Dashboard Informasi Pegawai | Visualisasi komposisi demografis dan organisasi pegawai |
| Kelola Pegawai | Input, pembaruan, dan penghapusan (pengarsipan) data pegawai |
| KPI Jam Kerja | Unggah kehadiran, perhitungan skor kedisiplinan, dan rekapitulasi |
| KPI 7 Aspek ASN | Unggah penilaian 7 aspek BerAKHLAK dan rekapitulasi kinerja |
| Raport Pegawai | Ringkasan biodata dan KPI per pegawai, ekspor PDF |
| Arsip Pegawai | Riwayat pegawai nonaktif, pensiun, atau pindah tugas |
| Log Aktivitas | Log aktivitas pengguna selama sesi berjalan |
| Logout | Konfirmasi dan keluar dari sesi aplikasi |

---

## Catatan Teknis dan Keputusan Desain

- **Konsistensi statistik KPI.** Seluruh kartu ringkasan, pie chart, dan radar chart pada menu KPI dihitung dari satu sumber baris data yang identik per pegawai, bukan dari seluruh baris mentah, sehingga tidak terjadi bias perhitungan ganda ketika satu pegawai memiliki lebih dari satu periode data dalam mode "Semua Data".
- **Sinkronisasi status keaktifan.** Data KPI, baik Jam Kerja maupun 7 Aspek, senantiasa disaring terhadap daftar pegawai yang benar-benar masih aktif pada basis data pegawai, bukan hanya berdasarkan data yang diunggah. Mekanisme ini mencegah pegawai yang telah pensiun atau dinonaktifkan ikut memengaruhi statistik kinerja.
- **Diferensiasi pesan kesalahan.** Aplikasi membedakan secara eksplisit antara kondisi "basis data belum pernah diisi" dengan "data tersedia namun tersaring habis karena tidak cocok dengan data pegawai aktif" (misalnya akibat kesalahan penulisan nama), sehingga proses debugging oleh pengguna maupun tim IT dapat berlangsung lebih cepat dan akurat.
- **Isolasi gaya tampilan (CSS).** Aturan tampilan sidebar dan area konten utama dipisahkan secara eksplisit untuk mencegah komponen input baru yang ditambahkan di kemudian hari mewarisi gaya yang tidak sesuai konteksnya (misalnya teks gelap di atas latar belakang gelap).

---

## Batasan dan Risiko yang Diketahui

Bagian ini disertakan secara sengaja agar pengguna dan pengembang lanjutan memahami trade-off arsitektural yang melekat pada desain aplikasi saat ini, bukan sekadar daftar kekurangan.

- **Konkurensi akses.** Karena penyimpanan berbasis berkas CSV tanpa mekanisme locking eksplisit, penulisan data secara simultan oleh lebih dari satu pengguna berpotensi menyebabkan kondisi *race condition* atau kehilangan data. Aplikasi ini paling sesuai untuk skenario penggunaan sekuensial atau jumlah pengguna aktif yang kecil.
- **Skalabilitas volume data.** Pendekatan pembacaan penuh berkas CSV ke memori (`pandas.read_csv`) akan mengalami penurunan performa seiring pertumbuhan volume data kepegawaian dalam jangka panjang.
- **Keamanan kredensial.** Sebagaimana dijelaskan pada bagian [Autentikasi Pengguna](#autentikasi-pengguna), skema autentikasi saat ini belum memenuhi standar keamanan produksi.

---

## Roadmap Pengembangan

- Migrasi penyimpanan data dari CSV ke sistem basis data (PostgreSQL atau SQLite) untuk mendukung akses multi-pengguna secara konkuren.
- Penerapan hashing kredensial pengguna dan dukungan multi-role (admin, staf, pimpinan).
- Ekspor rekap KPI Jam Kerja dan 7 Aspek ke format Excel (`.xlsx`), sebagai pelengkap format CSV yang sudah tersedia.
- Notifikasi otomatis (in-app) untuk pegawai yang akan memasuki masa pensiun dalam periode mendatang.
- Penyusunan pengujian otomatis (unit test) untuk fungsi-fungsi pengolahan data KPI.

---

## Lisensi dan Kontak

Proyek ini dikembangkan untuk kebutuhan internal Biro Umum Sekretariat Daerah. Ketentuan lisensi pada bagian ini perlu disesuaikan dengan kebijakan instansi atau organisasi terkait (misalnya proprietary/internal use only, atau lisensi open-source seperti MIT apabila proyek akan dibagikan secara publik).

Untuk pertanyaan, laporan bug, atau permintaan fitur, silakan membuka Issue pada repositori ini atau menghubungi tim pengembang internal.
