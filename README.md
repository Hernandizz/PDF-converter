# HD Document & PDF Suite Pro 📄➡️✨

Aplikasi pengolah dan konverter dokumen lengkap (**All-in-One PDF Suite**) yang berjalan **100% lokal**, aman, tanpa batasan ukuran file, dan menjaga kualitas dokumen tetap **Ultra-HD (Lossless / 300+ DPI)**.

---

## 🌟 Fitur Utama

### 1. 📄 Word ke PDF (Ultra-HD & Lossless)
- ✅ **`doc.DoNotCompressImages = True`**: Gambar embedded di Word tidak dikompresi / di-downsample oleh Word.
- ✅ **Standard Cetak 300+ DPI (`wdExportOptimizeForPrint`)**: Kualitas profesional untuk percetakan dan pembesaran zoom.
- ✅ **Vektor Murni (`BitmapMissingFonts = False`)**: Tipografi tetap berupa kurva vektor tajam pada zoom 1000%+.
- ✅ **Bookmark Otomatis**: Menghasilkan daftar isi / bookmark PDF interaktif dari heading Word.
- ✅ **DOCX Image Inspector**: Memeriksa resolusi asli gambar sebelum konversi.

### 2. 📝 PDF ke Word (.docx / .doc)
- ✅ Menggunakan engine native **Word COM PDF Reflow**.
- ✅ Membuka dokumen PDF langsung dan menyimpannya kembali sebagai format dokumen Word yang dapat diedit (*editable*).
- ✅ Pilihan output `.docx` atau `.doc`.

### 3. 📑 Gabung PDF (PDF Merge)
- ✅ Menggabungkan beberapa berkas PDF menjadi satu berkas PDF utuh secara berurutan.
- ✅ Dilengkapi kontrol **geser urutan halaman (Naik ⬆️ / Turun ⬇️)** langsung dari antarmuka web.
- ✅ Bebas menentukan nama berkas PDF hasil penggabungan.

### 4. ✂️ Pisah PDF (PDF Split & Extract)
- ✅ **Ekstrak Rentang Halaman**: Pilih nomor/rentang halaman tertentu (misal: `1-3, 5, 8-10` atau `1, 4`) ke dalam satu berkas PDF baru.
- ✅ **Pisah Setiap Halaman**: Memisahkan setiap lembar dokumen PDF menjadi file terpisah dan otomatis dikemas dalam arsip **ZIP**.
- ✅ **Deteksi Halaman Otomatis**: Menampilkan jumlah halaman dan metadata dokumen secara instan saat file dipilih.

### 5. 🖼️ Gambar ke PDF HD (Image to PDF)
- ✅ Mengubah satu atau banyak foto / scan berkas (`.jpg`, `.jpeg`, `.png`, `.webp`, `.bmp`, `.tiff`) menjadi dokumen PDF berkualitas tinggi.
- ✅ Mode tata letak: **Ukuran Asli Gambar (Fit)** atau **Standar Kertas A4 (300 DPI)**.
- ✅ Konversi transparansi PNG yang bersih (tanpa artefak latar hitam).

---

## 🚀 Cara Menjalankan

### Cara 1: Menggunakan Antarmuka Web (Web GUI - Direkomendasikan)
Cukup jalankan file launcher Windows:
1. Double click file **`run.bat`** (atau ketik `.\run.bat` di terminal).
2. Browser akan terbuka otomatis di alamat: **`http://127.0.0.1:5000`**
3. Pilih tab fitur yang Anda butuhkan (Word→PDF, PDF→Word, Gabung, Pisah, atau Gambar→PDF).
4. Tarik & lepas (drag & drop) berkas Anda, lalu klik tombol proses.
5. Unduh hasilnya atau klik tombol **Buka Folder** untuk melihat langsung di File Explorer.

Atau jalankan melalui terminal:
```bash
python app.py
```

---

### Cara 2: Menggunakan CLI / Command Line

#### Konversi Dokumen Word ke PDF:
```bash
# Satu file
python converter.py "Dokumen.docx" -o "Hasil.pdf"

# Seluruh folder Word
python converter.py "./folder_word" -o "./folder_pdf"

# Inspeksi kualitas gambar DOCX tanpa konversi
python converter.py --inspect-only "Dokumen.docx"
```

#### Manipulasi PDF via Script Python:
```python
from pdf_tools import merge_pdfs, split_pdf, images_to_pdf

# 1. Gabung PDF
merge_pdfs(["dokumen1.pdf", "dokumen2.pdf"], "hasil_gabung.pdf")

# 2. Pisah / Ekstrak Halaman
split_pdf("dokumen.pdf", output_dir="./outputs", mode="range", range_str="1-3, 5")

# 3. Konversi Gambar ke PDF
images_to_pdf(["foto1.jpg", "foto2.png"], "album.pdf", page_size="fit")
```

---

## 📦 Instalasi Dependensi

Aplikasi membutuhkan **Python 3.8+** dan **Microsoft Word (Office)** di Windows (untuk fitur Word ↔ PDF).
Instal dependensi Python:
```bash
pip install -r requirements.txt
```

---

## 📂 Struktur Proyek

```
PDF converter/
│
├── converter.py              # Engine konversi Word ke PDF HD Lossless
├── pdf_to_word_converter.py  # Engine konversi PDF ke Word via COM Automation
├── pdf_tools.py              # Engine All-in-One (Merge, Split, Extract, Image-to-PDF)
├── app.py                    # Backend Web Server Flask (REST API)
├── run.bat                   # Launcher 1-klik untuk Windows
├── requirements.txt          # Dependensi Python
├── README.md                 # Dokumentasi lengkap
│
├── static/
│   ├── css/
│   │   └── style.css         # Desain profesional Glassmorphism & layout responsif
│   └── js/
│       └── app.js            # Controller frontend (Drag & Drop, Reorder, Progress Bar)
│
├── templates/
│   └── index.html            # Dashboard Web GUI (5 Tab Suite)
│
├── uploads/                  # Direktori penyimpanan sementara file upload
└── outputs/                  # Direktori hasil konversi & pengolahan dokumen
```

---

## 🔒 Privasi & Keamanan 100% Lokal
Seluruh pemrosesan dilakukan **100% secara offline di komputer lokal Anda**. Tidak ada data atau dokumen yang dikirim ke server luar atau cloud pihak ketiga.
