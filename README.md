# HD Word to PDF Converter Pro 📄➡️✨

Aplikasi konverter dokumen Microsoft Word (`.docx` / `.doc`) ke PDF dengan hasil **Ultra-HD**, **tidak pecah**, dan **kualitas gambar 100% terjaga (Lossless / 300+ DPI)**.

---

## 🔍 Mengapa Hasil Word ke PDF Biasa Sering Pecah / Buram?

Kebanyakan converter atau fitur default *Save As PDF* di Word menghasilkan gambar yang buram karena:
1. **Word Image Downsampling**: Word secara otomatis memperkecil (downsample) resolusi gambar menjadi 96–150 PPI untuk menghemat ukuran file.
2. **Preset Export Ukuran Layar**: Sering kali converter menggunakan opsi `wdExportOptimizeForOnScreen` (untuk web/layar) yang mengompresi gambar JPEG secara agresif.
3. **Rasterisasi Teks/Font**: Beberapa tool merasterisasi font menjadi gambar bitmap saat disimpan ke PDF sehingga teks menjadi blur saat diperbesar (zoom).

### 💡 Solusi yang Diterapkan di Aplikasi Ini:
- ✅ **`doc.DoNotCompressImages = True`**: Memaksa Word mempertahankan 100% resolusi asli gambar embedded tanpa kompresi.
- ✅ **`wdExportOptimizeForPrint` (300+ DPI)**: Menggunakan standar cetak resolusi tinggi profesional, bukan standar web.
- ✅ **`BitmapMissingFonts = False`**: Memastikan seluruh tipografi tetap berupa kurva vektor murni yang tajam pada zoom 1000%+.
- ✅ **`CreateBookmarks = wdExportCreateHeadingBookmarks`**: Menghasilkan daftar isi / navigasi PDF yang interaktif.
- ✅ **DOCX Image Quality Inspector**: Alat inspeksi resolusi dan DPI asli gambar di dalam file DOCX sebelum konversi.

---

## 🚀 Cara Menjalankan

### Cara 1: Menggunakan Web Interface (Sangat Direkomendasikan)
Cukup jalankan file launcher Windows:
1. Double click file **`run.bat`** (atau ketik `.\run.bat` di terminal).
2. Browser akan terbuka otomatis di alamat: **`http://127.0.0.1:5000`**
3. Tarik & lepas (drag & drop) dokumen Word Anda.
4. Klik **Mulai Konversi ke HD PDF**, lalu unduh hasilnya atau klik **Buka Folder di Komputer**.

Atau melalui terminal manual:
```bash
python app.py
```

---

### Cara 2: Menggunakan CLI / Command Line

Konversi satu file:
```bash
python converter.py "Dokumen Saya.docx"
```

Menentukan lokasi output:
```bash
python converter.py "Dokumen Saya.docx" -o "Hasil_HD.pdf"
```

Konversi semua file Word dalam satu folder sekaligus:
```bash
python converter.py "./folder_word" -o "./folder_pdf"
```

Hanya periksa kualitas & resolusi gambar dalam DOCX (tanpa konversi):
```bash
python converter.py --inspect-only "Dokumen.docx"
```

---

## 📦 Instalasi Dependensi Manual (Jika Diperlukan)

Aplikasi membutuhkan Python 3.8+ dan Microsoft Word (Office) di Windows.
Instal pustaka Python yang dibutuhkan:
```bash
pip install -r requirements.txt
```

---

## 📂 Struktur Proyek

```
PDF converter/
│
├── converter.py         # Engine konversi HD & DOCX image inspector
├── app.py               # Backend Web Server (Flask)
├── run.bat              # Launcher 1-klik untuk Windows
├── requirements.txt     # Daftar pustaka Python
├── README.md            # Dokumentasi lengkap
├── .gitignore           # File pengabaian git
│
├── static/
│   ├── css/
│   │   └── style.css    # Desain Glassmorphism modern & responsif
│   └── js/
│       └── app.js       # Logika interaktif drag & drop dan progress bar
│
└── templates/
    └── index.html       # Antarmuka web dashboard
```

---

## 📌 Cara Push ke GitHub Repository

Berdasarkan pengaturan git sebelumnya:
```bash
git add .
git commit -m "feat: inisialisasi HD Word to PDF Converter dengan preservasi gambar 300 DPI"
git push -u origin main
```
