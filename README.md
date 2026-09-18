# AREKO BCA - Rekapitulasi Mutasi Rekening Koran Bank Central Asia

Aplikasi cerdas dan otomatis untuk menganalisis, mengekstraksi, dan merekapitulasi data rekening koran **Bank Central Asia (BCA)** dari berkas PDF menjadi Formulir Validasi Mutasi Rekening resmi (Format PDF A4 siap cetak dan Microsoft Excel Spreadsheet dinamis).

---

## ✨ Fitur Utama

1. **Multi-Engine Parsing Presisi Tinggi**:
   - Dukungan ekstraksi teks digital PDF native berkecepatan tinggi.
   - Fallback OCR beresolusi tinggi otomatis (Tesseract OCR Engine) untuk dokumen hasil scan.
   - Deteksi otomatis periode bulan transaksi (Bahasa Indonesia).
   - Validasi saldo berjalan (running balance), mutasi Debet/Kredit, dan rekonsiliasi saldo awal hingga saldo akhir 100% akurat sesuai ringkasan resmi BCA.

2. **Ringkasan Mutasi Komprehensif**:
   - Menghitung frekuensi Debet & Kredit.
   - Menghitung total perputaran dana mutasi Debet & Kredit.
   - Menganalisis saldo tertinggi (max), saldo terendah (min), dan saldo rata-rata (average).
   - Menghitung baris **Rata-Rata** keseluruhan bulan yang diunggah.

3. **Ekspor Dokumen Standar Perbankan**:
   - **PDF Resmi**: Dokumen A4 siap cetak dilengkapi informasi nasabah, nomor rekening, catatan Operation Head, dan kolom tanda tangan Sales Officer & Operation Head.
   - **Excel Spreadsheet (.xlsx)**: Lembar kerja dinamis lengkap dengan tabel rekapitulasi dan rincian transaksi per bulan di kolom rincian.

4. **Desain Antarmuka Modern & Bersih**:
   - Tampilan tema terang (*Light Mode*) yang profesional dan responsif.
   - Tabel pratinjau interaktif dengan penanda visual khusus pada baris rata-rata.

---

## 🛠️ Instalasi & Menjalankan Lokal

### 1. Kebutuhan Sistem
Pastikan sistem Anda telah memiliki **Python 3.9+** serta dependensi sistem `tesseract-ocr` dan `poppler-utils`:

- **macOS (via Homebrew)**:
  ```bash
  brew install tesseract poppler
  ```
- **Ubuntu/Debian / Streamlit Cloud**:
  Sudah terkonfigurasi di `packages.txt` (`tesseract-ocr`, `poppler-utils`).

### 2. Instal Dependensi Python
```bash
cd BCA
pip install -r requirements.txt
```

### 3. Jalankan Aplikasi
```bash
streamlit run app.py
```

Akses aplikasi melalui browser di `http://localhost:8501`.

---

## ☁️ Panduan Deploy ke Streamlit Cloud

1. Buat repository baru di GitHub (misalnya: `AREKO-BCA`).
2. Hubungkan remote dan lakukan push:
   ```bash
   git remote add origin https://github.com/<username>/AREKO-BCA.git
   git branch -M main
   git push -u origin main
   ```
3. Buka [share.streamlit.io](https://share.streamlit.io) dan klik **New App**.
4. Pilih repositori `AREKO-BCA`, branch `main`, dan set file path ke `app.py`.
5. Streamlit Cloud akan otomatis membaca `packages.txt` dan `requirements.txt` untuk menginstal Tesseract & Poppler.
