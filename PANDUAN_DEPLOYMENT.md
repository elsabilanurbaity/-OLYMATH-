# 🏆 OLYMATH – Panduan Deployment Lengkap

## Daftar Isi
1. [Persiapan File](#1-persiapan-file)
2. [Google Colab – Training Model](#2-google-colab--training-model)
3. [GitHub – Upload Project](#3-github--upload-project)
4. [Streamlit Cloud – Deploy Aplikasi](#4-streamlit-cloud--deploy-aplikasi)
5. [Troubleshooting](#5-troubleshooting)

---

## 1. Persiapan File

Pastikan kamu memiliki file-file berikut:

```
olymath/
├── app.py                    ← Aplikasi Streamlit utama
├── model.pkl                 ← Model hasil training di Colab
├── requirements.txt          ← Dependensi Python
├── OLYMATH_Training.ipynb    ← Notebook training (Google Colab)
└── PANDUAN_DEPLOYMENT.md     ← File ini
```

---

## 2. Google Colab – Training Model

### Langkah-langkah:

**A. Buka Google Colab**
1. Buka [colab.research.google.com](https://colab.research.google.com)
2. Klik **File → Upload notebook**
3. Upload file `OLYMATH_Training.ipynb`

**B. Upload dataset ke Google Drive**
1. Buka [drive.google.com](https://drive.google.com)
2. Buat folder baru bernama `OLYMATH`
3. Upload file `dataset.xlsx` ke dalam folder tersebut
4. Pastikan path: `MyDrive/OLYMATH/dataset.xlsx`

**C. Jalankan notebook**
1. Di Colab, klik **Runtime → Run all** (atau jalankan sel satu per satu dari atas)
2. Saat muncul prompt **Mount Google Drive**, klik **Connect to Google Drive** dan izinkan akses
3. Tunggu semua sel selesai dijalankan (sekitar 3-5 menit)
4. Cek apakah `model.pkl` sudah muncul di `MyDrive/OLYMATH/`

**D. Download model.pkl**
1. Buka Google Drive → folder `OLYMATH`
2. Klik kanan `model.pkl` → **Download**
3. Simpan di folder lokal project kamu

> ✅ **Hasil yang diharapkan:** Akurasi model sekitar **95%+**

---

## 3. GitHub – Upload Project

### A. Buat Repository Baru
1. Buka [github.com](https://github.com) → Login
2. Klik tombol **+** (pojok kanan atas) → **New repository**
3. Isi:
   - **Repository name:** `olymath`
   - **Description:** `Sistem Seleksi Olimpiade Matematika berbasis Machine Learning`
   - **Visibility:** Public *(wajib untuk Streamlit free tier)*
4. Klik **Create repository**

### B. Upload File
**Cara 1 – Via Web (Paling mudah):**
1. Di halaman repository, klik **Add file → Upload files**
2. Drag & drop semua file berikut:
   - `app.py`
   - `model.pkl`
   - `requirements.txt`
   - `OLYMATH_Training.ipynb` *(opsional)*
3. Klik **Commit changes**

**Cara 2 – Via Terminal (Jika sudah install Git):**
```bash
# Di folder project kamu
git init
git add app.py model.pkl requirements.txt
git commit -m "Initial commit: OLYMATH Streamlit App"
git branch -M main
git remote add origin https://github.com/USERNAME/olymath.git
git push -u origin main
```
*(Ganti `USERNAME` dengan username GitHub kamu)*

### C. Verifikasi
Pastikan repository kamu memiliki minimal:
```
📁 olymath (repository)
├── 📄 app.py
├── 📄 model.pkl
└── 📄 requirements.txt
```

---

## 4. Streamlit Cloud – Deploy Aplikasi

### A. Daftar / Login Streamlit Cloud
1. Buka [share.streamlit.io](https://share.streamlit.io)
2. Klik **Continue with GitHub** → Login dengan akun GitHub kamu

### B. Deploy Aplikasi Baru
1. Klik tombol **New app**
2. Isi form:
   - **Repository:** pilih `USERNAME/olymath`
   - **Branch:** `main`
   - **Main file path:** `app.py`
   - **App URL (opsional):** misalnya `olymath-olympiade`
3. Klik **Deploy!**

### C. Tunggu Proses Build
- Streamlit akan menginstal dependencies dari `requirements.txt` otomatis
- Proses ini memakan waktu **2-5 menit** pertama kali
- Kamu bisa memantau log di panel sebelah kanan

### D. Aplikasi Siap!
- URL aplikasimu akan berformat: `https://USERNAME-olymath.streamlit.app`
- Bagikan URL ini ke dosen dan teman-teman!

---

## 5. Troubleshooting

### ❌ Error: "model.pkl not found"
**Penyebab:** File `model.pkl` belum diupload ke GitHub  
**Solusi:** Pastikan `model.pkl` ada di root repository GitHub, sejajar dengan `app.py`

### ❌ Error: "Module not found: sklearn"
**Penyebab:** `requirements.txt` tidak terbaca  
**Solusi:** Pastikan file `requirements.txt` ada di root repository dengan isi yang benar

### ❌ Aplikasi lambat saat pertama dibuka
**Penyebab:** Streamlit "cold start" – normal untuk free tier  
**Solusi:** Tunggu 30-60 detik. Setelah itu aplikasi akan berjalan cepat

### ❌ Error upload Excel: "Kolom tidak ditemukan"
**Penyebab:** Nama kolom di file Excel berbeda dari yang diharapkan  
**Solusi:** Gunakan template Excel yang disediakan di aplikasi (klik "Download Template Excel")

### ❌ Nilai tinggi tapi confidence rendah
**Penjelasan:** Ini perilaku normal algoritma klasifikasi. Siswa yang berada di dekat batas antar-kategori (misalnya rata-rata 78-82) akan memiliki keyakinan model yang lebih rendah karena berada di "zona peralihan". Gunakan **Skor Kesiapan** sebagai patokan utama membandingkan siswa dalam kategori yang sama — skor ini selalu naik seiring nilai naik.

---

## 📌 Catatan Penting

- **model.pkl** harus selalu diupload ke GitHub bersama `app.py`
- Jika dataset diperbarui, jalankan ulang notebook di Colab dan upload ulang `model.pkl`
- Streamlit Cloud free tier memiliki batas resource, tapi cukup untuk demo tugas kuliah

---

*Dibuat untuk Tugas Data Mining | Program Studi Pendidikan Matematika*  
*OLYMATH – OLYmpiade MATHematics 🏆*
