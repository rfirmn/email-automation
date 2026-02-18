# 🎓 Automasi Pengiriman Sertifikat Masal

Aplikasi berbasis web sederhana untuk mengotomatisasi pembuatan dan pengiriman sertifikat menggunakan **Google Slides** dan **Email**.

Dibuat dengan [Streamlit](https://streamlit.io/) dan Python.

## 📋 Abstrak

Program ini berfungsi untuk:
1.  **Menggandakan** template sertifikat dari Google Slides.
2.  **Mengganti** placeholder nama (misal: `{{nama}}`) dengan nama asli peserta secara otomatis.
3.  **Mengekspor** slide tersebut menjadi file PDF.
4.  **Mengirimkan** PDF tersebut ke email peserta sebagai lampiran.
5.  **Menghapus** file duplikat sementara di Google Drive untuk menjaga kebersihan penyimpanan.

## 🛠️ Persyaratan Sistem

-   Python 3.8 atau lebih baru.
-   Akun Google Cloud Platform (untuk Google Drive & Slides API).
-   Akun Gmail (untuk pengiriman email via SMTP).

## 🚀 Cara Instalasi

1.  **Clone atau Download** repository ini.
2.  **Install Library** yang dibutuhkan:
    ```bash
    pip install -r requirements.txt
    ```

## ⚙️ Konfigurasi (PENTING)

Sebelum menjalankan aplikasi, Anda perlu menyiapkan beberapa hal:

### 1. Google Service Account
-   Buka [Google Cloud Console](https://console.cloud.google.com/).
-   Buat Project baru.
-   Aktifkan **Google Drive API** dan **Google Slides API**.
-   Masuk ke "Credentials" -> "Create Credentials" -> "Service Account".
-   Setelah dibuat, masuk ke tab "Keys" dan buat key baru (format **JSON**).
-   Simpan file JSON tersebut (nanti akan diupload ke aplikasi).

### 2. File `.env` (Kredensial Email)
-   Buat file bernama `.env` di folder root aplikasi (bisa copy dari `.env.example`).
-   Isi dengan data email pengirim Anda:
    ```env
    EMAIL_SENDER=emailanda@gmail.com
    EMAIL_PASSWORD=password_aplikasi_anda
    EMAIL_SUBJECT=Sertifikat Partisipasi
    ```
    > **Catatan:** Untuk Gmail, Anda WAJIB menggunakan **App Password** (bukan password login biasa). Aktifkan 2-Step Verification di akun Google Anda, lalu buat App Password di [sini](https://myaccount.google.com/apppasswords).

### 3. Isi Email (`email_body.txt`)
-   Edit file `email_body.txt` untuk mengatur isi pesan email.
-   Gunakan `{{nama}}` jika ingin menyisipkan nama peserta di dalam body email.

### 4. Template Google Slides
-   Buat desain sertifikat di Google Slides.
-   Gunakan teks `{{nama}}` di tempat nama peserta akan muncul.
-   **SHARE** file slide tersebut ke email Service Account (email yang berakhiran `@...iam.gserviceaccount.com` yang ada di dalam file JSON) dengan akses **Editor**.
-   Catat **ID** file slide dari URL (contoh: `docs.google.com/presentation/d/ID_FILE_DISINI/edit`).

## ▶️ Cara Penggunaan

1.  Jalankan aplikasi:
    ```bash
    streamlit run app.py
    ```
2.  Browser akan terbuka otomatis.
3.  **Upload** file `service_account.json` pada sidebar.
4.  Masukkan **ID Google Slides Template** pada sidebar.
5.  Masukkan daftar peserta di kolom input utama dengan format:
    ```
    Nama Peserta 1, email1@domain.com
    Nama Peserta 2, email2@domain.com
    ```
6.  Klik tombol **"Mulai Kirim Sertifikat"**.
7.  Tunggu proses selesai dan laporan pengiriman akan muncul.

## ⚠️ Troubleshooting

-   **Error Autentikasi**: Pastikan file JSON benar dan API sudah diaktifkan di Google Cloud Console.
-   **Error Permission**: Pastikan file Google Slides sudah di-share ke email Service Account.
-   **Error Email**: Pastikan App Password benar dan koneksi internet stabil.
