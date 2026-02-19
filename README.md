
# 🎓 Automasi Pengiriman Sertifikat Masal

Aplikasi berbasis web sederhana untuk mengotomatisasi pembuatan dan pengiriman sertifikat menggunakan **Google Slides** dan **Email**.

Dibuat dengan [Streamlit](https://streamlit.io/) dan Python.

## 📋 Fitur Utama

1.  **Menggandakan** template sertifikat dari Google Slides.
2.  **Mengganti** placeholder nama (misal: `{{nama}}`) dengan nama asli peserta secara otomatis.
3.  **Mengekspor** slide tersebut menjadi file PDF.
4.  **Mengirimkan** PDF tersebut ke email peserta sebagai lampiran.
5.  **Mendukung Akun Pribadi (OAuth)** agar tidak terkendala kuota Service Account.
6.  **Manajemen Penyimpanan** (khusus Service Account).

## 🛠️ Persyaratan Sistem

-   Python 3.8 atau lebih baru.
-   Akun Google (Bisa akun pribadi @gmail.com atau Workspace).
-   Akun Gmail (untuk pengiriman email via SMTP).

## 🚀 Cara Instalasi

1.  **Clone atau Download** repository ini.
2.  **Install Library** yang dibutuhkan:
    ```bash
    pip install -r requirements.txt
    ```

## ⚙️ Konfigurasi (PENTING)

Sebelum menjalankan aplikasi, Anda perlu menyiapkan kredensial Google. Ada dua metode:

### Opsi A: Menggunakan Akun Pribadi (Disarankan)
Metode ini menggunakan kuota Google Drive pribadi Anda (15GB).
1.  Buka [Google Cloud Console](https://console.cloud.google.com/).
2.  Buat Project baru -> Aktifkan **Google Drive API** dan **Google Slides API**.
3.  Masuk ke "APIs & Services" -> "OAuth consent screen".
    -   Pilih **External**.
    -   Isi info wajib.
    -   Tambahkan email Anda di bagian **Test Users**.
4.  Masuk ke "Credentials" -> "Create Credentials" -> "OAuth client ID".
    -   Application type: **Desktop App**.
    -   Download file JSON, rename menjadi `client_secret.json` (simpan untuk nanti).

### Opsi B: Menggunakan Service Account
Metode ini untuk use case server-to-server, tapi kuota penyimpanannya terbatas (15GB terpisah dari akun Anda).
1.  Di Google Cloud Console, buat **Service Account**.
2.  Buat Key (JSON), rename menjadi `service_account.json`.

---

### Konfigurasi Environment (`.env`)
1.  Buat file bernama `.env` di folder root aplikasi (bisa copy dari `.env.example`).
2.  Isi konfigurasi berikut:
    ```env
    EMAIL_SENDER=emailanda@gmail.com
    EMAIL_PASSWORD=password_aplikasi_anda
    EMAIL_SUBJECT=Sertifikat Partisipasi
    TARGET_FOLDER_ID=ID_FOLDER_GOOGLE_DRIVE_TUJUAN
    ```
    > **Catatan:**
    > - **App Password:** Untuk Gmail, Anda WAJIB menggunakan **App Password** (bukan password login biasa). [Cara buat App Password](https://myaccount.google.com/apppasswords).
    > - **Target Folder ID:** ID Folder Google Drive tempat sertifikat akan disimpan. (Ambil ID dari URL folder Drive).

### Isi Email (`email_body.txt`)
-   Edit file `email_body.txt` untuk mengatur isi pesan email.
-   Gunakan `{{nama}}` untuk menyisipkan nama peserta.

## ▶️ Cara Penggunaan

1.  Jalankan aplikasi:
    ```bash
    streamlit run app.py
    ```
2.  Di bagian **Konfigurasi** (Sidebar), pilih metode login:
    -   **OAuth Client ID**: Upload `client_secret.json`. Browser akan terbuka untuk login.
    -   **Service Account**: Upload `service_account.json`.
3.  Masukkan **ID Google Slides Template**.
    -   *Pastikan file Slide sudah Di-SHARE ke akun yang Anda pakai login (sebagai Editor).*
4.  Masukkan daftar peserta (Nama, Email).
5.  Klik **"Mulai Kirim Sertifikat"**.

## ⚠️ Troubleshooting

-   **Error 404 (File Not Found)**: Akun yang Anda pakai login **tidak punya akses** ke file Template. Buka Google Slides -> Share -> Masukkan email Anda -> Jadikan **Editor**.
-   **Token Expired/Salah Akun**: Klik tombol **"Logout / Reset Token"** di sidebar, lalu login ulang.
-   **Storage Penuh (Service Account)**: Gunakan fitur "Cek Kuota & Bersihkan" di sidebar (hanya muncul di mode Service Account).

    ```
