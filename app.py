import streamlit as st
import pandas as pd
import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import time
from dotenv import load_dotenv

# --- Load Environment Variables ---
load_dotenv()

# --- Page Configuration ---
st.set_page_config(
    page_title="Automasi Sertifikat Masal",
    page_icon="🎓",
    layout="wide"
)

# --- Validasi Email Sederhana ---
def is_valid_email(email):
    return "@" in email and "." in email

# --- Google API Auth ---
def authenticate_google(service_account_info):
    creds = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=[
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/presentations'
        ]
    )
    drive_service = build('drive', 'v3', credentials=creds)
    slides_service = build('slides', 'v1', credentials=creds)
    return drive_service, slides_service

# --- Fungsi Kirim Email ---
def send_email_with_attachment(sender_email, sender_password, recipient_email, subject, body, pdf_bytes, filename):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    part = MIMEApplication(pdf_bytes, Name=filename)
    part['Content-Disposition'] = f'attachment; filename="{filename}"'
    msg.attach(part)

    try:
        if "gmail" in sender_email:
            server = smtplib.SMTP('smtp.gmail.com', 587)
        else:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        return True, "Email terkirim."
    except Exception as e:
        return False, str(e)

# --- Load Email Body ---
def load_email_body():
    try:
        with open("email_body.txt", "r") as f:
            return f.read()
    except FileNotFoundError:
        return None

# --- UI Setup ---
st.title("🎓 Automasi Pengiriman Sertifikat Masal")

# --- Tutorial Expander ---
with st.expander("📚 Panduan Setup & Tutorial (Baca Dulu)", expanded=False):
    st.markdown("""
    ### Langkah Persiapan:
    1.  **Google Service Account**:
        - Upload file `service_account.json` di sidebar.
    2.  **Environment Variables**:
        - Buat file `.env` di folder aplikasi.
        - Isi dengan format:
          ```
          EMAIL_SENDER=emailanda@gmail.com
          EMAIL_PASSWORD=password_aplikasi_anda
          EMAIL_SUBJECT=Subjek Email
          ```
    3.  **Email Body**:
        - Edit file `email_body.txt` untuk mengubah isi email. Gunakan `{{nama}}` untuk nama peserta.
    4.  **Sertifikat Template**:
        - Pastikan ID Google Slides benar dan file sudah dishare ke email service account.
    """)

# --- Sidebar Configuration ---
st.sidebar.header("⚙️ Konfigurasi")

uploaded_file = st.sidebar.file_uploader("Upload service_account.json", type="json")
template_id = st.sidebar.text_input("ID Google Slides Template", help="ID bisa diambil dari URL Google Slides (bagian acak di tengah URL).")

st.sidebar.info("Email Pengirim, Password, Subjek, dan Body diambil dari konfigurasi file (.env & email_body.txt).")

# --- Main Area ---
st.subheader("📋 Data Peserta")
st.markdown("Masukkan data peserta dengan format: `Nama Lengkap, email@target.com` (Satu peserta per baris)")

raw_participants = st.text_area("List Peserta", height=200, placeholder="Budi Santoso, budi@example.com\nSiti Aminah, siti@example.com")

if st.button("🚀 Mulai Kirim Sertifikat", type="primary"):
    # Load settings from env and file
    email_sender = os.getenv("EMAIL_SENDER")
    email_password = os.getenv("EMAIL_PASSWORD")
    email_subject = os.getenv("EMAIL_SUBJECT")
    email_body_template = load_email_body()

    # Validasi Input dan Config
    missing_config = []
    if not email_sender: missing_config.append("EMAIL_SENDER di .env")
    if not email_password: missing_config.append("EMAIL_PASSWORD di .env")
    if not email_subject: missing_config.append("EMAIL_SUBJECT di .env")
    if not email_body_template: missing_config.append("File email_body.txt")

    if missing_config:
        st.error(f"Konfigurasi belum lengkap: {', '.join(missing_config)}")
    elif not uploaded_file or not template_id or not raw_participants:
        st.error("Mohon lengkapi Service Account, Template ID, dan Data Peserta.")
    else:
        # 1. Parse Data Peserta
        participants = []
        try:
            lines = raw_participants.strip().split('\n')
            for line in lines:
                parts = line.split(',')
                if len(parts) >= 2:
                    name = parts[0].strip()
                    email_addr = parts[1].strip()
                    if is_valid_email(email_addr):
                        participants.append({'nama': name, 'email': email_addr})
            
            df_participants = pd.DataFrame(participants)
            if df_participants.empty:
               st.error("Format data peserta tidak valid. Pastikan format: Nama, Email")
               st.stop()
               
            st.info(f"Terdeteksi {len(participants)} peserta.")
            
        except Exception as e:
            st.error(f"Gagal memparsing data peserta: {e}")
            st.stop()

        # 2. Init Google Auth
        try:
            service_account_info = json.load(uploaded_file)
            drive_service, slides_service = authenticate_google(service_account_info)
            st.success("Autentikasi Google Berhasil!")
        except Exception as e:
            st.error(f"Autentikasi Gagal: {e}")
            st.stop()

        # 3. Processing Loop
        progress_bar = st.progress(0)
        status_log = st.empty()
        results = []

        total = len(participants)
        
        for idx, p in enumerate(participants):
            name = p['nama']
            email = p['email']
            status_text = f"Memproses ({idx+1}/{total}): {name}..."
            status_log.text(status_text)
            
            start_time = time.strftime("%H:%M:%S")
            log_entry = {'Nama': name, 'Email': email, 'Waktu': start_time, 'Status': '', 'Detail': ''}
            
            try:
                # A. Copy Template
                copy_title = f"Sertifikat - {name}"
                body = {'name': copy_title}
                drive_response = drive_service.files().copy(
                    fileId=template_id, body=body).execute()
                copy_id = drive_response.get('id')
                
                # B. Replace Text
                requests = [
                    {
                        'replaceAllText': {
                            'containsText': {
                                'text': '{{nama}}',
                                'matchCase': True
                            },
                            'replaceText': name
                        }
                    }
                ]
                slides_service.presentations().batchUpdate(
                    presentationId=copy_id, body={'requests': requests}).execute()
                
                # C. Export to PDF
                request_pdf = drive_service.files().export_media(
                    fileId=copy_id, mimeType='application/pdf')
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request_pdf)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                
                pdf_bytes = fh.getvalue()
                
                # D. Send Email
                # Personalize body if it contains {{nama}}
                personal_body = email_body_template.replace("{{nama}}", name)
                
                sent, msg = send_email_with_attachment(
                    email_sender, email_password, email, 
                    email_subject, personal_body, pdf_bytes, f"Sertifikat_{name.replace(' ', '_')}.pdf"
                )
                
                if sent:
                    log_entry['Status'] = '✅ Berhasil'
                else:
                    log_entry['Status'] = '❌ Gagal Email'
                    log_entry['Detail'] = msg

                # E. Cleanup (Delete temp file)
                try:
                    drive_service.files().delete(fileId=copy_id).execute()
                except:
                    pass 

            except Exception as e:
                log_entry['Status'] = '❌ Error System'
                log_entry['Detail'] = str(e)
            
            results.append(log_entry)
            progress_bar.progress((idx + 1) / total)

        status_log.text("Proses Selesai!")
        
        # 4. Report
        st.divider()
        st.subheader("Laporan Pengiriman")
        df_results = pd.DataFrame(results)
        st.dataframe(df_results)
