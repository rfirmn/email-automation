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

# OAuth Imports
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle

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

# --- Google OAuth User Auth ---
def authenticate_user_oauth(client_config):
    SCOPES = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/presentations'
    ]
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
                
        if not creds:
            # Create a temporary file for client_secret if it's passed as dict
            # But here we expect client_config to be the dict from json.load
            # Flow.from_client_config is cleaner than saving to file
            flow = InstalledAppFlow.from_client_config(
                client_config, SCOPES)
            creds = flow.run_local_server(port=0)
            
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    drive_service = build('drive', 'v3', credentials=creds)
    slides_service = build('slides', 'v1', credentials=creds)
    return drive_service, slides_service

# --- Clear Token Function ---
def clear_token():
    if os.path.exists('token.pickle'):
        os.remove('token.pickle')
        return True
    return False

# --- Storage Management ---
def check_storage_quota(service):
    try:
        about = service.about().get(fields="storageQuota").execute()
        quota = about.get('storageQuota', {})
        usage = int(quota.get('usage', 0))
        limit_str = quota.get('limit')
        limit = int(limit_str) if limit_str else -1
        
        usage_mb = usage / (1024 * 1024)
        limit_mb = limit / (1024 * 1024) if limit != -1 else -1
        return usage_mb, limit_mb
    except Exception as e:
        return 0, 0

def cleanup_service_account_files(service):
    deleted_count = 0
    errors = []
    try:
        # List all files owned by 'me' (the service account) that are not in trash
        page_token = None
        while True:
            response = service.files().list(
                q="'me' in owners and trashed=false",
                spaces='drive',
                fields='nextPageToken, files(id, name)',
                pageToken=page_token
            ).execute()
            
            for file in response.get('files', []):
                try:
                    service.files().delete(fileId=file['id']).execute()
                    deleted_count += 1
                except Exception as e:
                    errors.append(f"Gagal hapus {file.get('name')}: {str(e)}")
            
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break
                
        # Also empty headers trash if possible (not always necessary if we delete permanently)
        try:
             service.files().emptyTrash().execute()
        except:
             pass
             
        return deleted_count, errors
    except Exception as e:
        return deleted_count, [str(e)]

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

st.sidebar.info("Email Pengirim, Password, Subjek, dan Body diambil dari konfigurasi file (.env & email_body.txt).")

# --- Authentication Method Selection ---
st.sidebar.markdown("---")
st.sidebar.header("🔑 Autentikasi Google")
auth_method = st.sidebar.radio(
    "Metode Login:",
    ("OAuth Client ID (Rekomendasi - Akun Pribadi)", "Service Account (Akun Kantor/Berbayar)")
)

uploaded_file = None
if auth_method == "Service Account (Akun Kantor/Berbayar)":
    uploaded_file = st.sidebar.file_uploader("Upload service_account.json", type="json")
    auth_type = "service_account"
else:
    uploaded_file = st.sidebar.file_uploader("Upload client_secret.json", type="json")
    auth_type = "oauth"
    if os.path.exists('token.pickle'):
        st.sidebar.success("Token tersimpan ditemukan! (Otomatis login)")
        if st.sidebar.button("Logout / Reset Token"):
            if clear_token():
                st.sidebar.warning("Token dihapus. Silakan refresh halaman.")
                st.rerun()

# --- Google Slide Template ---
template_id = st.sidebar.text_input("ID Google Slides Template", help="ID bisa diambil dari URL Google Slides (bagian acak di tengah URL).")

# --- Storage Management UI ---
# HANYA TAMPILKAN JIKA MENGGUNAKAN SERVICE ACCOUNT
if auth_type == "service_account":
    st.sidebar.markdown("---")
    st.sidebar.header("🧹 Manajemen Penyimpanan")
    
    if uploaded_file is not None:
        if st.sidebar.button("Cek Kuota & Bersihkan"):
            try:
                # Re-auth just for checking
                uploaded_file.seek(0)
                service_account_info = json.load(uploaded_file)
                drive_service, _ = authenticate_google(service_account_info)
                uploaded_file.seek(0) # Reset pointer

                usage, limit = check_storage_quota(drive_service)
                limit_str = f"{limit:.2f} MB" if limit != -1 else "Unlimited"
                st.sidebar.write(f"**Terpakai:** {usage:.2f} MB")
                st.sidebar.write(f"**Limit:** {limit_str}")
                
                if usage > 13000: # Warning near 15GB
                    st.sidebar.warning("Penyimpanan hampir penuh!")
                
                with st.sidebar.status("Membersihkan file temporary...", expanded=True) as status:
                    count, errors = cleanup_service_account_files(drive_service)
                    status.write(f"Berhasil menghapus {count} file.")
                    if errors:
                        status.warning(f"Error pada {len(errors)} file.")
                
                if count > 0:
                    st.sidebar.success(f"Berhasil mengosongkan ruang! ({count} file dihapus)")
                else:
                    st.sidebar.info("Tidak ada file yang perlu dihapus.")
                    
            except Exception as e:
                st.sidebar.error(f"Gagal cek storage: {e}")
    else:
        st.sidebar.caption("Upload Service Account JSON dulu untuk cek storage.")
else:
    # OAUTH MODE
    st.sidebar.markdown("---")
    st.sidebar.info("ℹ️ Mode OAuth menggunakan penyimpanan pribadi Anda.")
# --- Main Area ---
st.subheader("📋 Data Peserta")
st.markdown("Masukkan data peserta dengan format: `Nama Lengkap, email@target.com` (Satu peserta per baris)")

raw_participants = st.text_area("List Peserta", height=200, placeholder="Budi Santoso, budi@example.com\nSiti Aminah, siti@example.com")

if st.button("🚀 Mulai Kirim Sertifikat", type="primary"):
    # Load settings from env and file
    email_sender = os.getenv("EMAIL_SENDER")
    email_password = os.getenv("EMAIL_PASSWORD")
    email_subject = os.getenv("EMAIL_SUBJECT")
    target_folder_id = os.getenv("TARGET_FOLDER_ID")
    email_body_template = load_email_body()

    # Validasi Input dan Config
    missing_config = []
    if not email_sender: missing_config.append("EMAIL_SENDER di .env")
    if not email_password: missing_config.append("EMAIL_PASSWORD di .env")
    if not email_subject: missing_config.append("EMAIL_SUBJECT di .env")
    if not email_body_template: missing_config.append("File email_body.txt")

    if not email_body_template: missing_config.append("File email_body.txt")

    if missing_config:
        st.error(f"Konfigurasi belum lengkap: {', '.join(missing_config)}")
    elif (not uploaded_file and not os.path.exists('token.pickle')) or not template_id or not raw_participants:
        st.error("Mohon lengkapi Credential JSON, Template ID, dan Data Peserta.")
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
        # 2. Init Google Auth
        try:
            if auth_type == "service_account":
                service_account_info = json.load(uploaded_file)
                drive_service, slides_service = authenticate_google(service_account_info)
                st.success("Autentikasi Service Account Berhasil!")
            else:
                # OAuth
                if uploaded_file:
                    client_config = json.load(uploaded_file)
                elif os.path.exists('token.pickle'):
                    # If we have a token but no file, we hope the token is valid. 
                    # If invalid, we need client_config which might be missing if file not uploaded.
                    # For robust code, ideally we always ask for file or store client_config too.
                    # But often token.pickle is enough.
                    client_config = None 
                else:
                    st.error("Butuh file client_secret.json untuk login pertama kali.")
                    st.stop()
                    
                drive_service, slides_service = authenticate_user_oauth(client_config)
                st.success("Autentikasi OAuth User Berhasil!")
                
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
                
                # Use target folder if specified
                if target_folder_id:
                     body['parents'] = [target_folder_id]

                try:
                    drive_response = drive_service.files().copy(
                        fileId=template_id, body=body).execute()
                    copy_id = drive_response.get('id')
                except Exception as copy_error:
                    error_msg = str(copy_error)
                    if "404" in error_msg:
                        log_entry['Status'] = '❌ Template Tidak Ditemukan'
                        log_entry['Detail'] = (
                            f"File Template ID '{template_id}' tidak ditemukan atau tidak bisa diakses "
                            f"oleh akun yang login saat ini. Pastikan file ada dan Anda memiliki akses."
                        )
                    elif "403" in error_msg:
                         log_entry['Status'] = '❌ Akses Ditolak'
                         log_entry['Detail'] = "Akun tidak memiliki izin untuk mengedit/copy template ini."
                    else:
                        log_entry['Status'] = '❌ Gagal Copy Template'
                        log_entry['Detail'] = error_msg
                    
                    # Skip to next participant if copy failed
                    results.append(log_entry)
                    progress_bar.progress((idx + 1) / total)
                    continue

                try:
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
                    personal_body = email_body_template.replace("{{nama}}", name) if email_body_template else ""
                    
                    sent, msg = send_email_with_attachment(
                        email_sender, email_password, email, 
                        email_subject, personal_body, pdf_bytes, f"Sertifikat_{name.replace(' ', '_')}.pdf"
                    )
                    
                    if sent:
                        log_entry['Status'] = '✅ Berhasil'
                    else:
                        log_entry['Status'] = '❌ Gagal Email'
                        log_entry['Detail'] = msg

                finally:
                    # E. Cleanup (Delete temp file) - ALWAYS RUN
                    try:
                        drive_service.files().delete(fileId=copy_id).execute()
                    except Exception as cleanup_error:
                        # Log cleanup error but don't fail the row if email was sent
                        if log_entry['Status'] == '✅ Berhasil':
                             log_entry['Detail'] += f" (Gagal hapus temp: {str(cleanup_error)})"
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
