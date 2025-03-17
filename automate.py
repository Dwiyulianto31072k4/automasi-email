import streamlit as st
import os
import base64
import pickle
import logging
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import re

# === Konfigurasi Logging ===
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# === Konfigurasi Gmail API ===
SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]
TOKEN_PATH = "token.json"
CREDENTIALS_PATH = "client_secret_739705307269-e8vmb0lv0n493qln63is9ajomqaa0fmh.apps.googleusercontent.com.json"

# === Fungsi Autentikasi Gmail API ===
def authenticate_gmail():
    creds = None
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "wb") as token:
            pickle.dump(creds, token)
    
    # Simpan service di session state agar bisa digunakan di seluruh aplikasi
    st.session_state["service"] = build("gmail", "v1", credentials=creds)
    return st.session_state["service"]

# === Streamlit UI ===
st.title("📧 Automasi Distribusi Data Email")

# Tombol autentikasi Gmail API
if st.button("Authenticate Gmail API"):
    try:
        service = authenticate_gmail()
        st.success("Autentikasi Gmail API berhasil!")
    except Exception as e:
        st.error(f"Autentikasi gagal: {e}")

# === Upload File Excel ===
st.subheader("Upload File Data")
uploaded_file_1 = st.file_uploader("Upload File 1: CEK BAHAN DAN EMAIL BAHAN BH.xlsx", type=["xlsx"])
uploaded_file_2 = st.file_uploader("Upload File 2: 03. PIC Data Leads - 2025.xlsx", type=["xlsx"])

if uploaded_file_1 and uploaded_file_2:
    df_data = pd.read_excel(uploaded_file_1, sheet_name="TABEL FEB25")
    df_email = pd.read_excel(uploaded_file_2, sheet_name="PIC 2025")

    st.success("✅ Kedua file berhasil diupload!")
    st.dataframe(df_data.head())

    # === Bersihkan Data ===
    df_data = df_data.iloc[:, [1, 2, 3, 4, 5]]
    df_data.columns = ["OFFICE_CODE", "NAMA_CABANG", "LEADS_NMC", "LEADS_AMITRA", "GRAND_TOTAL"]
    df_data = df_data.dropna(subset=["OFFICE_CODE"])

    # Hapus baris yang tidak perlu
    df_data = df_data[df_data["OFFICE_CODE"] != "OFFICE_CODE"]

    # Konversi angka ke integer
    for col in ["LEADS_NMC", "LEADS_AMITRA", "GRAND_TOTAL"]:
        df_data[col] = pd.to_numeric(df_data[col], errors="coerce").fillna(0).astype(int)

    # Deteksi area
    def detect_area(row):
        if isinstance(row["OFFICE_CODE"], str) and "JUMLAH DATA AREA" in row["OFFICE_CODE"]:
            return row["OFFICE_CODE"]
        elif isinstance(row["NAMA_CABANG"], str) and "JUMLAH DATA AREA" in row["NAMA_CABANG"]:
            return row["NAMA_CABANG"]
        else:
            return None

    df_data["AREA"] = df_data.apply(detect_area, axis=1)
    df_data["AREA"] = df_data["AREA"].fillna(method="ffill")

    # Hapus baris tidak perlu
    df_cleaned = df_data.dropna(subset=["AREA"])
    df_cleaned = df_cleaned[~df_cleaned["OFFICE_CODE"].astype(str).str.contains("JUMLAH DATA AREA", na=False)]
    df_cleaned = df_cleaned[~df_cleaned["NAMA_CABANG"].astype(str).str.contains("Grand Total", na=False)]

    # === Fungsi Membuat Draft Email ===
    def create_draft(service, sender, to, cc, subject, body):
        message = MIMEMultipart()
        message["From"] = sender
        message["To"] = to
        message["Cc"] = cc
        message["Subject"] = subject
        message.attach(MIMEText(body, "html"))

        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
        draft = {"message": {"raw": raw_message}}

        try:
            service.users().drafts().create(userId="me", body=draft).execute()
            logging.info(f"Draft email untuk {subject} berhasil dibuat.")
            return True
        except Exception as e:
            logging.error(f"Error membuat draft email: {e}")
            return False

    # === Generate Email untuk Semua Area ===
    if st.button("Generate All Emails"):
        # Pastikan service sudah ada
        if "service" not in st.session_state:
            st.error("⚠️ Harap autentikasi Gmail API terlebih dahulu!")
        else:
            service = st.session_state["service"]

            for area, df_area in df_cleaned.groupby("AREA"):
                cleaned_area = str(area).replace("JUMLAH DATA AREA ", "").strip()

                # Format email body
                email_body = f"""
                <html>
                <body>
                    <p>Yth. Rekan-Rekan FIFGROUP - <b>{cleaned_area}</b>,</p>
                    <p>Data Good Customer telah tersedia.</p>
                    <table border="1" cellpadding="5" cellspacing="0" 
                           style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; text-align: center;">
                        <thead>
                            <tr style="background-color: #A8D08D; font-weight: bold;">
                                <th>Office Code</th>
                                <th>Nama Cabang</th>
                                <th>Leads NMC</th>
                                <th>Leads Amitra</th>
                                <th>Grand Total</th>
                            </tr>
                        </thead>
                        <tbody>
                """

                for _, row in df_area.iterrows():
                    email_body += f"""
                    <tr>
                        <td>{row['OFFICE_CODE']}</td>
                        <td>{row['NAMA_CABANG']}</td>
                        <td>{row['LEADS_NMC']}</td>
                        <td>{row['LEADS_AMITRA']}</td>
                        <td><b>{row['GRAND_TOTAL']}</b></td>
                    </tr>
                    """

                email_body += "</tbody></table></body></html>"

                # Buat draft email
                subject = f"[CRM DATA MINING INFO] - DISTRIBUSI DATA {cleaned_area}"
                success = create_draft(service, "your-email@gmail.com", "recipient@example.com", "", subject, email_body)

                if success:
                    st.success(f"✅ Draft email untuk {cleaned_area} berhasil dibuat!")
                else:
                    st.error(f"⚠️ Gagal membuat draft email untuk {cleaned_area}.")
