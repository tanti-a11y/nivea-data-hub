"""Jalankan SEKALI di komputermu (bukan di server) kalau memakai cara B (login akun Google).

Syarat: file credentials.json (OAuth client jenis Desktop app) ada di folder ini.
Hasil: blok [google_oauth] untuk ditempel ke Streamlit Secrets.
"""
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")
print("\nTempel ini ke Streamlit Secrets:\n")
print("[google_oauth]")
print(f'client_id = "{creds.client_id}"')
print(f'client_secret = "{creds.client_secret}"')
print(f'refresh_token = "{creds.refresh_token}"')
