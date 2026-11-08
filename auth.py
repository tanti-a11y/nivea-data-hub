"""Kredensial Google untuk server online.

Dua cara (pilih salah satu di Streamlit secrets):
  A. [gcp_service_account]  -> isi seluruh isi file service_account.json
  B. [google_oauth]         -> client_id, client_secret, refresh_token (dari login Google pemilikmu sekali saja)
Tidak ada kredensial yang disimpan di kode.
"""
import json
import os

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


class AuthError(Exception):
    pass


def _plain(x):
    """Ubah objek secrets Streamlit (mirip dict) jadi dict biasa."""
    if hasattr(x, "items"):
        return {k: _plain(v) for k, v in x.items()}
    return x


def auth_kind(secrets):
    if "gcp_service_account" in secrets:
        return "service_account"
    if "google_oauth" in secrets:
        return "oauth"
    return None


def get_credentials(secrets, scopes=SCOPES):
    kind = auth_kind(secrets)
    if kind == "service_account":
        from google.oauth2.service_account import Credentials
        info = _plain(secrets["gcp_service_account"])
        if "private_key" in info:
            info["private_key"] = info["private_key"].replace("\\n", "\n")
        return Credentials.from_service_account_info(info, scopes=scopes)
    if kind == "oauth":
        from google.oauth2.credentials import Credentials
        o = _plain(secrets["google_oauth"])
        miss = [k for k in ("client_id", "client_secret", "refresh_token") if not o.get(k)]
        if miss:
            raise AuthError("Bagian [google_oauth] di secrets belum lengkap: " + ", ".join(miss))
        return Credentials(
            None, refresh_token=o["refresh_token"], client_id=o["client_id"], client_secret=o["client_secret"],
            token_uri=o.get("token_uri", "https://oauth2.googleapis.com/token"), scopes=scopes)
    raise AuthError("Kredensial Google belum diisi di Streamlit secrets "
                    "([gcp_service_account] atau [google_oauth]). Lihat PANDUAN_DEPLOY.md.")


def service_account_email(secrets):
    if auth_kind(secrets) == "service_account":
        return _plain(secrets["gcp_service_account"]).get("client_email", "")
    return ""


def build_services(secrets):
    """(drive_service, sheets_service) untuk engine yang memakai googleapiclient."""
    from googleapiclient.discovery import build
    creds = get_credentials(secrets)
    try:
        import google_auth_httplib2
        import httplib2
        mk = lambda: google_auth_httplib2.AuthorizedHttp(creds, http=httplib2.Http(timeout=300))
        return build("drive", "v3", http=mk(), cache_discovery=False), build("sheets", "v4", http=mk(), cache_discovery=False)
    except ImportError:
        return (build("drive", "v3", credentials=creds, cache_discovery=False),
                build("sheets", "v4", credentials=creds, cache_discovery=False))


def write_runtime_files(folder, secrets, cfg):
    """Engine lama membaca kredensial & config dari file -> tulis ke folder sementara (tidak ikut ke GitHub)."""
    os.makedirs(folder, exist_ok=True)
    kind = auth_kind(secrets)
    if kind == "service_account":
        info = _plain(secrets["gcp_service_account"])
        if "private_key" in info:
            info["private_key"] = info["private_key"].replace("\\n", "\n")
        cred_name = "service_account.json"
        with open(os.path.join(folder, cred_name), "w", encoding="utf-8") as f:
            json.dump(info, f)
    elif kind == "oauth":
        o = _plain(secrets["google_oauth"])
        cred_name = "credentials.json"
        client = {"installed": {
            "client_id": o["client_id"], "client_secret": o["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": o.get("token_uri", "https://oauth2.googleapis.com/token"),
            "redirect_uris": ["http://localhost"]}}
        with open(os.path.join(folder, cred_name), "w", encoding="utf-8") as f:
            json.dump(client, f)
        token = {"token": None, "refresh_token": o["refresh_token"], "client_id": o["client_id"],
                 "client_secret": o["client_secret"], "token_uri": client["installed"]["token_uri"],
                 "scopes": SCOPES}
        with open(os.path.join(folder, "token.json"), "w", encoding="utf-8") as f:
            json.dump(token, f)
    else:
        raise AuthError("Kredensial Google belum diisi di Streamlit secrets.")
    cfg = dict(cfg, credentials_file=cred_name, token_file="token.json", report_dir="laporan")
    with open(os.path.join(folder, "config.json"), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.makedirs(os.path.join(folder, "laporan"), exist_ok=True)
    return os.path.join(folder, "config.json")
