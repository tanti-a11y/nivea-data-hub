#!/usr/bin/env python3
"""
sync_nivea_tracker.py
---------------------
Mengisi otomatis tracker "Tracker_Daily_Nivea_Fiks_okei_September.xlsx"
(tab Celeb, MCN, TOP VIP, RISING, KLUB VIP, Community) dari:

  1. File harian  Creator_List_YYYYMMDD-YYYYMMDD_..._Shop_Tokopedia.xlsx
        Creator username                 -> baris creator di tab
        Affiliate GMV                    -> GMV Total
        Affiliate LIVE GMV               -> GMV LS
        Affiliate shoppable video GMV    -> GMV SV
        Affiliate product card GMV       -> GMV SL
  3. File MCN: 'Affiliate partnerships overview_Agency list_YYYY-MM-DD-YYYY-MM-DD.xlsx' (folder MCN Daily)
        -> tab MCN: Agency name = Creator, GMV = GMV Total, 'LIVE streams' = LS, 'Videos' = SV
  2. Google Sheet data harian, tab LS dan tab SV
        LS = jumlah Live ID UNIK per creator per tanggal
        SV = jumlah Video ID UNIK per creator per tanggal

ATURAN
  - Creator HANYA diisi kalau namanya sudah ada di tab (Celeb, MCN, TOP VIP,
    RISING, KLUB VIP, Community).
  - Pengecualian: tab Community boleh DITAMBAH creator baru (yang belum ada di
    tab manapun) kalau GMV Total hari itu >= COMMUNITY_MIN_GMV.

CARA KERJA
  Script mengunduh tracker dari Drive, menyimpan backup lokal, mengisi
  sel-sel yang perlu SAJA (lewat edit langsung isi file .xlsx, jadi format,
  warna, dan dropdown PIC tidak berubah), lalu mengunggah kembali ke file
  yang sama di Drive.

PAKAI
  py sync_nivea_tracker.py                        -> isi tracker dari file Creator_List
  py sync_nivea_tracker.py --lihat-gmv            -> tampilkan total GMV tiap file di folder (tanpa mengisi apa pun)
  py sync_nivea_tracker.py --perbaiki-template    -> betulkan template (cukup sekali, SEBELUM mengisi)
  (file credentials.json & token.json sama seperti script sebelumnya)
"""

import html
import io
import json
import os
import re
import sys
import zipfile
import datetime as dt
from collections import OrderedDict
from typing import Optional
from xml.sax.saxutils import escape as xml_escape

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials as UserCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import openpyxl
import hashlib
import pickle
import time

# Semua permintaan Google otomatis diulang (sampai 6x, jeda bertambah) kalau koneksi putus / error sementara.
# Permintaan yang SENGAJA diberi num_retries=0 (penyisipan baris, supaya tidak dobel) tidak ikut diulang.
try:
    import googleapiclient.http as _gh
    _orig_execute = _gh.HttpRequest.execute

    def _execute_retry(self, http=None, num_retries=None):
        return _orig_execute(self, http=http, num_retries=(6 if num_retries is None else num_retries))
    _gh.HttpRequest.execute = _execute_retry
except Exception:
    pass

# =========================== CONFIG ===========================

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"

# Tracker baru (file .xlsx di Drive)
TRACKER_FILE_ID = "1YRU3d0IKJYQhSQAVAWQHD9RHWGm09SY5RsWMTXHgaNY"        # [November] Daily Tracker NIVEA 2026 (Google Sheets asli)

# Folder Drive tempat file Creator_List_*.xlsx disimpan
DRIVE_FOLDER_ID: Optional[str] = "12Z93kKe9gNedQ-vQwjjYOoE7aMb9xGiL"

# Folder Drive file MCN: "Affiliate partnerships overview_Agency list_YYYY-MM-DD-YYYY-MM-DD.xlsx"
#   Agency name -> kolom Creator di tab MCN | GMV -> GMV Total | Videos -> SV | LIVE streams -> LS
#   (GMV LS / GMV SV / GMV SL tidak ada di file ini, jadi tidak diisi)
MCN_FOLDER_ID: Optional[str] = "1-HrGSphbeEhpYiViixVCzJp20-BhOHOr"   # folder "MCN Daily"
MCN_TAB = "MCN"
MCN_FILE_DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})-(\d{4}-\d{2}-\d{2})")

# Google Sheet sumber LS/SV: data harian mentah di tab "LS" dan tab "SV".
#   LS = jumlah Live ID UNIK per creator per tanggal
#   SV = jumlah Video ID UNIK per creator per tanggal
LS_SV_SPREADSHEET_ID = "1avlQXBatlbgT8xjYi8s_h5S4oiIwBN1A-6TnlySA1QQ"   # "Tracker Content Daily_November" (Google Sheets asli)
KNOWN_TAB_GIDS = {
    "LS": [926156041],
    "SV": [1947295267],
}

# Tab yang diisi sesuai daftar nama yang sudah ada
ROSTER_TABS = ["Celeb", "MCN", "TOP VIP", "RISING", "KLUB VIP", "Community"]

# Tab yang boleh ditambah creator baru + syarat minimumnya
COMMUNITY_TAB = "Community"
COMMUNITY_MIN_GMV = 0   # creator baru masuk Community kalau GMV Total hari itu > 0. Naikkan angkanya kalau mau batas lebih tinggi.

BACKUP_DIR = "backup_tracker"

# Folder penyimpanan sementara: file Creator_List / Agency yang SUDAH terunduh & terbaca disimpan di sini,
# jadi kalau koneksi putus dan skrip dijalankan lagi, file itu tidak diunduh ulang (aman dihapus kapan saja).
CACHE_DIR = "cache_tracker"
_TRACKER_MTIME = None      # waktu terakhir tracker diubah (dari Drive); selama tidak berubah, hasil baca tracker dipakai ulang

# Nama file harian
FILENAME_PATTERN = re.compile(r"^Creator_List_(\d{8})-(\d{8})(?:_.*)?$", re.IGNORECASE)

# Kolom file harian -> field internal
SOURCE_COLUMN_MAP = {
    "creator username": "name",
    "affiliate gmv": "total",
    "affiliate live gmv": "gmv_ls",
    "affiliate shoppable video gmv": "gmv_sv",
    "affiliate product card gmv": "gmv_sl",
}

# ================================================================


# ----------------------------- util -----------------------------

def _cache_file(key: str) -> str:
    os.makedirs(CACHE_DIR, exist_ok=True)
    return os.path.join(CACHE_DIR, hashlib.md5(key.encode("utf-8")).hexdigest() + ".pkl")


def cache_get(key: str, max_age: Optional[float] = None):
    path = _cache_file(key)
    try:
        if not os.path.exists(path):
            return None
        if max_age is not None and time.time() - os.path.getmtime(path) > max_age:
            return None
        with open(path, "rb") as fh:
            return pickle.load(fh)
    except Exception:
        return None


def cache_put(key: str, obj):
    try:
        path = _cache_file(key)
        tmp = path + ".tmp"
        with open(tmp, "wb") as fh:
            pickle.dump(obj, fh, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, path)
    except Exception:
        pass


def cache_cleanup(max_days: float = 3.0):
    try:
        if not os.path.isdir(CACHE_DIR):
            return
        now = time.time()
        for fn in os.listdir(CACHE_DIR):
            fp = os.path.join(CACHE_DIR, fn)
            if now - os.path.getmtime(fp) > max_days * 86400:
                os.remove(fp)
    except Exception:
        pass


def is_net_error(e) -> bool:
    """Error jaringan sementara (koneksi putus / timeout)."""
    if isinstance(e, (ConnectionError, TimeoutError, OSError)):
        return True
    return type(e).__name__ in ("ServerNotFoundError", "SSLError", "DecodeRatioError", "RemoteDisconnected", "IncompleteRead")


def norm_name(s) -> str:
    """Samakan penulisan nama: huruf kecil, buang spasi/titik/underscore/simbol."""
    return re.sub(r"[^0-9a-z]", "", str(s).lower()) if s is not None else ""


def exact_key(s) -> str:
    """Kunci pencocokan utama: huruf kecil + trim (titik/underscore TETAP dihitung, karena akun TikTok bisa beda hanya di situ)."""
    return str(s).strip().lower() if s is not None else ""


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def idx_to_col(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def to_number(x) -> float:
    """Angka dari sel. Angka asli dipakai apa adanya; teks dibaca dengan aturan pemisah ribuan/desimal yang aman."""
    if x is None or x == "":
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    t = str(x).strip()
    if not t:
        return 0.0
    if re.fullmatch(r"[1-9]\d{0,2}(\.\d{3})+", t):                    # "1.500", "12.345", "1.234.567" = pemisah ribuan (format Indonesia)
        return float(t.replace(".", ""))
    if re.fullmatch(r"[+-]?\d+(\.\d+)?([eE][+-]?\d+)?", t):          # "1234.5", "-12", "1.5E+06"
        return float(t)
    sign = -1.0 if (t.startswith("-") or t.startswith("\u2212") or (t.startswith("(") and t.endswith(")"))) else 1.0
    t = re.sub(r"(?i)rp|idr", "", t)
    t = re.sub(r"[^0-9.,]", "", t)
    if not t:
        return 0.0
    if "." in t and "," in t:                                          # pemisah yang muncul paling akhir = desimal
        dec = "." if t.rfind(".") > t.rfind(",") else ","
        thou = "," if dec == "." else "."
        t = t.replace(thou, "").replace(dec, ".")
    elif "," in t:
        if re.fullmatch(r"\d{1,3}(,\d{3})+", t) or t.count(",") > 1:
            t = t.replace(",", "")                                      # 1,234,567
        else:
            t = t.replace(",", ".")                                     # 12,5 -> 12.5
    elif "." in t:
        if re.fullmatch(r"[1-9]\d{0,2}(\.\d{3})+", t) or t.count(".") > 1:
            t = t.replace(".", "")                                      # 1.234.567 (format Indonesia)
    try:
        return sign * float(t)
    except ValueError:
        return 0.0


_SHEETS_EPOCH = dt.date(1899, 12, 30)


def parse_date_cell(v) -> Optional[dt.date]:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 30000 <= v <= 70000:  # rentang serial tanggal yang masuk akal (tahun ~1982-2091)
            return _SHEETS_EPOCH + dt.timedelta(days=int(round(v)))
        return None
    if isinstance(v, str):
        t = v.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%d-%b-%Y", "%d %b %Y", "%d %B %Y"):
            try:
                return dt.datetime.strptime(t, fmt).date()
            except ValueError:
                continue
    return None


# ------------------------- Google auth --------------------------

def get_services():
    scopes = [
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, "r") as f:
                stored = json.load(f)
            stored_scopes = set(stored.get("scopes") or str(stored.get("scope", "")).split())
        except Exception:
            stored_scopes = set()
        if stored_scopes >= set(scopes):
            creds = UserCredentials.from_authorized_user_file(TOKEN_FILE, scopes)
        else:
            print("  (Izin di token.json belum lengkap -- minta login ulang.)")

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"  (Gagal refresh token lama ({e}) -- minta login ulang.)")
                creds = None
        if not creds or not creds.valid:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(
                    f"File '{CREDENTIALS_FILE}' tidak ditemukan. Taruh di folder yang sama dengan script ini."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, scopes)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    try:    # batas waktu koneksi lebih longgar (default 60 detik sering habis untuk file besar / sinyal lemah)
        import httplib2
        import google_auth_httplib2
        mk = lambda: google_auth_httplib2.AuthorizedHttp(creds, http=httplib2.Http(timeout=300))
        return build("drive", "v3", http=mk()), build("sheets", "v4", http=mk())
    except ImportError:
        return build("drive", "v3", credentials=creds), build("sheets", "v4", credentials=creds)


# ------------------------- Drive helpers ------------------------

GSHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def download_bytes(drive_service, file_id: str, mime_type: Optional[str] = None) -> bytes:
    import time
    last = None
    for attempt in range(1, 6):
        try:
            if mime_type == GSHEET_MIME:      # file Google Sheets asli -> diekspor jadi .xlsx dulu
                request = drive_service.files().export_media(fileId=file_id, mimeType=XLSX_MIME)
            else:
                request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
            buf = io.BytesIO()
            downloader = MediaIoBaseDownload(buf, request, chunksize=8 * 1024 * 1024)
            done = False
            while not done:
                _, done = downloader.next_chunk(num_retries=3)
            return buf.getvalue()
        except Exception as e:
            status = getattr(getattr(e, "resp", None), "status", None)
            retryable = isinstance(e, (TimeoutError, ConnectionError, OSError)) or status in (429, 500, 502, 503, 504) \
                or "timed out" in str(e).lower()
            if not retryable or attempt == 5:
                raise
            last = e
            wait = 5 * attempt
            print(f"     (unduhan terputus: {type(e).__name__} -- mencoba lagi dalam {wait} detik, percobaan {attempt}/4)")
            time.sleep(wait)
    raise last


def upload_over(drive_service, file_id: str, local_path: str):
    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    drive_service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()


def _list_folder(drive_service, name_part: str, folder_id: Optional[str]) -> list:
    query = f"name contains '{name_part}' and trashed = false"
    if folder_id:
        query += f" and '{folder_id}' in parents"
    files, token = [], None
    while True:
        resp = drive_service.files().list(
            q=query, spaces="drive",
            fields="nextPageToken, files(id, name, mimeType, modifiedTime)",
            pageToken=token, pageSize=200,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        files.extend(resp.get("files", []))
        token = resp.get("nextPageToken")
        if not token:
            break
    return files


def list_candidate_files(drive_service) -> list:
    return _list_folder(drive_service, "Creator_List", DRIVE_FOLDER_ID)


def list_mcn_files(drive_service) -> list:
    return _list_folder(drive_service, "Agency list", MCN_FOLDER_ID)


def file_date(name: str) -> Optional[dt.date]:
    m = FILENAME_PATTERN.match(name)
    if not m:
        return None
    a, b = m.group(1), m.group(2)
    if a != b:
        return None
    try:
        return dt.datetime.strptime(a, "%Y%m%d").date()
    except ValueError:
        return None


def mcn_file_date(name: str) -> Optional[dt.date]:
    m = MCN_FILE_DATE_RE.search(name)
    if not m or m.group(1) != m.group(2):
        return None
    try:
        return dt.datetime.strptime(m.group(1), "%Y-%m-%d").date()
    except ValueError:
        return None


def entry_date(f: dict) -> Optional[dt.date]:
    return mcn_file_date(f["name"]) if f.get("kind") == "mcn" else file_date(f["name"])


# Nama file harian: Creator_List_<tgl>-<tgl>_<CAP WAKTU EXPORT>_<jenis laporan/toko>.xlsx
EXPORT_NAME_RE = re.compile(r"^Creator_List_(\d{8})-(\d{8})_(\d{6,14})_(.*?)(?:\.xlsx)?$", re.IGNORECASE)
PAKAI_EXPORT_TERBARU_SAJA = True   # False = semua file dijumlahkan walau tanggal & laporannya sama


def dedupe_exports(todo: list):
    """
    todo = [(tanggal, file)]. Kalau ada beberapa file untuk TANGGAL dan JENIS LAPORAN yang sama (beda cap waktu export
    saja) itu laporan yang sama diunduh ulang -> hanya yang terbaru dipakai, supaya tidak terjumlah dobel.
    Jenis laporan berbeda (akhiran nama beda, mis. toko lain) tetap dijumlahkan.
    Return (dipakai, diabaikan).
    """
    if not PAKAI_EXPORT_TERBARU_SAJA:
        return list(todo), []
    groups = {}
    for d, f in todo:
        m = EXPORT_NAME_RE.match(f["name"])
        suffix, ts = (m.group(4).lower(), int(m.group(3))) if m else (f["name"].lower(), 0)
        groups.setdefault((d, suffix), []).append((ts, d, f))
    kept, dropped = [], []
    for items in groups.values():
        items.sort(key=lambda x: x[0])
        kept.append((items[-1][1], items[-1][2]))
        dropped += [(x[1], x[2]) for x in items[:-1]]
    kept.sort(key=lambda t: (t[0], t[1]["name"]))
    return kept, dropped


def _rp(v) -> str:
    return f"{v:,.0f}"


def _cut(text, n) -> str:
    text = str(text)
    return text if len(text) <= n else text[:n - 1] + "~"


def print_preview(src_by_date: dict, mcn_by_date: dict, st: dict, top: int = 15):
    """Tampilkan isi file (GMV Total / LS / SV / SL + jumlah LS & SV) dan tujuan penulisannya -- belum ada yang ditulis."""
    print("\n" + "=" * 20 + " PRATINJAU ISI FILE (belum ditulis ke tracker) " + "=" * 20)
    for date in sorted(src_by_date):
        pv = st.get("preview", {}).get(date)
        src = src_by_date[date]
        tot = {f: sum(r[f] for r in src.values()) for f in ("total", "gmv_ls", "gmv_sv", "gmv_sl")}
        print(f"\n[{date:%d-%b-%Y}]  Creator_List: {len(src)} creator di file")
        print(f"   GMV Total Rp {_rp(tot['total']):>16}   GMV LS Rp {_rp(tot['gmv_ls']):>16}")
        print(f"   GMV SV    Rp {_rp(tot['gmv_sv']):>16}   GMV SL Rp {_rp(tot['gmv_sl']):>16}")
        if not pv:
            print("   (tanggal ini tidak akan ditulis: di luar bulan tracker atau blok tanggalnya tidak ada)")
            continue
        rows, skipped = pv["rows"], pv["skipped"]
        ls_sum = sum(r["ls"] for r in rows if r["ls"] is not None)
        sv_sum = sum(r["sv"] for r in rows if r["sv"] is not None)
        has_ls = any(r["ls"] is not None for r in rows)
        has_sv = any(r["sv"] is not None for r in rows)
        print(f"   LS (sesi live) dari Google Sheet: {ls_sum if has_ls else '-'}   |   SV (video): {sv_sum if has_sv else '-'}   (hanya creator yang akan ditulis)")
        w = sum(r["total"] for r in rows)
        in_file = [r for r in rows if r.get("in_file", True)]
        only_cnt = [r for r in rows if not r.get("in_file", True)]
        print(f"   Akan ditulis : {len(in_file)} creator dari file, GMV Total Rp {_rp(w)}"
              + (f"  (+ {len(only_cnt)} creator yang hanya punya LS/SV)" if only_cnt else ""))
        print(f"   Tidak ditulis: {len(skipped)} creator, GMV Total Rp {_rp(sum(r['total'] for r in skipped))}")
        rows = in_file
        print(f"   {min(top, len(rows))} GMV terbesar yang akan ditulis:")
        print(f"     {'Creator':<24} {'Tab':<9} {'GMV Total':>14} {'GMV LS':>13} {'GMV SV':>13} {'GMV SL':>12} {'LS':>3} {'SV':>3}  Status")
        for r in sorted(rows, key=lambda x: -x["total"])[:top]:
            print(f"     {_cut(r['name'], 24):<24} {r['tab']:<9} {_rp(r['total']):>14} {_rp(r['gmv_ls']):>13} {_rp(r['gmv_sv']):>13} "
                  f"{_rp(r['gmv_sl']):>12} {('-' if r['ls'] is None else r['ls']):>3} {('-' if r['sv'] is None else r['sv']):>3}  {r['status']}")
    for date in sorted(mcn_by_date):
        data = mcn_by_date[date]
        print(f"\n[{date:%d-%b-%Y}]  MCN / Agency list: {len(data)} agency di file")
        print(f"   GMV Total Rp {_rp(sum(r['gmv'] for r in data.values())):>16}   (GMV LS / GMV SV / GMV SL tidak ada di file ini)")
        print(f"     {'Agency':<26} {'GMV Total':>15} {'LS':>5} {'SV':>5}  Status")
        for r in sorted(st.get("mcn", {}).get("preview", {}).get(date, []), key=lambda x: -x["gmv"]):
            print(f"     {_cut(r['name'], 26):<26} {_rp(r['gmv']):>15} {('-' if r['ls'] is None else int(r['ls'])):>5} "
                  f"{('-' if r['sv'] is None else int(r['sv'])):>5}  {r['status']}")
    print("=" * 87)


def save_preview_xlsx(src_by_date: dict, mcn_by_date: dict, st: dict, path: str):
    """Semua baris (yang ditulis & yang dilewati) ke file Excel biasa, supaya bisa dicek/difilter sebelum konfirmasi."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Creator_List"
    ws.append(["Tanggal", "Creator", "Tab tujuan", "Baris di tab", "GMV Total", "GMV LS", "GMV SV", "GMV SL", "LS (sesi)", "SV (video)", "Status"])
    for date in sorted(st.get("preview", {})):
        pv = st["preview"][date]
        for r in sorted(pv["rows"], key=lambda x: (not x.get("in_file", True), -x["total"])) + sorted(pv["skipped"], key=lambda x: -x["total"]):
            ws.append([date, r["name"], r["tab"], r["row"], r["total"], r["gmv_ls"], r["gmv_sv"], r["gmv_sl"], r["ls"], r["sv"], r["status"]])
    for row in ws.iter_rows(min_row=2):
        row[0].number_format = "dd-mmm-yyyy"
        for c in row[4:8]:
            c.number_format = "#,##0"
    wm = wb.create_sheet("MCN")
    wm.append(["Tanggal", "Agency", "GMV Total", "LS (LIVE streams)", "SV (Videos)", "Status"])
    for date in sorted(st.get("mcn", {}).get("preview", {})):
        for r in sorted(st["mcn"]["preview"][date], key=lambda x: -x["gmv"]):
            wm.append([date, r["name"], r["gmv"], r["ls"], r["sv"], r["status"]])
    for row in wm.iter_rows(min_row=2):
        row[0].number_format = "dd-mmm-yyyy"
        row[2].number_format = "#,##0"
    wr = wb.create_sheet("Ringkasan")
    wr.append(["Sumber", "Tanggal", "GMV Total", "GMV LS", "GMV SV", "GMV SL", "Jumlah baris"])
    for date in sorted(src_by_date):
        src = src_by_date[date]
        wr.append(["Creator_List", date] + [sum(r[f] for r in src.values()) for f in ("total", "gmv_ls", "gmv_sv", "gmv_sl")] + [len(src)])
    for date in sorted(mcn_by_date):
        d = mcn_by_date[date]
        wr.append(["MCN (Agency list)", date, sum(r["gmv"] for r in d.values()), None, None, None, len(d)])
    for row in wr.iter_rows(min_row=2):
        row[1].number_format = "dd-mmm-yyyy"
        for c in row[2:6]:
            c.number_format = "#,##0"
    for sheet in wb.worksheets:
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = 22 if col[0].column > 1 else 14
    wb.save(path)


def describe_creator_stats(st: dict, tot: float) -> list:
    """Catatan yang menjelaskan kenapa jumlah per creator bisa beda dari jumlah mentah kolom GMV di file."""
    rp = lambda v: f"{v:,.0f}"
    notes = []
    if st.get("total_rows"):
        notes.append(f"{st['total_rows']} baris Total dilewati (Rp {rp(st['total_rows_sum'])})")
    if st.get("blank_rows"):
        notes.append(f"{st['blank_rows']} baris TANPA username (Rp {rp(st['blank_total'])}) digabung jadi satu baris '{NO_NAME_LABEL}' "
                     f"-- ikut dihitung, supaya total sama dengan file")
    if st.get("bad_cells"):
        notes.append(f"{st['bad_cells']} sel GMV berupa teks yang tidak terbaca (dianggap 0), contoh: "
                     + ", ".join(repr(x) for x in st["bad_samples"]))
    for title, n, t in st.get("other_sheets", []):
        notes.append(f"sheet lain '{title}' juga berisi tabel ({n:,} creator, Rp {rp(t)}) -- TIDAK dihitung")
    unexplained = st.get("all_raw", 0.0) - st.get("total_rows_sum", 0.0) - tot
    if abs(unexplained) > 1:
        notes.append(f"selisih Rp {rp(unexplained)} belum terjelaskan")
    return notes


_PARSE_CACHE = {}


def _file_cache_key(kind: str, f: dict) -> str:
    return f"{kind}|{f['id']}|{f.get('modifiedTime', '')}"


def read_creator_file(drive_service, f):
    """Baca file Creator_List (sekali saja; hasil disimpan sementara di memori DAN di disk). Return (data, statistik_mentah)."""
    key = ("creator", f["id"])
    if key not in _PARSE_CACHE:
        ck = _file_cache_key("creator", f)
        hit = cache_get(ck)
        if hit is not None:
            print("     (dari penyimpanan sementara -- tidak diunduh ulang)")
            _PARSE_CACHE[key] = hit
        else:
            stats = _new_creator_stats()
            data = parse_creator_list(download_bytes(drive_service, f["id"], f.get("mimeType")), stats)
            _PARSE_CACHE[key] = (data, stats)
            cache_put(ck, _PARSE_CACHE[key])
    return _PARSE_CACHE[key]


def read_mcn_file(drive_service, f):
    """Baca file Agency list (sekali saja). Return (data, catatan)."""
    key = ("mcn", f["id"])
    if key not in _PARSE_CACHE:
        ck = _file_cache_key("mcn", f)
        hit = cache_get(ck)
        if hit is not None:
            print("     (dari penyimpanan sementara -- tidak diunduh ulang)")
            _PARSE_CACHE[key] = hit
        else:
            _PARSE_CACHE[key] = parse_agency_list(download_bytes(drive_service, f["id"], f.get("mimeType")))
            cache_put(ck, _PARSE_CACHE[key])
    return _PARSE_CACHE[key]


def _sorted_entries(entries: list) -> list:
    return sorted(entries, key=lambda f: (f.get("kind") != "creator", entry_date(f) or dt.date.max, f["name"]))


def print_gmv_preview(drive_service, entries: list):
    """
    Tampilkan total GMV tiap file di folder (tanpa menyentuh tracker) supaya bisa dicocokkan dengan hitunganmu sendiri.
    Nomor sama dengan nomor di daftar pilihan.
    """
    entries = _sorted_entries(entries)
    rp = lambda v: f"{v:,.0f}"
    creators = [(i, f) for i, f in enumerate(entries, 1) if f.get("kind") == "creator"]
    mcns = [(i, f) for i, f in enumerate(entries, 1) if f.get("kind") == "mcn"]

    if creators:
        _, ignored = dedupe_exports([(entry_date(f), f) for _, f in creators if entry_date(f)])
        ignored_ids = {f["id"] for _, f in ignored}
        print("\nTOTAL GMV PER FILE -- Creator_List (Rp)")
        print(f"  {'No':>3}  {'Tanggal':<11} {'Creator':>8} {'GMV Total':>16} {'GMV LS':>15} {'GMV SV':>15} {'GMV SL':>15}  Catatan")
        used = {"total": 0.0, "ls": 0.0, "sv": 0.0, "sl": 0.0}
        for n, (i, f) in enumerate(creators, 1):
            d = entry_date(f)
            ds = d.strftime("%d-%b-%Y") if d else "(?)"
            print(f"  ...membaca {n}/{len(creators)}", end="\r", flush=True)
            try:
                data, st = read_creator_file(drive_service, f)
            except Exception as e:
                print(f"  {i:>3}  {ds:<11} [ERROR] {_short_err(e)}" + " " * 20)
                continue
            tot = sum(r["total"] for r in data.values())
            ls_, sv_, sl_ = (sum(r[k] for r in data.values()) for k in ("gmv_ls", "gmv_sv", "gmv_sl"))
            notes = describe_creator_stats(st, tot)
            if f["id"] in ignored_ids:
                notes.append("DIABAIKAN: ada export lebih baru di tanggal yang sama")
            else:
                for k, v in (("total", tot), ("ls", ls_), ("sv", sv_), ("sl", sl_)):
                    used[k] += v
            print(f"  {i:>3}  {ds:<11} {len(data):>8,} {rp(tot):>16} {rp(ls_):>15} {rp(sv_):>15} {rp(sl_):>15}  " + "; ".join(notes) + " " * 12)
        print(f"  {'':>3}  {'JUMLAH (file yang akan dipakai)':<20} {rp(used['total']):>16} {rp(used['ls']):>15} {rp(used['sv']):>15} {rp(used['sl']):>15}")

    if mcns:
        print("\nTOTAL GMV PER FILE -- MCN / Agency list (Rp)")
        print(f"  {'No':>3}  {'Tanggal':<11} {'Agency':>8} {'GMV':>16} {'LS':>9} {'SV':>9}  Catatan")
        tot_all = 0.0
        for n, (i, f) in enumerate(mcns, 1):
            d = entry_date(f)
            ds = d.strftime("%d-%b-%Y") if d else "(?)"
            print(f"  ...membaca {n}/{len(mcns)}", end="\r", flush=True)
            try:
                data, notes = read_mcn_file(drive_service, f)
            except Exception as e:
                print(f"  {i:>3}  {ds:<11} [ERROR] {_short_err(e).splitlines()[0] if _short_err(e) else ''}" + " " * 20)
                continue
            tot = sum(r["gmv"] for r in data.values())
            tot_all += tot
            lsn = sum(r["ls"] or 0 for r in data.values()) if any(r["ls"] is not None for r in data.values()) else None
            svn = sum(r["sv"] or 0 for r in data.values()) if any(r["sv"] is not None for r in data.values()) else None
            print(f"  {i:>3}  {ds:<11} {len(data):>8,} {rp(tot):>16} {'-' if lsn is None else rp(lsn):>9} {'-' if svn is None else rp(svn):>9}  " + "; ".join(notes) + " " * 12)
        print(f"  {'':>3}  {'JUMLAH':<20} {rp(tot_all):>16}")
        print("  (GMV agency umumnya sudah mencakup GMV creator di bawahnya -- jangan dijumlahkan dengan Creator_List.)")
    print()


def prompt_user_selection(entries: list, on_preview=None) -> list:
    """entries = file dengan kunci 'kind' ('creator' / 'mcn'). Nomor urut berlanjut antar bagian."""
    entries = _sorted_entries(entries)
    print()
    for kind, title in (("creator", "File Creator_List (folder Daily CP) -> tab Celeb, TOP VIP, RISING, KLUB VIP, Community"),
                        ("mcn", "File MCN / Agency list (folder MCN Daily) -> tab MCN")):
        group = [f for f in entries if f.get("kind") == kind]
        if not group:
            continue
        print(title + ":")
        for f in group:
            d = entry_date(f)
            flag = d.strftime("%d-%b-%Y") if d else "!! (nama/tanggal tidak dikenali, akan di-skip)"
            print(f"  {entries.index(f) + 1:>3}. [{flag}]  {f['name']}")
        print()
    while True:
        print("Pilih nomor file yang mau diisi ke tracker.")
        print("Contoh: 1,3,5  atau  2-6  atau  semua  |  'lihat' = tampilkan total GMV tiap file dulu  |  kosong / 'batal' = keluar")
        raw = input("> ").strip().lower()
        if raw in ("lihat", "preview", "cek") and on_preview is not None:
            on_preview(entries)
            continue
        break
    if raw in ("", "batal", "cancel"):
        return []
    if raw in ("semua", "all"):
        return entries
    chosen = []
    for part in raw.split(","):
        part = part.strip()
        m = re.fullmatch(r"(\d+)\s*-\s*(\d+)", part)
        if m:
            lo, hi = int(m.group(1)), int(m.group(2))
            idxs = range(min(lo, hi), max(lo, hi) + 1)
        elif part.isdigit():
            idxs = [int(part)]
        else:
            print(f"  (dilewati: '{part}')")
            continue
        for idx in idxs:
            if 1 <= idx <= len(entries):
                if entries[idx - 1] not in chosen:
                    chosen.append(entries[idx - 1])
            else:
                print(f"  (nomor {idx} di luar daftar)")
    return chosen


# ---------------------- baca file Creator_List ------------------

def _classify_header(h):
    """Kenali judul kolom file Creator_List (longgar: abaikan huruf besar/kecil, tanda baca, satuan seperti '(Rp)')."""
    t = re.sub(r"[^a-z0-9]+", " ", str(h).lower()).strip()
    if not t:
        return None
    if t in ("creator username", "creator user name", "username", "user name") or \
            (t.startswith("creator") and "username" in t.replace(" ", "")):
        return "name"
    if "gmv" not in t:
        return None
    if "shoppable video" in t or "short video" in t:
        return "gmv_sv"
    if "product card" in t:
        return "gmv_sl"
    if "live" in t:
        return "gmv_ls"
    if re.fullmatch(r"affiliate gmv( rp| idr)?", t):
        return "total"
    return None


NO_NAME_LABEL = "(Tanpa Username)"


def _new_creator_stats() -> dict:
    return {"rows": 0, "raw_total": 0.0, "total_rows": 0, "total_rows_sum": 0.0, "all_raw": 0.0,
            "blank_rows": 0, "blank_total": 0.0, "bad_cells": 0, "bad_samples": [], "other_sheets": []}


def _scan_creator_rows(rows, stats=None):
    """Return (result, header_ditemukan, kolom_terdeteksi, beberapa_baris_pertama)."""
    header_map, started, result, first_rows = {}, False, {}, []
    for row in rows:
        if len(first_rows) < 6:
            first_rows.append(row)
        if not started:
            hits = {}
            for i, v in enumerate(row):
                if v is None:
                    continue
                k = _classify_header(v)
                if k and k not in hits.values():
                    hits[i] = k
            if "name" in hits.values() and "total" in hits.values():
                header_map, started = hits, True
            continue
        rec = {field: row[i] if i < len(row) else None for i, field in header_map.items()}
        name = rec.get("name")
        raw_g = rec.get("total")
        g = to_number(raw_g)
        if stats is not None:
            stats["all_raw"] += g                                    # jumlah mentah kolom GMV, SEMUA baris data
            if isinstance(raw_g, str) and raw_g.strip() not in ("", "-", "0", "0.0") and g == 0.0:
                stats["bad_cells"] += 1
                if len(stats["bad_samples"]) < 3:
                    stats["bad_samples"].append(raw_g.strip()[:20])
        key = exact_key(str(name).strip()) if name is not None else ""
        if not key:
            # Tidak ada username: GMV-nya TIDAK dibuang. Semua baris seperti ini digabung jadi satu "creator"
            # bernama "(Tanpa Username)" supaya total yang dibaca sama dengan total di file.
            vals = {f: to_number(rec.get(f)) for f in ("total", "gmv_ls", "gmv_sv", "gmv_sl")}
            if any(vals.values()):
                nk = exact_key(NO_NAME_LABEL)
                cur = result.setdefault(nk, {"name": NO_NAME_LABEL, "total": 0.0, "gmv_ls": 0.0, "gmv_sv": 0.0, "gmv_sl": 0.0})
                for f_, v_ in vals.items():
                    cur[f_] += v_
                if stats is not None:
                    stats["blank_rows"] += 1
                    stats["blank_total"] += g
            continue
        name = str(name).strip()
        if stats is not None:
            stats["rows"] += 1
            stats["raw_total"] += g
        if _norm_sheet(name) in ("total", "grand total", "subtotal", "sub total", "jumlah", "total keseluruhan", "sum"):
            if stats is not None:
                stats["total_rows"] += 1
                stats["total_rows_sum"] += g
            else:
                print(f"  (baris '{name}' dilewati: itu baris jumlah, bukan creator)")
            continue
        cur = result.setdefault(key, {"name": name, "total": 0.0, "gmv_ls": 0.0, "gmv_sv": 0.0, "gmv_sl": 0.0})
        for f in ("total", "gmv_ls", "gmv_sv", "gmv_sl"):
            cur[f] += to_number(rec.get(f))
    return result, started, set(header_map.values()), first_rows


def parse_creator_list(raw_bytes: bytes, stats: Optional[dict] = None) -> dict:
    """Return { exact_key: {"name","total","gmv_ls","gmv_sv","gmv_sl"} } (angka digabung kalau nama dobel)."""
    need = {"name", "total", "gmv_ls", "gmv_sv", "gmv_sl"}
    diag, sheet_names = [], []
    for read_only in (True, False):          # cara 1: cepat; cara 2: tanpa read-only (untuk file yang dimensinya tidak tercatat)
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=read_only)
        sheet_names = list(wb.sheetnames)
        found, others = None, []
        for ws in wb.worksheets:
            if read_only:
                try:
                    ws.reset_dimensions()
                except Exception:
                    pass
            sheet_stats = _new_creator_stats()
            result, started, cols, first_rows = _scan_creator_rows(ws.iter_rows(values_only=True), sheet_stats)
            if started:
                if found is None:                                    # sheet pertama yang berisi tabel = yang dipakai
                    missing = need - cols
                    if missing:
                        raise ValueError(
                            "Kolom berikut tidak ketemu di file ini: "
                            + ", ".join(sorted({"name": "Creator username", "total": "Affiliate GMV",
                                                "gmv_ls": "Affiliate LIVE GMV", "gmv_sv": "Affiliate shoppable video GMV",
                                                "gmv_sl": "Affiliate product card GMV"}[m] for m in missing)))
                    found = (result, sheet_stats)
                else:                                                # sheet lain yang juga berisi tabel -> hanya dilaporkan
                    others.append((ws.title, len(result), sum(r["total"] for r in result.values())))
                continue
            if not any(t == ws.title for t, _ in diag) or not read_only:
                diag = [(t, r) for t, r in diag if t != ws.title] + [(ws.title, first_rows)]
        wb.close()
        if found is not None:
            if stats is not None:
                stats.update(found[1])
                stats["other_sheets"] = others
            return found[0]

    lines = [f"Header kolom (Creator username, Affiliate GMV, ...) tidak ditemukan. Sheet di file: {sheet_names}."]
    for title, rows in diag[:2]:
        lines.append(f"  Isi awal sheet '{title}':")
        if not rows:
            lines.append("    (kosong)")
        for r in rows:
            lines.append("    " + " | ".join(str(c)[:22] for c in list(r)[:10] if c is not None))
    raise ValueError("\n".join(lines))


# ---------------------- baca file Agency list (MCN) -------------

def _classify_agency_header(h):
    t = re.sub(r"[^a-z0-9]+", " ", str(h).lower()).strip()
    if t in ("agency name", "agency"):
        return "name"
    if re.fullmatch(r"gmv( rp| idr)?", t):
        return "gmv"
    if re.fullmatch(r"live ?streams?( count)?|lives?", t):
        return "ls"
    if re.fullmatch(r"videos?( count)?", t):
        return "sv"
    return None


def _scan_agency_rows(rows):
    """Return (result, header_ditemukan, kolom_terdeteksi, beberapa_baris_pertama)."""
    header_map, started, result, first_rows = {}, False, {}, []
    for row in rows:
        if len(first_rows) < 6:
            first_rows.append(row)
        if not started:
            hits = {}
            for i, v in enumerate(row):
                if v is None:
                    continue
                k = _classify_agency_header(v)
                if k and k not in hits.values():
                    hits[i] = k
            if "name" in hits.values() and "gmv" in hits.values():
                header_map, started = hits, True
            continue
        rec = {field: row[i] if i < len(row) else None for i, field in header_map.items()}
        name = rec.get("name")
        if name is None or str(name).strip() == "":
            if to_number(rec.get("gmv")):
                name = "(Tanpa Agency)"
            else:
                continue
        key = exact_key(name)
        if not key or key == "total":
            continue
        cur = result.setdefault(key, {"name": str(name).strip(), "gmv": 0.0,
                                      "ls": 0.0 if "ls" in header_map.values() else None,
                                      "sv": 0.0 if "sv" in header_map.values() else None})
        cur["gmv"] += to_number(rec.get("gmv"))
        for f in ("ls", "sv"):
            if cur[f] is not None:
                cur[f] += to_number(rec.get(f))
    return result, started, set(header_map.values()), first_rows


def _count_by_agency(rows, id_hint):
    """Cadangan: sheet 'Videos' / 'LIVE streams' berisi satu baris per video/live -> hitung per agency (ID unik kalau ada)."""
    for ri, row in enumerate(rows[:15]):
        heads = [re.sub(r"[^a-z0-9]+", " ", str(c).lower()).strip() if c is not None else "" for c in row]
        if "agency name" in heads:
            ni = heads.index("agency name")
            ii = next((j for j, h in enumerate(heads) if h.endswith(" id") and id_hint in h), None)
            seen, counts = {}, {}
            for r in rows[ri + 1:]:
                if ni >= len(r) or r[ni] in (None, ""):
                    continue
                k = exact_key(r[ni])
                if ii is not None and ii < len(r) and r[ii] not in (None, ""):
                    seen.setdefault(k, set()).add(_id_key(r[ii]))
                else:
                    counts[k] = counts.get(k, 0) + 1
            for k, st in seen.items():
                counts[k] = counts.get(k, 0) + len(st)
            return counts
    return None


def parse_agency_list(raw_bytes: bytes):
    """
    Return (data, notes)
      data = { exact_key: {"name", "gmv", "ls" (atau None), "sv" (atau None)} }
    LS/SV diambil dari kolom 'LIVE streams' / 'Videos'. Kalau kolomnya tidak ada, dicoba dari sheet bernama
    'LIVE streams' / 'Videos' (satu baris per live/video). Kalau tetap tidak ada -> None (kolom tracker tidak disentuh).
    """
    notes, diag, sheet_names = [], [], []
    for read_only in (True, False):
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=read_only)
        sheet_names = list(wb.sheetnames)
        all_rows = {}
        for ws in wb.worksheets:
            if read_only:
                try:
                    ws.reset_dimensions()
                except Exception:
                    pass
            all_rows[ws.title] = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()

        for title, rows in all_rows.items():
            result, started, cols, first_rows = _scan_agency_rows(rows)
            if not started:
                diag = [(t, r) for t, r in diag if t != title] + [(title, first_rows)]
                continue
            if "ls" not in cols or "sv" not in cols:
                for field, names, hint, label in (("ls", ("live streams", "live stream", "lives", "live"), "live", "LIVE streams"),
                                                  ("sv", ("videos", "video"), "video", "Videos")):
                    if field in cols:
                        continue
                    other = next((t for t in all_rows if re.sub(r"[^a-z0-9]+", " ", t.lower()).strip() in names), None)
                    counts = _count_by_agency(all_rows[other], hint) if other else None
                    if counts is None:
                        notes.append(f"Kolom/sheet '{label}' tidak ketemu di file ini -> {field.upper()} tidak diisi. "
                                     f"(sheet di file: {sheet_names})")
                        continue
                    for k, rec in result.items():
                        rec[field] = float(counts.get(k, 0))
                    notes.append(f"{field.upper()} dihitung dari sheet '{other}' (jumlah baris/ID unik per agency).")
            return result, notes

    lines = [f"Header 'Agency name' + 'GMV' tidak ditemukan. Sheet di file: {sheet_names}."]
    for title, rows in diag[:3]:
        lines.append(f"  Isi awal sheet '{title}':")
        if not rows:
            lines.append("    (kosong)")
        for r in rows:
            lines.append("    " + " | ".join(str(c)[:22] for c in list(r)[:10] if c is not None))
    raise ValueError("\n".join(lines))


# ------------------ baca LS / SV dari Google Sheet --------------

def list_tabs(sheets_service, spreadsheet_id: str) -> list:
    meta = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties(sheetId,title)"
    ).execute()
    return [(s["properties"]["title"], s["properties"]["sheetId"]) for s in meta.get("sheets", [])]


def fetch_tab_values(sheets_service, spreadsheet_id: str, title: str) -> list:
    ck = f"tab|{spreadsheet_id}|{title}"
    hit = cache_get(ck, max_age=30 * 60)          # data sumber LS/SV: dipakai ulang max 30 menit kalau skrip diulang
    if hit is not None:
        print(f"     (tab '{title}' dari penyimpanan sementara)")
        return hit
    safe = title.replace("'", "''")
    resp = _with_retry(lambda: sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"'{safe}'!A:AZ",
        valueRenderOption="UNFORMATTED_VALUE",
    ), f"membaca tab '{title}'")
    vals = resp.get("values", [])
    cache_put(ck, vals)
    return vals


def _cell(row, i):
    return row[i] if i < len(row) else None


def _id_key(v) -> str:
    """ID panjang (Live ID / Video ID) -> teks. Angka dibulatkan tanpa notasi ilmiah, teks apa adanya."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip().lstrip("'")


def extract_counts(rows: list, kind: str, info: Optional[dict] = None):
    """
    Kenali bentuk tab lalu hitung jumlah per (creator, tanggal).
    Return (layout, counts, covered):
       counts  = { date: { exact_key(creator): jumlah } }
       covered = fungsi(date) -> True kalau tab ini memang memuat data untuk tanggal tsb
    layout: 'mentah' (satu baris = satu data harian; dihitung ID UNIK), 'matriks', 'periode', atau None.
    info (opsional, diisi): {"no_id_rows": n, "uniq": True/False, "last_date": date}
    """
    info = info if info is not None else {}
    lowered = [[str(c).strip().lower() if c is not None else "" for c in r] for r in rows[:25]]
    id_header = {"LS": "live id", "SV": "video id"}.get(kind.upper())

    # 1) tab mentah: ada header 'date/tanggal' dan 'creator'
    for ri, row in enumerate(lowered[:12]):
        if ("creator" in row) and ("date" in row or "tanggal" in row):
            ci = row.index("creator")
            di = row.index("date") if "date" in row else row.index("tanggal")
            ii = row.index(id_header) if (id_header and id_header in row) else None
            ids, counts, dates_seen, no_id = {}, {}, [], 0
            for r in rows[ri + 1:]:
                d = parse_date_cell(_cell(r, di))
                nm = exact_key(_cell(r, ci))
                if not (d and nm):
                    continue
                dates_seen.append(d)
                if ii is None:                       # tidak ada kolom ID -> terpaksa hitung baris
                    counts.setdefault(d, {})
                    counts[d][nm] = counts[d].get(nm, 0) + 1
                    continue
                idv = _id_key(_cell(r, ii))
                if not idv:
                    no_id += 1
                    continue
                ids.setdefault((d, nm), set()).add(idv)
            for (d, nm), idset in ids.items():
                counts.setdefault(d, {})[nm] = len(idset)
            info.update({"no_id_rows": no_id, "uniq": ii is not None,
                         "last_date": max(dates_seen) if dates_seen else None})
            hi = max(dates_seen) if dates_seen else None
            return "mentah", counts, (lambda d, hi=hi: hi is not None and d <= hi)

    # 2) matriks: header 'creator' + >=3 sel tanggal pada baris yang sama
    for ri, r in enumerate(rows[:15]):
        low = lowered[ri] if ri < len(lowered) else []
        if "creator" not in low:
            continue
        ci = low.index("creator")
        date_cols = {i: parse_date_cell(v) for i, v in enumerate(r) if i != ci and parse_date_cell(v)}
        if len(date_cols) >= 3:
            counts = {}
            for rr in rows[ri + 1:]:
                nm = exact_key(_cell(rr, ci))
                if not nm:
                    continue
                for i, d in date_cols.items():
                    v = _cell(rr, i)
                    if isinstance(v, (int, float)) and v > 0:
                        counts.setdefault(d, {})
                        counts[d][nm] = counts[d].get(nm, 0) + int(round(v))
            covered_dates = set(date_cols.values())
            return "matriks", counts, (lambda d, cd=covered_dates: d in cd)

    # 3) summary periode: header 'creator' + kolom 'jumlah ...' dan Start/End Date di atasnya
    for ri, low in enumerate(lowered[:15]):
        if "creator" not in low:
            continue
        ci = low.index("creator")
        cnt_idx = None
        for i, h in enumerate(low):
            if h.startswith("jumlah") and (kind.lower() in h or "session" in h or "video" in h):
                cnt_idx = i
                break
        if cnt_idx is None:
            continue
        start = end = None
        for rr in rows[:ri]:
            for i, v in enumerate(rr):
                t = str(v).strip().lower() if v is not None else ""
                if t == "start date":
                    start = parse_date_cell(_cell(rr, i + 1))
                if t == "end date":
                    end = parse_date_cell(_cell(rr, i + 1))
        if start and end and start == end:
            counts = {start: {}}
            for rr in rows[ri + 1:]:
                nm = exact_key(_cell(rr, ci))
                v = _cell(rr, cnt_idx)
                if nm and isinstance(v, (int, float)) and v > 0:
                    counts[start][nm] = int(round(v))
            return "periode", counts, (lambda d, s=start: d == s)
        return "periode-banyak-hari", {}, (lambda d: False)

    return None, {}, (lambda d: False)


def _short_err(e) -> str:
    return " ".join(str(e).split())[:300]


def load_tabs_via_drive(drive_service, file_id: str) -> dict:
    """
    Cadangan kalau Sheets API menolak file (mis. file Excel .xlsx yang di-upload ke Drive):
    unduh file lewat Drive, lalu baca tab 'LS' dan 'SV'. Return {judul_tab: baris-baris}.
    """
    meta = drive_service.files().get(fileId=file_id, fields="name,mimeType", supportsAllDrives=True).execute()
    print(f"  Membaca '{meta.get('name')}' lewat Drive (tipe: {meta.get('mimeType')}) ...")
    data = download_bytes(drive_service, file_id, meta.get("mimeType"))
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    out = {}
    for ws in wb.worksheets:
        if ws.title.strip().lower() in ("ls", "sv"):
            try:
                ws.reset_dimensions()
            except Exception:
                pass
            out[ws.title] = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    if not out:
        raise RuntimeError(f"Tab 'LS' / 'SV' tidak ada di file itu. Tab yang ada: {list(wb.sheetnames)}")
    return out


def load_ls_sv_counts(sheets_service, needed_dates: set, drive_service=None):
    """
    Return (counts, notes)
      counts = {"LS": {date: {norm: n}}, "SV": {...}}  (hanya untuk tanggal yang berhasil dicakup)
    """
    drive_rows = None
    try:
        tabs = list_tabs(sheets_service, LS_SV_SPREADSHEET_ID)
        fetch = lambda title: fetch_tab_values(sheets_service, LS_SV_SPREADSHEET_ID, title)
        print("  Tab di Google Sheet LS/SV:")
    except Exception as e:
        if drive_service is None:
            raise
        print(f"  Sheets API tidak bisa membuka file ini ({_short_err(e)})")
        print("  -> mencoba lewat Drive (unduh sebagai file Excel) ...")
        drive_rows = load_tabs_via_drive(drive_service, LS_SV_SPREADSHEET_ID)
        tabs = [(t, None) for t in drive_rows]
        fetch = lambda title: drive_rows[title]
        print("  Tab yang terbaca:")
    for title, gid in tabs:
        print(f"     - '{title}'" + (f" (gid={gid})" if gid is not None else ""))

    out = {"LS": {}, "SV": {}}
    notes = []
    for kind in ("LS", "SV"):
        kind_re = re.compile(rf"(?<![a-z]){kind}(?![a-z])", re.I)
        cands = [(title, gid) for title, gid in tabs if title.strip().lower() == kind.lower()]
        if not cands:     # cadangan: cari lewat gid
            cands = [(title, gid) for title, gid in tabs if gid is not None and gid in KNOWN_TAB_GIDS.get(kind, [])]
        if not cands:
            notes.append(f"{kind}: tab '{kind}' (gid {KNOWN_TAB_GIDS.get(kind)}) tidak ditemukan di Google Sheet.")

        remaining = set(needed_dates)
        for title, gid in cands:
            if not remaining:
                break
            try:
                rows = fetch(title)
            except Exception as e:
                notes.append(f"{kind}: gagal baca tab '{title}': {_short_err(e)}")
                continue
            tab_info = {}
            layout, counts, covered = extract_counts(rows, kind, tab_info)
            if layout is None:
                notes.append(f"{kind}: tab '{title}' tidak dikenali bentuknya (isi 5 baris pertama: {rows[:5]})")
                continue
            if layout == "periode-banyak-hari":
                notes.append(f"{kind}: tab '{title}' adalah summary untuk rentang banyak hari (Start≠End), "
                             f"tidak bisa dipecah per hari -- dilewati.")
                continue
            got = {d for d in remaining if covered(d)}
            for d in got:
                out[kind][d] = counts.get(d, {})
            if got:
                cara = ("hitung " + ("Live ID" if kind == "LS" else "Video ID") + " unik per creator per tanggal"
                        if tab_info.get("uniq") else "hitung baris (kolom ID tidak ditemukan!)")
                print(f"  {kind}: pakai tab '{title}' -> {cara}, {len(got)} tanggal.")
                if layout == "mentah" and not tab_info.get("uniq"):
                    notes.append(f"{kind}: kolom '{'Live ID' if kind == 'LS' else 'Video ID'}' tidak ketemu di tab '{title}', "
                                 f"jadi yang dihitung jumlah BARIS (bisa lebih besar dari jumlah unik).")
                if tab_info.get("no_id_rows"):
                    notes.append(f"{kind}: {tab_info['no_id_rows']} baris tanpa ID dilewati (tidak ikut dihitung).")
            remaining -= got
        if remaining:
            notes.append(f"{kind}: data di tab belum sampai tanggal {sorted(d.isoformat() for d in remaining)} "
                         f"-- kolom {kind} di tracker dibiarkan kosong untuk tanggal itu.")
    return out, notes


# ---------------------- baca struktur tracker -------------------

def _analyze_tracker_openpyxl(tracker_bytes: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(tracker_bytes), data_only=False)
    ms = wb["Setting"]["C4"].value
    if isinstance(ms, dt.datetime):
        ms = ms.date()
    if not isinstance(ms, dt.date):
        raise ValueError("Setting!C4 (Awal bulan) tidak terbaca sebagai tanggal.")

    layouts = {}
    for name in ROSTER_TABS:
        if name not in wb.sheetnames:
            print(f"  [PERINGATAN] Tab '{name}' tidak ada di tracker -- dilewati.")
            continue
        ws = wb[name]
        sub_row = None
        for r in range(6, 26):
            if any(str(ws.cell(row=r, column=c).value or "").strip() == "GMV Total" for c in range(4, 30)):
                sub_row = r
                break
        if sub_row is None:
            print(f"  [PERINGATAN] Tab '{name}': baris header 'GMV Total' tidak ketemu -- dilewati.")
            continue

        starts = [c for c in range(4, ws.max_column + 1) if str(ws.cell(row=sub_row, column=c).value or "").strip() == "GMV Total"]
        day_base = 14
        for c in starts:
            mm = re.search(r"Setting!\$B\$(\d+)", str(ws.cell(row=4, column=c).value or ""))
            if mm:
                day_base = int(mm.group(1))   # blok pertama = hari ke-0
                break
        blocks = {}
        for pos, c in enumerate(starts):
            f4 = ws.cell(row=4, column=c).value
            m = re.search(r"Setting!\$B\$(\d+)", str(f4)) if f4 else None
            if not m:
                continue  # blok TOTAL BULAN dll.
            day_index = int(m.group(1)) - day_base
            if not (0 <= day_index <= 30) or day_index in blocks:
                continue
            nxt = starts[pos + 1] if pos + 1 < len(starts) else c + 9
            offs = {}
            for cc in range(c, min(nxt, c + 9)):
                h = ws.cell(row=sub_row, column=cc).value
                key = {"GMV Total": "total", "GMV LS": "gmv_ls", "GMV SV": "gmv_sv",
                       "GMV SL": "gmv_sl", "LS": "ls", "SV": "sv"}.get(h)
                if key:
                    offs[key] = cc
            if {"total", "gmv_ls", "gmv_sv", "gmv_sl"} <= set(offs):
                blocks[day_index] = offs

        first_row = sub_row + 1
        total_row = None
        for r in range(first_row, ws.max_row + 1):
            if str(ws.cell(row=r, column=2).value or "").strip().upper() == "TOTAL":
                total_row = r
                break
        if total_row is None:
            print(f"  [PERINGATAN] Tab '{name}': baris TOTAL tidak ketemu -- dilewati.")
            continue

        exact, loose = {}, {}
        last_row, last_no = None, 0
        for r in range(first_row, total_row):
            nm = ws.cell(row=r, column=2).value
            if nm in (None, ""):
                continue
            ek = exact_key(nm)
            exact.setdefault(ek, []).append(r)
            loose.setdefault(norm_name(nm), set()).add(ek)
            last_row = r
            no = ws.cell(row=r, column=1).value
            if isinstance(no, (int, float)):
                last_no = int(no)
        # kolom TOTAL BULAN di kanan (GH..GM di tab 6-kolom): dipakai untuk rumus baris creator baru
        tail = None
        all_starts = sorted(starts)
        formula_cols = [c for c in all_starts if str(ws.cell(row=4, column=c).value or "").startswith("=")]
        if len(all_starts) >= 2 and all_starts[1] - all_starts[0] == 6 and formula_cols:
            first_data_col = all_starts[0]
            tail_first = max(formula_cols) + 6
            tail = {"cols": [], "first": first_data_col, "last": tail_first - 1, "hdr_row": sub_row}
            for cc in range(tail_first, tail_first + 6):
                h = ws.cell(row=sub_row, column=cc).value
                if h in ("GMV Total", "GMV LS", "GMV SV", "GMV SL", "LS", "SV"):
                    tail["cols"].append((cc, h))
            if len(tail["cols"]) != 6:
                tail = None
        layouts[name] = {
            "sub_row": sub_row, "first_row": first_row, "total_row": total_row,
            "last_row": last_row, "last_no": last_no, "blocks": blocks, "exact": exact, "loose": loose,
            "tail": tail,
        }
    return ms, layouts


class SheetNotFound(Exception):
    pass


def _load_shared_strings(zin) -> list:
    if "xl/sharedStrings.xml" not in zin.namelist():
        return []
    x = zin.read("xl/sharedStrings.xml").decode("utf-8")
    out = []
    for m in re.finditer(r"<si>(.*?)</si>|<si\s*/>", x, re.S):
        body = m.group(1) or ""
        out.append(html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))))
    return out


def _xml_cell_value(attrs: str, inner: str, ss: list):
    t = re.search(r'\bt="(\w+)"', attrs)
    t = t.group(1) if t else "n"
    v = re.search(r"<v>(.*?)</v>", inner or "", re.S)
    if t == "s":
        return ss[int(v.group(1))] if v else None
    if t == "inlineStr":
        return html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner or "", re.S)))
    if v is None:
        return None
    if t in ("str", "e"):
        return html.unescape(v.group(1))
    try:
        n = float(v.group(1))
    except ValueError:
        return v.group(1)
    return int(n) if n.is_integer() else n


def _parse_header_cells(xml: str, ss: list, last_row: int = 25) -> dict:
    """Sel pada baris 1..last_row -> {(row, col): {"v": nilai, "f": teks rumus / None}} (rumus bersama diterjemahkan)."""
    from openpyxl.formula.translate import Translator
    cut = xml.find(f'<row r="{last_row + 1}"')
    head = xml if cut < 0 else xml[:cut]
    cells, masters = {}, {}
    for m in CELL_RE.finditer(head):
        col_letters, row = m.group(1), int(m.group(2))
        attrs, inner = m.group(3), m.group(4) or ""
        f = None
        fm = re.search(r"<f([^>]*?)(?:/>|>(.*?)</f>)", inner, re.S)
        if fm:
            fattrs = fm.group(1)
            ftext = html.unescape(fm.group(2)) if fm.group(2) is not None else None
            if 't="shared"' in fattrs:
                si = re.search(r'si="(\d+)"', fattrs).group(1)
                if ftext:
                    masters[si] = (f"{col_letters}{row}", ftext)
                    f = ftext
                elif si in masters:
                    origin, mtext = masters[si]
                    f = Translator("=" + mtext, origin=origin).translate_formula(f"{col_letters}{row}")[1:]
            else:
                f = ftext
        cells[(row, col_to_idx(col_letters))] = {"v": _xml_cell_value(attrs, inner, ss), "f": f}
    return cells


_AB_RE = re.compile(r'<c r="([AB])(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def _day_index_map(starts, hc, ms, tab_name) -> dict:
    """
    Kolom awal tiap blok -> nomor hari (0 = tanggal 1). Urutan cara membaca:
      1. rumus di baris 4 yang menunjuk ke Setting!$B$N (blok pertama = hari ke-0, jadi tidak tergantung baris awal kalender),
      2. nilai tanggal hasil hitung di baris 4,
      3. urutan kolom (31 blok pertama) -- dicetak sebagai catatan.
    """
    matched = {}
    for c in starts:
        c4 = hc.get((4, c)) or {}
        f4 = ("=" + c4["f"]) if c4.get("f") else c4.get("v")
        mm = re.search(r"Setting!\$B\$(\d+)", str(f4)) if f4 else None
        if mm:
            matched[c] = int(mm.group(1))
    if matched:
        base = matched[min(matched)]
        return {c: n - base for c, n in matched.items()}
    by_date = {}
    for c in starts:
        v = (hc.get((4, c)) or {}).get("v")
        if isinstance(v, (int, float)) and not isinstance(v, bool) and 30000 < v < 70000:
            by_date[c] = (_SHEETS_EPOCH + dt.timedelta(days=int(v)) - ms).days
    if by_date:
        return by_date
    if len(starts) >= 28:
        print(f"  [CATATAN] Tab '{tab_name}': tanggal blok tidak terbaca dari baris 4 -- dipakai urutan kolom "
              f"(blok pertama = tanggal 1). Pastikan urutannya memang berurutan per tanggal.")
        return {c: i for i, c in enumerate(starts[:31])}
    return {}


def _analyze_tracker_fast(tracker_bytes: bytes):
    """Sama hasilnya dengan cara openpyxl, tapi membaca XML langsung (jauh lebih ringan untuk tracker besar)."""
    zin = zipfile.ZipFile(io.BytesIO(tracker_bytes))
    paths = _sheet_paths(zin)
    if "Setting" not in paths:
        raise SheetNotFound(f"Sheet 'Setting' tidak ada di tracker. Nama sheet yang ada: {list(paths)}")
    ss = _load_shared_strings(zin)

    set_xml = zin.read(paths["Setting"]).decode("utf-8")
    m = re.search(r'<c r="C4"([^>]*?)(?:/>|>(.*?)</c>)', set_xml, re.S)
    v = _xml_cell_value(m.group(1), m.group(2), ss) if m else None
    if not isinstance(v, (int, float)):
        raise ValueError("Setting!C4 (Awal bulan) tidak terbaca sebagai tanggal.")
    ms = _SHEETS_EPOCH + dt.timedelta(days=int(v))

    layouts = {}
    for name in ROSTER_TABS:
        if name not in paths:
            print(f"  [PERINGATAN] Tab '{name}' tidak ada di tracker -- dilewati. Nama sheet yang ada: {list(paths)}")
            continue
        xml = zin.read(paths[name]).decode("utf-8")
        hc = _parse_header_cells(xml, ss)
        val = lambda r, c: (hc.get((r, c)) or {}).get("v")

        sub_row = None
        for r in range(6, 26):
            if any(str(val(r, c) or "").strip() == "GMV Total" for c in range(4, 30)):
                sub_row = r
                break
        if sub_row is None:
            print(f"  [PERINGATAN] Tab '{name}': baris header 'GMV Total' tidak ketemu (dicari di baris 6-25) -- dilewati.")
            continue

        starts = sorted(c for (r, c), d in hc.items() if r == sub_row and c >= 4 and str(d["v"] or "").strip() == "GMV Total")
        day_of = _day_index_map(starts, hc, ms, name)
        blocks = {}
        for pos, c in enumerate(starts):
            day_index = day_of.get(c)
            if day_index is None or not (0 <= day_index <= 30) or day_index in blocks:
                continue
            nxt = starts[pos + 1] if pos + 1 < len(starts) else c + 9
            offs = {}
            for cc in range(c, min(nxt, c + 9)):
                key = {"GMV Total": "total", "GMV LS": "gmv_ls", "GMV SV": "gmv_sv",
                       "GMV SL": "gmv_sl", "LS": "ls", "SV": "sv"}.get(val(sub_row, cc))
                if key:
                    offs[key] = cc
            if {"total", "gmv_ls", "gmv_sv", "gmv_sl"} <= set(offs):
                blocks[day_index] = offs

        first_row = sub_row + 1
        pos0 = xml.find(f'<row r="{first_row}"')
        names, a_vals, total_row = {}, {}, None
        if pos0 >= 0:
            for cm in _AB_RE.finditer(xml, pos0):
                col, r = cm.group(1), int(cm.group(2))
                if r < first_row:
                    continue
                value = _xml_cell_value(cm.group(3), cm.group(4), ss)
                if col == "A":
                    a_vals[r] = value
                else:
                    if isinstance(value, str) and value.strip().upper() == "TOTAL":
                        total_row = r
                        break
                    if value not in (None, ""):
                        names[r] = value
        if total_row is None:
            print(f"  [PERINGATAN] Tab '{name}': baris TOTAL tidak ketemu -- dilewati.")
            continue

        exact, loose = {}, {}
        last_row, last_no = None, 0
        for r in sorted(names):
            nm = names[r]
            ek = exact_key(nm)
            exact.setdefault(ek, []).append(r)
            loose.setdefault(norm_name(nm), set()).add(ek)
            last_row = r
            no = a_vals.get(r)
            if isinstance(no, (int, float)):
                last_no = int(no)

        tail = None
        day_cols = [c for c in starts if day_of.get(c) is not None and 0 <= day_of[c] <= 30]
        if len(starts) >= 2 and starts[1] - starts[0] == 6 and day_cols:
            tail_first = max(day_cols) + 6
            tail = {"cols": [], "first": starts[0], "last": tail_first - 1, "hdr_row": sub_row}
            for cc in range(tail_first, tail_first + 6):
                h = val(sub_row, cc)
                if h in ("GMV Total", "GMV LS", "GMV SV", "GMV SL", "LS", "SV"):
                    tail["cols"].append((cc, h))
            if len(tail["cols"]) != 6:
                tail = None
        layouts[name] = {
            "sub_row": sub_row, "first_row": first_row, "total_row": total_row,
            "last_row": last_row, "last_no": last_no, "blocks": blocks, "exact": exact, "loose": loose,
            "tail": tail,
        }
    return ms, layouts


def analyze_tracker(tracker_bytes: bytes):
    try:
        return _analyze_tracker_fast(tracker_bytes)
    except SheetNotFound:
        raise
    except Exception as e:
        print(f"  (Pembacaan cepat gagal: {type(e).__name__}: {e} -- memakai cara lambat.)")
        return _analyze_tracker_openpyxl(tracker_bytes)


# --------------------------- rencana isi ------------------------

def build_plan(month_start, layouts, src_by_date, counts):
    """
    Aturan pencocokan nama:
      1. Persis sama (tanpa peduli huruf besar/kecil)  -> dipakai.
      2. Kalau tidak ada: samakan dengan membuang spasi/titik/underscore, TAPI hanya kalau di tab itu
         cuma ada satu kandidat (kalau ambigu -> dianggap tidak ketemu).
      3. Satu creator hanya diisi di SATU tab, urutan prioritas = urutan ROSTER_TABS
         (supaya creator yang tercatat di dua tab tidak terhitung dobel).
      4. Kalau nama dobel dalam satu tab -> hanya baris pertama yang diisi.
    Return plan = {"cells": {tab: {(row,col): nilai}}, "new_rows": {tab: [...]}, "stats": {...}}
    """
    order = [t for t in ROSTER_TABS if t in layouts and t != MCN_TAB]   # MCN diisi dari file Agency list
    cells = {n: {} for n in layouts}
    new_rows = {n: [] for n in layouts}
    stats = {"per_tab": {n: 0 for n in layouts}, "skipped_dates": [], "added": 0,
             "loose": [], "cross_tab": set(), "dup_rows": set()}

    for lay in layouts.values():
        lay["_next_row"] = lay["total_row"]
        lay["_next_no"] = lay["last_no"] + 1

    cache = {}
    added_today = set()

    def in_mcn(key):
        """Nama ini terdaftar di tab MCN? (diisi lewat cadangan dari Creator_List, jangan ditambah ke Community)"""
        lay_m = layouts.get(MCN_TAB)
        if not lay_m:
            return False
        if key in lay_m["exact"]:
            return True
        names = lay_m["loose"].get(norm_name(key))
        return bool(names) and len(names) == 1

    def resolve(key):
        if key in cache:
            return cache[key]
        hit = None
        matched_tabs = [t for t in order if key in layouts[t]["exact"]]
        if matched_tabs:
            t = matched_tabs[0]
            rows = layouts[t]["exact"][key]
            hit = (t, rows[0], None)
            if len(matched_tabs) > 1:
                stats["cross_tab"].add(key)
            if len(rows) > 1:
                stats["dup_rows"].add((t, key))
        else:
            ln = norm_name(key)
            for t in order:
                names = layouts[t]["loose"].get(ln)
                if names and len(names) == 1:
                    nm = next(iter(names))
                    hit = (t, layouts[t]["exact"][nm][0], nm)
                    stats["loose"].append((key, nm, t))
                    break
        cache[key] = hit
        return hit

    for date in sorted(src_by_date):
        day_index = (date - month_start).days
        if not (0 <= day_index <= 30) or date.month != month_start.month:
            stats["skipped_dates"].append(date)
            continue
        src = src_by_date[date]
        ls_c = counts.get("LS", {}).get(date)
        sv_c = counts.get("SV", {}).get(date)

        added_today = set()
        # creator baru -> Community (hanya kalau belum tercatat di tab manapun)
        if COMMUNITY_TAB in layouts:
            com = layouts[COMMUNITY_TAB]
            for key, rec in src.items():
                if rec["total"] > 0 and rec["total"] >= COMMUNITY_MIN_GMV and resolve(key) is None and not in_mcn(key):
                    row = com["_next_row"]
                    com["_next_row"] += 1
                    new_rows[COMMUNITY_TAB].append({"row": row, "no": com["_next_no"], "name": rec["name"]})
                    com["_next_no"] += 1
                    com["exact"][key] = [row]
                    com["loose"].setdefault(norm_name(key), set()).add(key)
                    cache.pop(key, None)
                    added_today.add(key)
                    stats["added"] += 1

        # kumpulkan per BARIS TUJUAN (nama yang ditulis beda di file GMV vs sheet LS/SV tetap jatuh ke baris yang sama)
        targets = {}

        def target_of(key):
            hit = resolve(key)
            if hit is None:
                return None
            tab, row, _ = hit
            if not layouts[tab]["blocks"].get(day_index):
                return None
            return targets.setdefault((tab, row), {"total": 0.0, "gmv_ls": 0.0, "gmv_sv": 0.0, "gmv_sl": 0.0, "ls": 0, "sv": 0,
                                                   "name": None, "new": False, "loose": hit[2], "in_file": False})

        skipped_rows = []
        for key, rec in src.items():
            t = target_of(key)
            if t is not None:
                for f in ("total", "gmv_ls", "gmv_sv", "gmv_sl"):
                    t[f] += rec[f]
                t["name"] = t["name"] or rec["name"]
                t["in_file"] = True
                if key in added_today:
                    t["new"] = True
            else:
                why = ("terdaftar di tab MCN (nama tidak ada di file Agency) -> diisi dari Creator_List ke tab MCN" if in_mcn(key) else "belum terdaftar & GMV di bawah batas") \
                    if resolve(key) is None else "blok tanggal tidak ada di tab"
                skipped_rows.append({"name": rec["name"], "tab": "-", "row": None, "total": rec["total"], "gmv_ls": rec["gmv_ls"],
                                     "gmv_sv": rec["gmv_sv"], "gmv_sl": rec["gmv_sl"], "ls": None, "sv": None,
                                     "status": f"TIDAK ditulis: {why}"})
        for key, n in (ls_c or {}).items():
            t = target_of(key)
            if t is not None:
                t["ls"] += n
                t["name"] = t["name"] or key
        for key, n in (sv_c or {}).items():
            t = target_of(key)
            if t is not None:
                t["sv"] += n
                t["name"] = t["name"] or key

        stats.setdefault("recon", {})[date] = {"file": sum(r["total"] for r in src.values()),
                                              "written": sum(t["total"] for t in targets.values())}
        pv_rows = []
        for (tab, row), t in targets.items():
            pv_rows.append({"name": t["name"] or "?", "tab": tab, "row": row, "total": t["total"], "in_file": t["in_file"],
                            "gmv_ls": t["gmv_ls"], "gmv_sv": t["gmv_sv"], "gmv_sl": t["gmv_sl"],
                            "ls": t["ls"] if ls_c is not None else None, "sv": t["sv"] if sv_c is not None else None,
                            "status": ("BARU di Community" if t["new"] else ("cocok (nama mirip)" if t.get("loose") else "terdaftar"))
                                      if t["in_file"] else "hanya LS/SV (tidak ada di file GMV)"})
        stats.setdefault("preview", {})[date] = {"rows": pv_rows, "skipped": skipped_rows, "n_file": len(src)}
        # nilai LAMA di baris yang tidak ada di file hari itu dikosongkan -> total tracker = total file
        for tab in order:
            blk0 = layouts[tab]["blocks"].get(day_index)
            if not blk0:
                continue
            all_rows = {r for rs in layouts[tab]["exact"].values() for r in rs}
            for r0 in all_rows:
                if (tab, r0) in targets:
                    continue
                for fld in ("total", "gmv_ls", "gmv_sv", "gmv_sl"):
                    if fld in blk0:
                        cells[tab].setdefault((r0, blk0[fld]), None)
                if ls_c is not None and "ls" in blk0:
                    cells[tab].setdefault((r0, blk0["ls"]), None)
                if sv_c is not None and "sv" in blk0:
                    cells[tab].setdefault((r0, blk0["sv"]), None)
        for (tab, row), t in targets.items():
            blk = layouts[tab]["blocks"][day_index]
            cc = cells[tab]
            cc[(row, blk["total"])] = t["total"]
            cc[(row, blk["gmv_ls"])] = t["gmv_ls"]
            cc[(row, blk["gmv_sv"])] = t["gmv_sv"]
            cc[(row, blk["gmv_sl"])] = t["gmv_sl"]
            if "ls" in blk and ls_c is not None:
                cc[(row, blk["ls"])] = t["ls"]
            if "sv" in blk and sv_c is not None:
                cc[(row, blk["sv"])] = t["sv"]
            stats["per_tab"][tab] += 1
    return {"cells": cells, "new_rows": new_rows, "stats": stats}


def build_mcn_plan(plan, month_start, layouts, mcn_by_date):
    """Isi tab MCN dari file Agency list: GMV Total, LS, SV saja (GMV LS/SV/SL tidak disentuh)."""
    st = plan["stats"].setdefault("mcn", {"filled": 0, "unmatched": {}, "loose": [], "skipped_dates": [], "dup": set()})
    lay = layouts.get(MCN_TAB)
    if lay is None:
        st["no_tab"] = True
        return
    cells = plan["cells"].setdefault(MCN_TAB, {})
    for date in sorted(mcn_by_date):
        day_index = (date - month_start).days
        blk = lay["blocks"].get(day_index)
        if not (0 <= day_index <= 30) or date.month != month_start.month or not blk:
            st["skipped_dates"].append(date)
            continue
        st.setdefault("recon", {}).setdefault(date, {"file": 0.0, "written": 0.0})["file"] = sum(r["gmv"] for r in mcn_by_date[date].values())
        for key, rec in mcn_by_date[date].items():
            rows = lay["exact"].get(key)
            if rows:
                row = rows[0]
                if len(rows) > 1:
                    st["dup"].add(key)
            else:
                cand = lay["loose"].get(norm_name(key))
                if cand and len(cand) == 1:
                    nm = next(iter(cand))
                    row = lay["exact"][nm][0]
                    st["loose"].append((rec["name"], nm))
                else:
                    st["unmatched"].setdefault(rec["name"], 0.0)
                    st["unmatched"][rec["name"]] += rec["gmv"]
                    st.setdefault("preview", {}).setdefault(date, []).append(
                        {"name": rec["name"], "gmv": rec["gmv"], "ls": rec["ls"], "sv": rec["sv"], "status": "TIDAK ada di tab MCN"})
                    continue
            st.setdefault("preview", {}).setdefault(date, []).append(
                {"name": rec["name"], "gmv": rec["gmv"], "ls": rec["ls"], "sv": rec["sv"], "status": f"-> tab MCN baris {row}"})
            cells[(row, blk["total"])] = rec["gmv"]
            rc = st.setdefault("recon", {}).setdefault(date, {"file": 0.0, "written": 0.0})
            rc["written"] += rec["gmv"]
            if rec["ls"] is not None and "ls" in blk:
                cells[(row, blk["ls"])] = rec["ls"]
            if rec["sv"] is not None and "sv" in blk:
                cells[(row, blk["sv"])] = rec["sv"]
            st["filled"] += 1
            plan["stats"]["per_tab"][MCN_TAB] = plan["stats"]["per_tab"].get(MCN_TAB, 0) + 1
        written_rows = {r for (r, c) in cells if c == blk["total"] and cells[(r, c)] is not None}
        for r0 in {r for rs in lay["exact"].values() for r in rs}:
            if r0 not in written_rows:
                for fld in ("total", "ls", "sv"):
                    if fld in blk:
                        cells.setdefault((r0, blk[fld]), None)


def build_mcn_fallback(plan, month_start, layouts, mcn_by_date, src_by_date, counts):
    """
    Nama di tab MCN yang TIDAK ada di file Agency list (folder MCN) dicari di file Creator_List (folder Drive utama);
    kalau ketemu, GMV Total, GMV LS, GMV SV, GMV SL (dan LS/SV dari sheet LS/SV) diisi dari sana, sama seperti tab creator. Nama yang sudah ada di tab creator lain
    (Celeb, TOP VIP, dst.) dilewati supaya GMV tidak terhitung dobel.
    """
    lay = layouts.get(MCN_TAB)
    if lay is None or not src_by_date:
        return
    st = plan["stats"].setdefault("mcn", {"filled": 0, "unmatched": {}, "loose": [], "skipped_dates": [], "dup": set()})
    fb = st.setdefault("fallback", {"total": {}, "rows": {}, "in_other_tab": set()})
    cells = plan["cells"].setdefault(MCN_TAB, {})
    others = [t for t in ROSTER_TABS if t in layouts and t != MCN_TAB]

    def in_other_tab(key):
        for t in others:
            if key in layouts[t]["exact"]:
                return t
            nm = layouts[t]["loose"].get(norm_name(key))
            if nm and len(nm) == 1:
                return t
        return None

    for date in sorted(src_by_date):
        day_index = (date - month_start).days
        blk = lay["blocks"].get(day_index)
        if not (0 <= day_index <= 30) or date.month != month_start.month or not blk:
            continue
        src = src_by_date[date]
        loose_src = {}
        for k in src:
            loose_src.setdefault(norm_name(k), []).append(k)
        ls_c = (counts or {}).get("LS", {}).get(date) if counts else None
        sv_c = (counts or {}).get("SV", {}).get(date) if counts else None
        agency_rows = {r for (r, c), v in cells.items() if c == blk["total"] and v is not None}
        for key, rows in lay["exact"].items():
            row = rows[0]
            if row in agency_rows:
                continue                       # sudah terisi dari file Agency list
            rec = src.get(key)
            k_src = key
            if rec is None:
                cand = loose_src.get(norm_name(key))
                if cand and len(cand) == 1:
                    k_src = cand[0]
                    rec = src[k_src]
            if rec is None:
                continue
            ot = in_other_tab(key) or in_other_tab(k_src)
            if ot:
                fb["in_other_tab"].add((rec["name"], ot))
                continue
            cells[(row, blk["total"])] = rec["total"]
            for fld in ("gmv_ls", "gmv_sv", "gmv_sl"):      # sama seperti tab creator: semua kolom GMV ikut diisi
                if fld in blk:
                    cells[(row, blk[fld])] = rec[fld]
            if "ls" in blk and ls_c is not None:
                cells[(row, blk["ls"])] = ls_c.get(k_src, ls_c.get(key, 0))
            if "sv" in blk and sv_c is not None:
                cells[(row, blk["sv"])] = sv_c.get(k_src, sv_c.get(key, 0))
            fb["total"][date] = fb["total"].get(date, 0.0) + rec["total"]
            fb["rows"].setdefault(date, []).append({"name": rec["name"], "gmv": rec["total"], "row": row})
            rc = plan["stats"].setdefault("recon", {}).get(date)
            if rc is not None:
                rc["written"] += rec["total"]
            st["filled"] += 1
            plan["stats"]["per_tab"][MCN_TAB] = plan["stats"]["per_tab"].get(MCN_TAB, 0) + 1


# ------------------------ tambal XML .xlsx ----------------------

ROW_RE = re.compile(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def fmt_num(v) -> str:
    if isinstance(v, float) and v == int(v):
        v = int(v)
    if isinstance(v, int):
        return str(v)
    return repr(round(float(v), 6))


def _cell_xml(col_idx, row, attrs, inner):
    ref = f"{idx_to_col(col_idx)}{row}"
    if inner:
        return f'<c r="{ref}"{attrs}>{inner}</c>'
    return f'<c r="{ref}"{attrs}/>'


def _parse_cells(inner):
    cells = OrderedDict()
    if not inner:
        return cells
    for m in CELL_RE.finditer(inner):
        cells[col_to_idx(m.group(1))] = [m.group(3) or "", m.group(4) or ""]
    return cells


def _style_attr(attrs):
    m = re.search(r'\ss="(\d+)"', attrs)
    return f' s="{m.group(1)}"' if m else ""


def patch_sheet_xml(xml, cells, new_rows, first_row, total_row, tail_info=None):
    """
    cells    : {(row, col): number}
    new_rows : [{"row","no","name"}] -> disisipkan tepat sebelum baris TOTAL (baris TOTAL turun)
    """
    k = len(new_rows)
    s0 = xml.index("<sheetData>") + len("<sheetData>")
    s1 = xml.index("</sheetData>")
    head, body, tail = xml[:s0], xml[s0:s1], xml[s1:]

    rows = [(int(m.group(1)), m.group(2), m.group(3), m.group(0)) for m in ROW_RE.finditer(body)]
    if "".join(r[3] for r in rows) != body:
        raise RuntimeError("Struktur baris XML tidak seperti yang diperkirakan; file tidak diubah.")

    by_row = {}
    for (r, c), v in cells.items():
        by_row.setdefault(r, {})[c] = v

    # gaya (style) kolom diambil dari baris creator pertama, untuk sel yang belum ada
    default_style = {}
    template_last = None
    for rnum, attrs, inner, raw in rows:
        if rnum == first_row:
            for c, (a, _) in _parse_cells(inner).items():
                default_style[c] = _style_attr(a)
        if rnum == total_row - 1:
            template_last = (attrs, inner)

    out = []
    skipped_formula = 0
    for rnum, attrs, inner, raw in rows:
        if k and rnum == total_row:
            if template_last is None:
                raise RuntimeError("Baris creator terakhir tidak ketemu untuk dijadikan contoh.")
            t_attrs, t_inner = template_last
            t_cells = _parse_cells(t_inner)
            for nr in new_rows:
                row_cells = OrderedDict()
                if 1 in t_cells:
                    row_cells[1] = [t_cells[1][0].replace(' t="s"', ''), f"<v>{nr['no']}</v>"]
                b_attrs = re.sub(r'\st="[^"]*"', "", t_cells.get(2, [""])[0])
                name = xml_escape(nr["name"])
                space = ' xml:space="preserve"' if name != name.strip() else ""
                row_cells[2] = [b_attrs + ' t="inlineStr"', f"<is><t{space}>{name}</t></is>"]
                if 3 in t_cells:
                    row_cells[3] = list(t_cells[3])
                for c, v in by_row.get(nr["row"], {}).items():
                    if v is None:
                        continue
                    row_cells[c] = [default_style.get(c, ""), f"<v>{fmt_num(v)}</v>"]
                if tail_info:   # rumus TOTAL BULAN per creator (GH..GM), sama polanya dengan baris lain
                    a0, a1 = idx_to_col(tail_info["first"]), idx_to_col(tail_info["last"])
                    for tc, label in tail_info["cols"]:
                        formula = (f'IF($B{nr["row"]}="","",SUMIF(${a0}${tail_info["hdr_row"]}:${a1}${tail_info["hdr_row"]},'
                                   f'"{label}",${a0}{nr["row"]}:${a1}{nr["row"]}))')
                        row_cells[tc] = [default_style.get(tc, ""), f"<f>{xml_escape(formula)}</f>"]
                body_cells = "".join(_cell_xml(c, nr["row"], a, i) for c, (a, i) in sorted(row_cells.items()))
                new_attrs = t_attrs
                out.append(f'<row r="{nr["row"]}"{new_attrs}>{body_cells}</row>')

        if k and rnum >= total_row:
            shifted = re.sub(r'<row r="\d+"', f'<row r="{rnum + k}"', raw, count=1)
            shifted = re.sub(r'<c r="([A-Z]+)\d+"', lambda m: f'<c r="{m.group(1)}{rnum + k}"', shifted)
            if rnum == total_row:
                last = total_row + k - 1
                shifted = re.sub(r'<f>SUM\(([A-Z]+)\d+:([A-Z]+)\d+\)</f>',
                                 lambda m: f'<f>SUM({m.group(1)}{first_row}:{m.group(2)}{last})</f>', shifted)
            out.append(shifted)
            continue

        if rnum in by_row:
            cl = _parse_cells(inner)
            for c, v in by_row[rnum].items():
                if c in cl:
                    a, i = cl[c]
                    if "<f" in i:
                        skipped_formula += 1
                        continue
                    a = re.sub(r'\st="[^"]*"', "", a)
                    cl[c] = [a, "" if v is None else f"<v>{fmt_num(v)}</v>"]
                elif v is not None:
                    cl[c] = [default_style.get(c, ""), f"<v>{fmt_num(v)}</v>"]
            body_cells = "".join(_cell_xml(c, rnum, a, i) for c, (a, i) in sorted(cl.items()))
            out.append(f'<row r="{rnum}"{attrs}>{body_cells}</row>')
        else:
            out.append(raw)

    new_xml = head + "".join(out) + tail

    if k:
        def bump_dim(m):
            return f'<dimension ref="{m.group(1)}{int(m.group(2)) + k}"/>'
        new_xml = re.sub(r'<dimension ref="([A-Z]+\d+:[A-Z]+)(\d+)"/>', bump_dim, new_xml, count=1)
        new_xml = new_xml.replace(f"C{first_row}:C{total_row - 1}</xm:sqref>",
                                  f"C{first_row}:C{total_row - 1 + k}</xm:sqref>")
    return new_xml, skipped_formula


def apply_plan(tracker_bytes: bytes, plan: dict, layouts: dict) -> bytes:
    zin = zipfile.ZipFile(io.BytesIO(tracker_bytes))
    wbxml = zin.read("xl/workbook.xml").decode("utf-8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    sheet_path = _sheet_paths(zin)

    patched = {}
    skipped = 0
    for tab, lay in layouts.items():
        cells = plan["cells"].get(tab, {})
        nrows = plan["new_rows"].get(tab, [])
        if not cells and not nrows:
            continue
        # sel untuk baris baru juga harus ikut (sudah ada di cells); pastikan kunci baris baru terisi
        path = sheet_path[tab]
        xml = zin.read(path).decode("utf-8")
        new_xml, sk = patch_sheet_xml(xml, cells, nrows, lay["first_row"], lay["total_row"], lay.get("tail"))
        skipped += sk
        patched[path] = new_xml.encode("utf-8")

    # workbook.xml : paksa hitung ulang saat dibuka
    if "fullCalcOnLoad" not in wbxml:
        wbxml = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wbxml, count=1)
    patched["xl/workbook.xml"] = wbxml.encode("utf-8")

    # hapus calcChain (Excel membuatnya lagi otomatis)
    drop = set()
    if "xl/calcChain.xml" in zin.namelist():
        drop.add("xl/calcChain.xml")
        rels = re.sub(r'<Relationship\b[^>]*calcChain[^>]*/>', "", rels)
        patched["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
        ct = zin.read("[Content_Types].xml").decode("utf-8")
        ct = re.sub(r'<Override\b[^>]*calcChain[^>]*/>', "", ct)
        patched["[Content_Types].xml"] = ct.encode("utf-8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in drop:
                continue
            data = patched.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
    if skipped:
        print(f"  (info) {skipped} sel berisi rumus dilewati supaya rumus tidak rusak.")
    return buf.getvalue()



def verify_written(new_bytes, plan: dict, layouts: dict, src_by_date: dict, mcn_by_date: dict, month_start, col_sum_fn=None):
    """Baca ULANG file hasil (bukan rencana) dan jumlahkan GMV Total tiap tanggal per tab. Bandingkan dengan file."""
    ok_all = True
    cache = {}
    if col_sum_fn is None:
        zin = zipfile.ZipFile(io.BytesIO(new_bytes))
        paths = _sheet_paths(zin)
        ss = _load_shared_strings(zin)

    def col_sum(tab, col, first, last):
        if tab not in cache:
            cache[tab] = zin.read(paths[tab]).decode("utf-8")
        tot = 0.0
        for m in CELL_RE.finditer(cache[tab]):
            if col_to_idx(m.group(1)) != col:
                continue
            r = int(m.group(2))
            if r < first or r > last:
                continue
            v = _xml_cell_value(m.group(3) or "", m.group(4) or "", ss)
            if isinstance(v, (int, float)):
                tot += v
        return tot

    if col_sum_fn is not None:
        col_sum = col_sum_fn
        pf = getattr(col_sum_fn, "prefetch", None)
        if pf is not None:       # baca semua kolom GMV Total yang akan dicek sekaligus
            need = []
            for date in sorted(set(src_by_date) | set(mcn_by_date)):
                di = (date - month_start).days
                if not (0 <= di <= 30) or date.month != month_start.month:
                    continue
                for tab, lay in layouts.items():
                    blk = lay["blocks"].get(di)
                    if not blk:
                        continue
                    k = len(plan["new_rows"].get(tab, [])) if tab != MCN_TAB else 0
                    need.append((tab, blk["total"], lay["first_row"], lay["total_row"] + k - 1))
            try:
                pf(need)
            except Exception as e:
                print(f"     (membaca sekaligus gagal: {type(e).__name__}; dibaca satu-satu)")
    print("  Cek ULANG dari file hasil (jumlah kolom GMV Total di tracker vs file):")
    for date in sorted(src_by_date):
        di = (date - month_start).days
        if not (0 <= di <= 30) or date.month != month_start.month:
            continue
        file_tot = sum(r["total"] for r in src_by_date[date].values())
        fb_tot = ((plan["stats"].get("mcn") or {}).get("fallback") or {}).get("total", {}).get(date, 0.0)
        file_tot -= fb_tot          # GMV creator yang masuk tab MCN dihitung di tab MCN
        got, per = 0.0, []
        for tab, lay in layouts.items():
            if tab == MCN_TAB or di not in lay["blocks"]:
                continue
            k = len(plan["new_rows"].get(tab, []))
            sm = col_sum(tab, lay["blocks"][di]["total"], lay["first_row"], lay["total_row"] + k - 1)
            got += sm
            per.append(f"{tab} {sm:,.0f}")
        gap = file_tot - got
        okd = abs(gap) < 1
        ok_all = ok_all and okd
        print(f"     {date:%d-%b-%Y}: file Rp {file_tot:,.0f} | tracker Rp {got:,.0f} | selisih Rp {gap:,.0f}  {'OK' if okd else '<-- BEDA'}")
        if not okd:
            print("        per tab: " + "; ".join(per))
            sk = plan["stats"].get("preview", {}).get(date, {}).get("skipped", [])
            if sk:
                print(f"        {len(sk)} creator di file tidak masuk tracker (belum terdaftar/di bawah batas), total Rp "
                      f"{sum(x['total'] for x in sk):,.0f}; terbesar: " + ", ".join(f"{x['name']} ({x['total']:,.0f})" for x in sorted(sk, key=lambda x: -x['total'])[:5]))
    lay = layouts.get(MCN_TAB)
    fb_dates = set((((plan["stats"].get("mcn") or {}).get("fallback") or {}).get("total", {})))
    if lay and (mcn_by_date or fb_dates):
        for date in sorted(set(mcn_by_date) | fb_dates):
            di = (date - month_start).days
            if not (0 <= di <= 30) or di not in lay["blocks"] or date.month != month_start.month:
                continue
            file_tot = sum(r["gmv"] for r in mcn_by_date.get(date, {}).values())
            file_tot += ((plan["stats"].get("mcn") or {}).get("fallback") or {}).get("total", {}).get(date, 0.0)
            got = col_sum(MCN_TAB, lay["blocks"][di]["total"], lay["first_row"], lay["total_row"] - 1)
            okd = abs(file_tot - got) < 1
            ok_all = ok_all and okd
            print(f"     MCN {date:%d-%b-%Y}: file Rp {file_tot:,.0f} | tracker Rp {got:,.0f} | selisih Rp {file_tot - got:,.0f}  {'OK' if okd else '<-- BEDA (agency tidak ada di tab MCN?)'}")
    return ok_all


# ------------- tulis langsung ke Google Sheets asli -------------


def _with_retry(make_request, what="permintaan", retry_net=True):
    """Jalankan permintaan Google; kalau kena batas kuota (429), error sementara, atau koneksi putus, tunggu lalu coba lagi.
    retry_net=False: JANGAN diulang kalau koneksi putus (untuk perintah yang tidak aman diulang, mis. menyisipkan baris)."""
    for attempt in range(1, 9):
        try:
            req = make_request()
            return req.execute() if retry_net else req.execute(num_retries=0)
        except Exception as e:
            status = getattr(getattr(e, "resp", None), "status", None)
            if status in (429, 500, 502, 503, 504) and attempt < 8:
                wait = 20 + 10 * attempt
                print(f"     (kuota/gangguan Google [{status}] pada {what}; menunggu {wait} detik lalu mencoba lagi, percobaan {attempt}/7)")
                time.sleep(wait)
                continue
            if retry_net and is_net_error(e) and attempt < 8:
                wait = 5 * attempt
                print(f"     (koneksi terputus saat {what}: {type(e).__name__}; mencoba lagi dalam {wait} detik, percobaan {attempt}/7)")
                time.sleep(wait)
                continue
            raise


def _gv(sheets_service, ranges, render, chunk=8):
    """Baca banyak rentang, dipecah per beberapa rentang supaya satu jawaban Google tidak terlalu besar
    (httplib2 versi baru menolak jawaban yang mengembang > 100x: DecodeRatioError). Kalau tetap kena, dipecah lebih kecil.
    Hasil bacaan tracker disimpan di disk selama tracker belum diubah -> kalau skrip diulang, tidak membaca ulang."""
    out, i = [], 0
    while i < len(ranges):
        part = ranges[i:i + chunk]
        ck = f"gv|{_TRACKER_MTIME}|{render}|" + "|".join(part) if _TRACKER_MTIME else None
        hit = cache_get(ck, max_age=6 * 3600) if ck else None
        if hit is not None:
            out += hit
            i += len(part)
            continue
        try:
            res = _with_retry(lambda: sheets_service.spreadsheets().values().batchGet(
                spreadsheetId=TRACKER_FILE_ID, ranges=part, valueRenderOption=render), "membaca")
        except Exception as e:
            if type(e).__name__ in ("DecodeRatioError", "DecodeLimitError") and len(part) > 1:
                chunk = max(1, len(part) // 2)
                continue
            raise
        got = [vr.get("values", []) for vr in res.get("valueRanges", [])]
        if ck:
            cache_put(ck, got)
        out += got
        i += len(part)
    return out


def analyze_native(sheets_service):
    """Baca struktur tracker langsung dari Google Sheets lewat API (tanpa ekspor; ekspor dibatasi 10 MB)."""
    meta = sheets_service.spreadsheets().get(spreadsheetId=TRACKER_FILE_ID,
                                             fields="sheets.properties(sheetId,title,gridProperties)").execute()
    props = {_norm_sheet(x["properties"]["title"]): x["properties"] for x in meta["sheets"]}
    if "setting" not in props:
        raise SheetNotFound(f"Sheet 'Setting' tidak ada. Sheet yang ada: {[p['title'] for p in props.values()]}")
    q = lambda t: "'" + t.replace("'", "''") + "'"
    v = _gv(sheets_service, [f"{q(props['setting']['title'])}!C4"], "UNFORMATTED_VALUE")[0]
    serial = v[0][0] if v and v[0] else None
    if not isinstance(serial, (int, float)):
        raise ValueError("Setting!C4 (Awal bulan) tidak terbaca sebagai tanggal.")
    ms = _SHEETS_EPOCH + dt.timedelta(days=int(serial))

    layouts = {}
    for name in ROSTER_TABS:
        pr = props.get(_norm_sheet(name))
        if pr is None:
            print(f"  [PERINGATAN] Tab '{name}' tidak ada di tracker -- dilewati.")
            continue
        t = q(pr["title"])
        max_row = pr["gridProperties"]["rowCount"]
        hv, hf = _gv(sheets_service, [f"{t}!A1:ZZ25"], "FORMATTED_VALUE")[0], \
            _gv(sheets_service, [f"{t}!A1:ZZ25"], "FORMULA")[0]
        hc = {}
        for ri, row in enumerate(hv, start=1):
            for ci, x in enumerate(row, start=1):
                if x not in ("", None):
                    hc[(ri, ci)] = {"v": x, "f": None}
        for ri, row in enumerate(hf, start=1):
            for ci, x in enumerate(row, start=1):
                if isinstance(x, str) and x.startswith("="):
                    d = hc.setdefault((ri, ci), {"v": None, "f": None})
                    d["f"] = x[1:]
        val = lambda r, c: (hc.get((r, c)) or {}).get("v")
        sub_row = None
        for r in range(6, 26):
            if any(str(val(r, c) or "").strip() == "GMV Total" for c in range(4, 30)):
                sub_row = r
                break
        if sub_row is None:
            print(f"  [PERINGATAN] Tab '{name}': baris header 'GMV Total' tidak ketemu (dicari di baris 6-25) -- dilewati.")
            continue
        starts = sorted(c for (r, c), d in hc.items() if r == sub_row and c >= 4 and str(d["v"] or "").strip() == "GMV Total")
        day_of = _day_index_map(starts, hc, ms, name)
        blocks = {}
        for pos, c in enumerate(starts):
            day_index = day_of.get(c)
            if day_index is None or not (0 <= day_index <= 30) or day_index in blocks:
                continue
            nxt = starts[pos + 1] if pos + 1 < len(starts) else c + 9
            offs = {}
            for cc in range(c, min(nxt, c + 9)):
                key = {"GMV Total": "total", "GMV LS": "gmv_ls", "GMV SV": "gmv_sv",
                       "GMV SL": "gmv_sl", "LS": "ls", "SV": "sv"}.get(val(sub_row, cc))
                if key:
                    offs[key] = cc
            if {"total", "gmv_ls", "gmv_sv", "gmv_sl"} <= set(offs):
                blocks[day_index] = offs

        first_row = sub_row + 1
        ab = _gv(sheets_service, [f"{t}!A{first_row}:B{max_row}"], "UNFORMATTED_VALUE")[0]
        names, a_vals, total_row = {}, {}, None
        for i, row in enumerate(ab):
            r = first_row + i
            a = row[0] if len(row) > 0 else None
            b = row[1] if len(row) > 1 else None
            if isinstance(b, str) and b.strip().upper() == "TOTAL":
                total_row = r
                break
            a_vals[r] = a
            if b not in (None, ""):
                names[r] = int(b) if isinstance(b, float) and b.is_integer() else b
        if total_row is None:
            if names:
                # tab tanpa baris TOTAL (mis. Community): baris TOTAL "virtual" = satu baris di bawah creator terakhir.
                # Creator baru disisipkan DI ATAS creator terakhir, jadi rentang rumus ikut melebar.
                total_row = max(names) + 1
                print(f"  [CATATAN] Tab '{name}' tidak punya baris TOTAL -- dipakai sampai baris {max(names)} (creator terakhir); creator baru disisipkan di atasnya.")
            else:
                print(f"  [PERINGATAN] Tab '{name}': baris TOTAL tidak ketemu dan tidak ada nama creator -- dilewati.")
                continue
        exact, loose = {}, {}
        last_row, last_no = None, 0
        for r in sorted(names):
            nm = names[r]
            ek = exact_key(nm)
            exact.setdefault(ek, []).append(r)
            loose.setdefault(norm_name(nm), set()).add(ek)
            last_row = r
            no = a_vals.get(r)
            if isinstance(no, (int, float)):
                last_no = int(no)
        tail = None
        day_cols = [c for c in starts if day_of.get(c) is not None and 0 <= day_of[c] <= 30]
        if len(starts) >= 2 and starts[1] - starts[0] == 6 and day_cols:
            tail_first = max(day_cols) + 6
            tail = {"cols": [], "first": starts[0], "last": tail_first - 1, "hdr_row": sub_row}
            for cc in range(tail_first, tail_first + 6):
                h = val(sub_row, cc)
                if h in ("GMV Total", "GMV LS", "GMV SV", "GMV SL", "LS", "SV"):
                    tail["cols"].append((cc, h))
            if len(tail["cols"]) != 6:
                tail = None
        layouts[name] = {
            "sub_row": sub_row, "first_row": first_row, "total_row": total_row,
            "last_row": last_row, "last_no": last_no, "blocks": blocks, "exact": exact, "loose": loose,
            "tail": tail, "_title": pr["title"],
        }
    return ms, layouts


def _native_template_info(sheets_service, plan: dict, layouts: dict) -> dict:
    """Dari Google Sheets: sel rumus & sel berisi di kolom yang akan ditulis, plus kolom rumus pada baris creator terakhir."""
    info = {}
    q = lambda t: "'" + t.replace("'", "''") + "'"
    for tab, lay in layouts.items():
        cells = plan["cells"].get(tab, {})
        nrows = plan["new_rows"].get(tab, [])
        if not cells and not nrows:
            continue
        cols = sorted({c for (_, c) in cells})
        first, last = lay["first_row"], lay["total_row"] - 1
        t = q(lay["_title"])
        rngs = [f"{t}!{idx_to_col(c)}{first}:{idx_to_col(c)}{last}" for c in cols]
        rngs.append(f"{t}!A{last}:ZZ{last}")
        got = _gv(sheets_service, rngs, "FORMULA")
        formulas, nonempty, old = set(), set(), {}
        for c, colvals in zip(cols, got[:-1]):
            for i, row in enumerate(colvals):
                x = row[0] if row else ""
                if x in ("", None):
                    continue
                r = first + i
                old[(r, c)] = x
                nonempty.add((r, c))
                if isinstance(x, str) and x.startswith("="):
                    formulas.add((r, c))
        tmpl_cols, tmpl_formula_cols = set(), set()
        lastrow_vals = got[-1][0] if got[-1] else []
        for ci, x in enumerate(lastrow_vals, start=1):
            if x in ("", None):
                continue
            tmpl_cols.add(ci)
            if isinstance(x, str) and x.startswith("="):
                tmpl_formula_cols.add(ci)
        info[tab] = {"formulas": formulas, "tmpl_formula_cols": tmpl_formula_cols, "tmpl_cols": tmpl_cols,
                     "nonempty": nonempty, "old": old}
    return info


def _runs(items):
    """[(row, col, val)] -> daftar (col, row_awal, [nilai...]) untuk baris berurutan pada kolom yang sama."""
    by_col = {}
    for r, c, v in items:
        by_col.setdefault(c, []).append((r, v))
    out = []
    for c, lst in by_col.items():
        lst.sort()
        start, vals, prev = None, [], None
        for r, v in lst:
            if start is None or r != prev + 1:
                if start is not None:
                    out.append((c, start, vals))
                start, vals = r, []
            vals.append(v)
            prev = r
        out.append((c, start, vals))
    return out


def _same_value(a, b) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(float(a) - float(b)) < 1e-6
    return str(a) == str(b)


def push_to_native(sheets_service, plan: dict, layouts: dict, info=None):
    """Tulis rencana ke Google Sheets asli lewat Sheets API. Mengembalikan {tab: jumlah_baris_baru}."""
    meta = sheets_service.spreadsheets().get(spreadsheetId=TRACKER_FILE_ID,
                                             fields="sheets.properties(sheetId,title)").execute()
    props = {_norm_sheet(x["properties"]["title"]): x["properties"] for x in meta["sheets"]}
    if info is None:
        info = _native_template_info(sheets_service, plan, layouts)

    # 1) sisipkan baris baru (di DALAM rentang, supaya rumus SUM/ringkasan ikut melebar).
    #    Per tab dikirim SATU permintaan: sisip + salin format/rumus + kosongkan angka salinan + tulis No & nama.
    #    Jadi kalau koneksi putus, hasilnya "jadi semua" atau "tidak sama sekali" -> kalau skrip diulang tidak ada baris kosong/dobel.
    shift = {}
    for tab, lay in layouts.items():
        nrs = sorted(plan["new_rows"].get(tab, []), key=lambda x: x["row"])
        k = len(nrs)
        if not k:
            continue
        sid = props[_norm_sheet(tab)]["sheetId"]
        T = lay["total_row"]
        reqs = [
            {"insertDimension": {"range": {"sheetId": sid, "dimension": "ROWS", "startIndex": T - 2, "endIndex": T - 2 + k},
                                 "inheritFromBefore": True}},
            {"copyPaste": {"source": {"sheetId": sid, "startRowIndex": T - 2 + k, "endRowIndex": T - 1 + k},
                           "destination": {"sheetId": sid, "startRowIndex": T - 2, "endRowIndex": T - 2 + k},
                           "pasteType": "PASTE_NORMAL"}},
        ]
        inf0 = info.get(tab, {})
        clear_cols = sorted(c for c in inf0.get("tmpl_cols", set()) if c not in inf0.get("tmpl_formula_cols", set()) and c > 2)
        runs, run = [], []
        for c in clear_cols:
            if run and c == run[-1] + 1:
                run.append(c)
            else:
                if run:
                    runs.append(run)
                run = [c]
        if run:
            runs.append(run)
        for rn in runs:
            reqs.append({"updateCells": {"range": {"sheetId": sid, "startRowIndex": T - 2, "endRowIndex": T - 2 + k,
                                                   "startColumnIndex": rn[0] - 1, "endColumnIndex": rn[-1]},
                                         "rows": [{"values": [{} for _ in rn]} for _ in range(k)],
                                         "fields": "userEnteredValue"}})
        reqs.append({"updateCells": {"range": {"sheetId": sid, "startRowIndex": T - 2, "endRowIndex": T - 2 + k,
                                               "startColumnIndex": 0, "endColumnIndex": 2},
                                     "rows": [{"values": [{"userEnteredValue": {"numberValue": nr["no"]}},
                                                          {"userEnteredValue": {"stringValue": str(nr["name"])}}]} for nr in nrs],
                                     "fields": "userEnteredValue"}})
        try:
            _with_retry(lambda: sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=TRACKER_FILE_ID, body={"requests": reqs}), f"menyisipkan baris di '{tab}'", retry_net=False)
        except Exception as e:
            if is_net_error(e):
                raise RuntimeError(
                    f"Koneksi terputus saat menyisipkan {k} baris baru di tab '{tab}'. Belum pasti sudah masuk atau belum. "
                    f"Jalankan skrip lagi saja: kalau barisnya sudah masuk, namanya otomatis terdeteksi (tidak dobel); "
                    f"kalau belum, akan disisipkan lagi.") from e
            raise
        shift[tab] = (T, k)

    def actual_row(tab, r):
        if tab not in shift:
            return r
        T, k = shift[tab]
        if r >= T:
            return T - 1 + (r - T)          # baris baru
        if r == T - 1:
            return T - 1 + k                # creator terakhir lama, bergeser ke bawah
        return r

    # 2) kumpulkan nilai
    data = []
    n_cells = 0
    for tab, lay in layouts.items():
        cells = plan["cells"].get(tab, {})
        nrows = plan["new_rows"].get(tab, [])
        if not cells and not nrows:
            continue
        inf = info[tab]
        items = []
        for (r, c), v in cells.items():
            if (r, c) in inf["formulas"]:
                continue
            if v is None and (r, c) not in inf["nonempty"]:
                continue        # sudah kosong, tidak perlu ditulis
            newv = "" if v is None else (int(v) if float(v) == int(v) else round(float(v), 6))
            if (r, c) in inf["old"] and _same_value(inf["old"][(r, c)], newv):
                continue        # sudah berisi nilai yang sama (mis. sudah tertulis sebelum koneksi putus) -> tidak perlu ditulis lagi
            items.append((actual_row(tab, r), c, newv))
        # (No, nama, dan pembersihan angka salinan baris baru sudah dikerjakan di langkah 1)
        # sel baru (nilai terakhir menang)
        dedup = {}
        for r, c, v in items:
            dedup[(r, c)] = v
        real_title = props[_norm_sheet(tab)]["title"].replace("'", "''")
        for c, r0, vals in _runs([(r, c, v) for (r, c), v in dedup.items()]):
            col = idx_to_col(c)
            data.append({"range": f"'{real_title}'!{col}{r0}:{col}{r0 + len(vals) - 1}", "values": [[v] for v in vals]})
            n_cells += len(vals)
    CH = 1500        # banyak rentang per permintaan (kuota Google dihitung per PERMINTAAN, maks 60/menit)
    for n_req, i in enumerate(range(0, len(data), CH), start=1):
        _with_retry(lambda: sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=TRACKER_FILE_ID, body={"valueInputOption": "RAW", "data": data[i:i + CH]}), "menulis")
        print(f"     menulis bagian {n_req}/{(len(data) + CH - 1) // CH} ...")
    return n_cells, shift


def make_native_col_sum(sheets_service, shift):
    """Penjumlah kolom untuk 'Cek ULANG'. .prefetch(daftar) membaca BANYAK kolom sekaligus (jauh lebih sedikit permintaan ke Google)."""
    meta = sheets_service.spreadsheets().get(spreadsheetId=TRACKER_FILE_ID, fields="sheets.properties(title)").execute()
    titles = {_norm_sheet(x["properties"]["title"]): x["properties"]["title"] for x in meta["sheets"]}
    cache = {}

    def _rng(tab, col, first, last):
        title = titles[_norm_sheet(tab)].replace("'", "''")
        c = idx_to_col(col)
        return f"'{title}'!{c}{first}:{c}{last}"

    def prefetch(reqs):
        todo = sorted({r for r in reqs if r not in cache})
        by_size = {}
        for r in todo:
            by_size.setdefault(max(1, r[3] - r[2] + 1), []).append(r)
        for nrows, items in by_size.items():
            per = max(1, min(40, 200000 // nrows))          # ~200 ribu sel per permintaan
            for k in range(0, len(items), per):
                part = items[k:k + per]
                got = _gv(sheets_service, [_rng(*r) for r in part], "UNFORMATTED_VALUE", chunk=per)
                for r, vals in zip(part, got):
                    cache[r] = sum(v for row in vals for v in row if isinstance(v, (int, float)))

    def col_sum(tab, col, first, last):
        key = (tab, col, first, last)
        if key not in cache:
            prefetch([key])
        return cache[key]
    col_sum.prefetch = prefetch
    return col_sum


# ------------------- perbaikan template tracker -----------------

def _norm_sheet(n) -> str:
    return re.sub(r"\s+", " ", str(n).replace("\xa0", " ")).strip().lower()


class _SheetMap(dict):
    """{nama sheet: path file XML}. Pencarian nama tidak peduli huruf besar/kecil, spasi ganda/di ujung, atau spasi tak terlihat."""

    def _find(self, key):
        if dict.__contains__(self, key):
            return key
        nk = _norm_sheet(key)
        for k in self:
            if _norm_sheet(k) == nk:
                return k
        return None

    def __contains__(self, key):
        return self._find(key) is not None

    def __getitem__(self, key):
        k = self._find(key)
        if k is None:
            raise KeyError(f"Sheet '{key}' tidak ada di tracker. Nama sheet yang ada: {list(self)}")
        return dict.__getitem__(self, k)

    def get(self, key, default=None):
        k = self._find(key)
        return default if k is None else dict.__getitem__(self, k)


def _sheet_paths(zin) -> dict:
    import xml.etree.ElementTree as ET
    wb = ET.fromstring(zin.read("xl/workbook.xml"))
    rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    rid_target = {rel.get("Id"): rel.get("Target") for rel in rels if rel.get("Id") and rel.get("Target")}
    out = _SheetMap()
    for el in wb.iter():
        if el.tag.split("}")[-1] != "sheet":
            continue
        rid = next((v for k, v in el.attrib.items() if k.split("}")[-1] == "id"), None)
        name = el.get("name")
        if name is None or rid not in rid_target:
            continue
        t = rid_target[rid]
        out[name] = "xl/" + t.lstrip("/").replace("xl/", "", 1) if t.startswith("/") else "xl/" + t
    return out


def template_needs_fix(tracker_bytes: bytes) -> bool:
    """
    True kalau di Summary PIC baris berlabel 'Community' membaca tab KLUB VIP (atau sebaliknya) -- bawaan template lama
    yang tertukar. Hanya saran: kalau tidak bisa dicek, hasilnya False (tidak pernah menghentikan program).
    """
    try:
        zin = zipfile.ZipFile(io.BytesIO(tracker_bytes))
        paths = _sheet_paths(zin)
        path = paths.get("Summary PIC") or next((p for n, p in paths.items() if "summary" in _norm_sheet(n)), None)
        if path is None:
            print(f"  (info) Sheet 'Summary PIC' tidak ketemu, pengecekan template dilewati. Nama sheet di file: {list(paths)}")
            return False
        ss = _load_shared_strings(zin)
        xml = zin.read(path).decode("utf-8")
        cells = {}
        for m in CELL_RE.finditer(xml):
            col, row = m.group(1), int(m.group(2))
            if row > 90 or col not in ("A", "C", "D"):
                continue
            inner = m.group(4) or ""
            fm = re.search(r"<f[^>]*>(.*?)</f>", inner, re.S)
            cells[(row, col)] = (_xml_cell_value(m.group(3), inner, ss), html.unescape(fm.group(1)) if fm else "")
        crossed = False
        for (row, col), (label, _) in cells.items():
            if col != "A" or not isinstance(label, str):
                continue
            lab = _norm_sheet(label)
            if lab not in ("klub vip", "community"):
                continue
            own, other = ("klub vip", "community") if lab == "klub vip" else ("community", "klub vip")
            for c in ("C", "D"):
                f = _norm_sheet((cells.get((row, c)) or ("", ""))[1])
                if f and other in f and own not in f:
                    crossed = True
        if crossed:
            return True
        # rumus TOTAL BULAN per creator (GH:GM) di baris creator terakhir tab Community harus ada
        import contextlib
        with contextlib.redirect_stdout(io.StringIO()):
            _, lays = analyze_tracker(tracker_bytes)
        comm = lays.get("Community")
        if comm and comm.get("tail") and comm.get("last_row"):
            cx = zin.read(paths["Community"]).decode("utf-8")
            col0 = idx_to_col(comm["tail"]["cols"][0][0])
            if not re.search(r'<c r="%s%d"[^>]*>(?:(?!</c>).)*<f' % (col0, comm["last_row"]), cx, re.S):
                return True
        return False
    except Exception as e:
        print(f"  (info) Pengecekan template dilewati ({type(e).__name__}: {e})")
        return False


def _set_formula(xml: str, ref: str, formula: str) -> str:
    m = re.search(r'<c r="%s"([^>]*?)(?:/>|>(.*?)</c>)' % ref, xml, re.S)
    if not m:
        raise RuntimeError(f"Sel {ref} tidak ditemukan saat memperbaiki template.")
    attrs = re.sub(r'\st="[^"]*"', "", m.group(1))
    return xml[:m.start()] + f'<c r="{ref}"{attrs}><f>{xml_escape(formula)}</f></c>' + xml[m.end():]


def _sheet_ref(name: str) -> str:
    """Nama sheet untuk dipakai di rumus (diberi kutip kalau ada spasi/tanda baca)."""
    return name if re.fullmatch(r"\w+", name) else "'" + name.replace("'", "''") + "'"


def _swap_sheet_refs(text: str, wrong_ref: str, right_ref: str) -> str:
    return re.sub(re.escape(wrong_ref) + "!", lambda m: right_ref + "!", text, flags=re.I)


def fix_template(tracker_bytes: bytes):
    """
    Memperbaiki 2 masalah bawaan template lama (tidak menyentuh data/format lain):
      1. Tab 'KLUB VIP' dan 'Community' tertukar di Summary PIC & Setting (juga judul A1 kedua tab).
      2. Tab Community berisi ribuan creator, tapi rumus (jumlah harian, TOTAL BULAN per creator, baris TOTAL,
         Summary PIC) hanya mencakup sebagian baris -> diperluas ke seluruh daftar, tahan penambahan creator baru.
    Semua posisi (baris header, baris label di Setting/Summary, jangkauan data) dicari dari LABEL dan layout tiap tab,
    bukan dari nomor baris tetap -- jadi aman walau header satu tab beda baris dengan tab lain.
    Return (bytes_baru, daftar_catatan).
    """
    _, layouts = analyze_tracker(tracker_bytes)
    for needed in ("Community", "KLUB VIP"):
        if needed not in layouts:
            raise SheetNotFound(f"Tab '{needed}' tidak terbaca di tracker, perbaikan template tidak bisa dilakukan.")
    comm, klub = layouts["Community"], layouts["KLUB VIP"]
    first, last, total = comm["first_row"], comm["last_row"], comm["total_row"]
    tail = comm.get("tail")
    if not tail:
        raise RuntimeError("Kolom TOTAL BULAN (GH:GM) di tab Community tidak terdeteksi.")
    if not klub.get("tail"):
        raise RuntimeError("Kolom TOTAL BULAN di tab KLUB VIP tidak terdeteksi.")

    zin = zipfile.ZipFile(io.BytesIO(tracker_bytes))
    paths = _sheet_paths(zin)
    ss = _load_shared_strings(zin)
    comm_name, klub_name = paths._find("Community"), paths._find("KLUB VIP")
    comm_ref, klub_ref = _sheet_ref(comm_name), _sheet_ref(klub_name)
    notes, patched = [], {}
    LABELS = ("klub vip", "community")

    # ---------- Setting: baris 'hari berjalan' KLUB VIP / Community (dicari dari label, bukan nomor baris) ----------
    p = paths["Setting"]
    xml = zin.read(p).decode("utf-8")
    hc = _parse_header_cells(xml, ss, last_row=45)
    swapped = 0
    for r in sorted({r for (r, c), d in hc.items() if isinstance(d["v"], str) and _norm_sheet(d["v"]) in LABELS}):
        label = next(_norm_sheet(d["v"]) for (rr, c), d in sorted(hc.items()) if rr == r and isinstance(d["v"], str) and _norm_sheet(d["v"]) in LABELS)
        m = re.search(r'<row r="%d"[^>]*>.*?</row>' % r, xml, re.S)
        if not m:
            continue
        row = m.group(0)
        has = lambda ref: re.search(re.escape(ref) + "!", row, re.I) is not None
        if label == "klub vip" and has(comm_ref) and not has(klub_ref):
            row = _swap_sheet_refs(row, comm_ref, klub_ref)
        elif label == "community" and has(klub_ref) and not has(comm_ref):
            row = _swap_sheet_refs(row, klub_ref, comm_ref)
        else:
            continue
        xml = xml[:m.start()] + row + xml[m.end():]
        swapped += 1
    patched[p] = xml
    if swapped:
        notes.append("Setting: hari berjalan baris KLUB VIP & Community kini membaca tab yang benar.")

    # ---------- Summary PIC: baris berlabel Klub VIP / Community yang tertukar (dicari dari label di kolom A) ----------
    p = paths.get("Summary PIC") or next((pp for n, pp in paths.items() if "summary" in _norm_sheet(n)), None)
    if p is None:
        raise SheetNotFound(f"Sheet 'Summary PIC' tidak ketemu. Nama sheet di file: {list(paths)}")
    xml = zin.read(p).decode("utf-8")
    sc = _parse_header_cells(xml, ss, last_row=90)
    big = 60000
    metric_labels = ("GMV Total", "GMV LS", "GMV SV", "GMV SL")

    def tail_cols(tab_layout):
        return {h: idx_to_col(c) for c, h in tab_layout["tail"]["cols"]}

    fixed_rows = 0
    for r in sorted({r for (r, c), d in sc.items() if c == 1 and isinstance(d["v"], str) and _norm_sheet(d["v"]) in LABELS}):
        label = _norm_sheet(sc[(r, 1)]["v"])
        fC = _norm_sheet((sc.get((r, 3)) or {}).get("f") or "")
        own, other = ("klub vip", "community") if label == "klub vip" else ("community", "klub vip")
        if not fC or other not in fC or own in fC:
            continue                                                   # sudah benar / bukan baris rumus
        is_klub = label == "klub vip"
        tab, ref = (klub, klub_ref) if is_klub else (comm, comm_ref)
        a = tab["first_row"]
        b = tab["last_row"] if is_klub else big
        cols = tail_cols(tab)
        if fC.startswith("countifs("):                                  # baris per PIC
            new = {f"C{r}": f'COUNTIFS({ref}!$C${a}:$C${b},$B{r},{ref}!$B${a}:$B${b},"<>")'}
            for col, lab in zip("DEFG", metric_labels):
                new[f"{col}{r}"] = f"SUMIF({ref}!$C${a}:$C${b},$B{r},{ref}!${cols[lab]}${a}:${cols[lab]}${b})"
        elif fC.startswith("countif("):                                 # baris TOTAL tab
            if is_klub:
                new = {f"C{r}": f'COUNTIF({ref}!$B${a}:$B${b},"<>")'}
                for col, lab in zip("DEFG", metric_labels):
                    new[f"{col}{r}"] = f"{ref}!{cols[lab]}{tab['total_row']}"
            else:
                Bc = f"{ref}!$B${a}:$B${b}"
                new = {f"C{r}": f'COUNTIFS({Bc},"<>",{Bc},"<>TOTAL")'}
                for col, lab in zip("DEFG", metric_labels):
                    new[f"{col}{r}"] = f'INDEX({ref}!${cols[lab]}${a}:${cols[lab]}${b},MATCH("TOTAL",{Bc},0))'
        else:
            continue
        mf = (sc.get((r, 13)) or {}).get("f")                           # kolom 'Ref: Target Tab' -> cukup tukar nama sheet
        if mf:
            new[f"M{r}"] = _swap_sheet_refs(mf, comm_ref if is_klub else klub_ref, ref)
        for cell_ref, formula in new.items():
            xml = _set_formula(xml, cell_ref, formula)
        fixed_rows += 1
    patched[p] = xml
    if fixed_rows:
        notes.append(f"Summary PIC: {fixed_rows} baris KLUB VIP / Community kini membaca tab yang benar (Community: seluruh daftar creator).")

    # ---------- judul A1 kedua tab (hanya kalau memang tertukar) ----------
    for tab_key, tab_name, own, other, title in (("Community", comm_name, "community", "klub vip", "NIVEA TIKTOK – DAILY TRACKER COMMUNITY"),
                                                  ("KLUB VIP", klub_name, "klub vip", "community", "NIVEA TIKTOK – DAILY TRACKER KLUB VIP")):
        pp = paths[tab_key]
        x = patched.get(pp) or zin.read(pp).decode("utf-8")
        a1 = (_parse_header_cells(x, ss, last_row=1).get((1, 1)) or {}).get("v")
        if isinstance(a1, str) and other in _norm_sheet(a1) and own not in _norm_sheet(a1):
            m = re.search(r'<c r="A1"([^>]*?)(?:/>|>(.*?)</c>)', x, re.S)
            attrs = re.sub(r'\st="[^"]*"', "", m.group(1))
            x = x[:m.start()] + f'<c r="A1"{attrs} t="inlineStr"><is><t>{xml_escape(title)}</t></is></c>' + x[m.end():]
            patched[pp] = x
            notes.append(f"Judul tab {tab_name} diperbaiki (tadinya tertukar).")

    # ---------- Community: rumus ----------
    p = paths["Community"]
    xml = patched.get(p) or zin.read(p).decode("utf-8")

    # (a) baris header: jumlah harian (SUM jangkauan pendek -> seluruh daftar, tanpa baris TOTAL) & ref total bulan
    cut = xml.find(f'<row r="{first}"')
    head, rest = xml[:cut], xml[cut:]
    def day_sum(m):
        c = m.group(1)
        return f'<f>IF({c}$4="","",SUMIF($B${first}:$B${big},"&lt;&gt;TOTAL",{c}${first}:{c}${big}))</f>'
    head, n_day = re.subn(r'<f>IF\(([A-Z]+)\$4="","",SUM\(\1\d+:\1\d+\)\)</f>', day_sum, head)
    gh = idx_to_col(tail["cols"][0][0])
    head = re.sub(r'(<c r="%s\d+"[^>]*>)<f>%s\d+</f>' % (gh, gh),
                  lambda m: m.group(1) + f'<f>INDEX(${gh}${first}:${gh}${big},MATCH("TOTAL",$B${first}:$B${big},0))</f>', head)
    xml = head + rest

    # (b) TOTAL BULAN per creator untuk SEMUA baris creator
    tcols = [idx_to_col(c) for c, _ in tail["cols"]]
    labels = [h for _, h in tail["cols"]]
    mfirst = re.search(r'<row r="%d"[^>]*>.*?</row>' % first, xml, re.S).group(0)
    info = {}
    for tc in tcols:
        mm = re.search(r'<c r="%s%d"([^>]*)><f t="shared" ref="([A-Z]+\d+:[A-Z]+\d+)" si="(\d+)">' % (tc, first), mfirst)
        if mm:
            sm = re.search(r's="(\d+)"', mm.group(1))
            info[tc] = (mm.group(2), mm.group(3), sm.group(1) if sm else None)
    use_shared = len(info) == len(tcols)
    style = {}
    for tc in tcols:
        sm = re.search(r'<c r="%s%d"[^>]*?\ss="(\d+)"' % (tc, first), mfirst)
        style[tc] = sm.group(1) if sm else (info[tc][2] if tc in info else None)
    if use_shared:
        for tc, (ref, si, _) in info.items():
            xml = xml.replace(f'ref="{ref}" si="{si}"', f'ref="{tc}{first}:{tc}{last}" si="{si}"', 1)
    a0, a1 = idx_to_col(tail["first"]), idx_to_col(tail["last"])
    added_rows = 0

    def tail_cell(tc, lab, r):
        s_attr = f' s="{style[tc]}"' if style.get(tc) else ""
        if use_shared:
            return f'<c r="{tc}{r}"{s_attr}><f t="shared" si="{info[tc][1]}"/></c>'
        formula = f'IF($B{r}="","",SUMIF(${a0}${tail["hdr_row"]}:${a1}${tail["hdr_row"]},"{lab}",${a0}{r}:${a1}{r}))'
        return f'<c r="{tc}{r}"{s_attr}><f>{xml_escape(formula)}</f></c>'

    def add_tail(m):
        nonlocal added_rows
        r = int(m.group(1))
        if r < first or r > last:
            return m.group(0)
        attrs, inner = m.group(2), m.group(3) or ""
        if re.search(r'<c r="%s%d"[^>]*>(?:(?!</c>).)*<f' % (tcols[0], r), inner, re.S):
            return m.group(0)                                   # baris ini sudah punya rumus
        for tc in tcols:                                        # buang sel kosong lama di kolom itu bila ada
            inner = re.sub(r'<c r="%s%d"[^>]*?(?:/>|></c>)' % (tc, r), "", inner)
        cells = "".join(tail_cell(tc, lab, r) for tc, lab in zip(tcols, labels))
        added_rows += 1
        return f'<row r="{r}"{attrs}>{inner}{cells}</row>'
    xml = ROW_RE.sub(add_tail, xml)
    if added_rows:
        notes.append(f"Community: rumus TOTAL BULAN per creator ({tcols[0]}:{tcols[-1]}) ditambahkan ke {added_rows:,} baris.")

    # (c) baris TOTAL: jangkauan SUM -> seluruh daftar creator
    mt = re.search(r'<row r="%d"[^>]*>.*?</row>' % total, xml, re.S)
    trow = re.sub(r'<f>SUM\(([A-Z]+)\d+:([A-Z]+)\d+\)</f>', lambda m: f'<f>SUM({m.group(1)}{first}:{m.group(2)}{last})</f>', mt.group(0))
    xml = xml[:mt.start()] + trow + xml[mt.end():]
    patched[p] = xml
    notes.append(f"Community: jumlah harian, baris TOTAL, dan Summary kini mencakup baris {first}-{last}.")

    # ---------- paksa hitung ulang + buang calcChain ----------
    wbxml = zin.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wbxml:
        wbxml = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wbxml, count=1)
    patched["xl/workbook.xml"] = wbxml
    drop = set()
    if "xl/calcChain.xml" in zin.namelist():
        drop.add("xl/calcChain.xml")
        rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        patched["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship\b[^>]*calcChain[^>]*/>', "", rels)
        ct = zin.read("[Content_Types].xml").decode("utf-8")
        patched["[Content_Types].xml"] = re.sub(r'<Override\b[^>]*calcChain[^>]*/>', "", ct)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in drop:
                continue
            data = patched[item.filename] if item.filename in patched else zin.read(item.filename)
            zout.writestr(item, data.encode("utf-8") if isinstance(data, str) else data)
    return buf.getvalue(), notes


def run_fix_mode():
    print("MODE PERBAIKAN TEMPLATE")
    drive_service, _ = get_services()
    tmeta = drive_service.files().get(fileId=TRACKER_FILE_ID, fields="name,mimeType", supportsAllDrives=True).execute()
    if tmeta.get("mimeType") == GSHEET_MIME:
        print("  Tracker ini Google Sheets asli; mode perbaikan template hanya untuk file Excel (.xlsx). "
              "Perbaiki rumus langsung di Google Sheets.")
        return
    data = download_bytes(drive_service, TRACKER_FILE_ID)
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bpath = os.path.join(BACKUP_DIR, f"tracker_sebelum_perbaikan_{stamp}.xlsx")
    with open(bpath, "wb") as fh:
        fh.write(data)
    print(f"  Backup: {os.path.abspath(bpath)}")
    if not template_needs_fix(data):
        print("Template ini sudah dalam kondisi benar (Summary PIC sudah membaca tab yang tepat). Tidak ada yang diubah.")
        return
    try:
        new_bytes, notes = fix_template(data)
    except (KeyError, SheetNotFound) as e:
        print(f"  Tidak bisa memperbaiki otomatis: {e}")
        return
    for n in notes:
        print(f"  - {n}")
    out = f"Tracker_diperbaiki_{stamp}.xlsx"
    with open(out, "wb") as fh:
        fh.write(new_bytes)
    print(f"  Hasil disimpan lokal: {os.path.abspath(out)}")
    print("Upload ke Drive (menimpa file tracker)?  y = ya | n = tidak (pakai file lokal saja)")
    if input("> ").strip().lower() == "y":
        upload_over(drive_service, TRACKER_FILE_ID, out)
        print("  Selesai. Tracker di Drive sudah diperbaiki.")
    else:
        print("  Tidak diunggah.")


# ------------------------------ main ----------------------------

def main():
    global _TRACKER_MTIME
    print("Menghubungkan ke Google Drive & Sheets ...")
    cache_cleanup()
    drive_service, sheets_service = get_services()

    print("Mencari file di folder Drive (Creator_List dan MCN) ...")
    entries = [dict(f, kind="creator") for f in list_candidate_files(drive_service)]
    try:
        entries += [dict(f, kind="mcn") for f in list_mcn_files(drive_service)]
    except Exception as e:
        print(f"  [PERINGATAN] Folder MCN tidak bisa dibuka ({_short_err(e)}) -- file MCN tidak ditampilkan.")
    if not entries:
        print("Tidak ada file 'Creator_List' maupun 'Agency list' di folder itu.")
        return
    chosen = prompt_user_selection(entries, on_preview=lambda ents: print_gmv_preview(drive_service, ents))
    todo, todo_mcn = [], []
    for f in chosen:
        d = entry_date(f)
        if d is None:
            print(f"  [SKIP] '{f['name']}': nama/tanggal tidak dikenali (harus file 1 hari: tanggal awal = tanggal akhir).")
            continue
        (todo_mcn if f["kind"] == "mcn" else todo).append((d, f))
    if not todo and not todo_mcn:
        print("Tidak ada file valid yang dipilih. Selesai.")
        return
    todo.sort(key=lambda t: t[0])
    todo_mcn.sort(key=lambda t: t[0])
    todo, dropped = dedupe_exports(todo)
    for d, f in dropped:
        print(f"  [DIABAIKAN] '{f['name']}': ada export lebih baru untuk tanggal {d:%d-%b-%Y} dengan laporan yang sama "
              f"(kalau dipakai dua-duanya, GMV-nya terhitung dobel).")
    by_day = {}
    for d, f in todo:
        by_day.setdefault(d, []).append(f["name"])
    for d, names in by_day.items():
        if len(names) > 1:
            print(f"  [CATATAN] Tanggal {d:%d-%b-%Y} memakai {len(names)} laporan berbeda yang DIJUMLAHKAN: " + " + ".join(names))

    print("Mengunduh tracker dari Drive ...")
    tmeta = drive_service.files().get(fileId=TRACKER_FILE_ID, fields="name,mimeType,modifiedTime", supportsAllDrives=True).execute()
    is_native = tmeta.get("mimeType") == GSHEET_MIME
    _TRACKER_MTIME = tmeta.get("modifiedTime") if is_native else None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    if is_native:
        print(f"  '{tmeta.get('name')}' adalah Google Sheets asli -> dibaca & ditulis langsung lewat Sheets API "
              f"(tidak lewat ekspor, karena file terlalu besar).")
        tracker_bytes = None
        print("Membaca struktur tracker ...")
        month_start, layouts = analyze_native(sheets_service)
    else:
        tracker_bytes = download_bytes(drive_service, TRACKER_FILE_ID)
        backup_path = os.path.join(BACKUP_DIR, f"tracker_sebelum_{stamp}.xlsx")
        with open(backup_path, "wb") as fh:
            fh.write(tracker_bytes)
        print(f"  Backup tracker (sebelum diubah) disimpan: {os.path.abspath(backup_path)}")
        if template_needs_fix(tracker_bytes):
            print("  [PERINGATAN] Template tracker masih versi lama (tab KLUB VIP/Community tertukar di Summary PIC, "
                  "rumus Community hanya sampai baris 31). Jalankan dulu:  py sync_nivea_tracker.py --perbaiki-template")
        print("Membaca struktur tracker ...")
        month_start, layouts = analyze_tracker(tracker_bytes)
    print(f"  Awal bulan di Setting: {month_start.strftime('%d-%b-%Y')}")
    for tab, lay in layouts.items():
        print(f"  - {tab}: {len(lay['exact'])} creator, {len(lay['blocks'])} blok tanggal")

    src_by_date = {}
    failed = []
    for d, f in todo:
        print(f"Membaca '{f['name']}' ...")
        try:
            data0, fstats = read_creator_file(drive_service, f)
            data = {k: dict(v) for k, v in data0.items()}
        except Exception as e:
            print(f"  [ERROR] {type(e).__name__}: {e}")
            failed.append(f["name"])
            continue
        if d in src_by_date:  # dua file di tanggal sama -> gabungkan
            for k2, rec in data.items():
                cur = src_by_date[d].setdefault(k2, dict(rec, total=0.0, gmv_ls=0.0, gmv_sv=0.0, gmv_sl=0.0))
                for fld in ("total", "gmv_ls", "gmv_sv", "gmv_sl"):
                    cur[fld] += rec[fld]
        else:
            src_by_date[d] = data
        print(f"  {len(data)} creator, total GMV Rp {sum(r['total'] for r in data.values()):,.0f}")
        for n_ in describe_creator_stats(fstats, sum(r["total"] for r in data0.values())):
            print(f"     ! {n_}")
    mcn_by_date, mcn_notes = {}, []
    for d, f in todo_mcn:
        print(f"Membaca file MCN '{f['name']}' ...")
        try:
            data0, nts = read_mcn_file(drive_service, f)
            data = {k: dict(v) for k, v in data0.items()}
        except Exception as e:
            print(f"  [ERROR] {type(e).__name__}: {e}")
            failed.append(f["name"])
            continue
        mcn_notes += [n for n in nts if n not in mcn_notes]
        if d in mcn_by_date:
            for k2, rec in data.items():
                cur = mcn_by_date[d].setdefault(k2, dict(rec, gmv=0.0, ls=None if rec["ls"] is None else 0.0,
                                                         sv=None if rec["sv"] is None else 0.0))
                cur["gmv"] += rec["gmv"]
                for fld in ("ls", "sv"):
                    if rec[fld] is not None:
                        cur[fld] = (cur[fld] or 0.0) + rec[fld]
        else:
            mcn_by_date[d] = data
        print(f"  {len(data)} agency, total GMV Rp {sum(r['gmv'] for r in data.values()):,.0f}")
    if failed:
        print()
        print(f"!! {len(failed)} file gagal dibaca (koneksi?): " + "; ".join(failed[:5]) + (" ..." if len(failed) > 5 else ""))
        print("   Tracker TIDAK diubah. Jalankan skrip lagi dengan pilihan yang sama: file yang sudah terbaca tersimpan "
              f"di folder '{CACHE_DIR}' dan tidak diunduh ulang, jadi hanya file yang gagal yang diunduh lagi.")
        return
    if not src_by_date and not mcn_by_date:
        print("Tidak ada data yang bisa dibaca. Selesai.")
        return

    counts, notes = {"LS": {}, "SV": {}}, []
    if src_by_date:
        print("Mengambil jumlah LS / SV dari Google Sheet ...")
        try:
            counts, notes = load_ls_sv_counts(sheets_service, set(src_by_date), drive_service)
        except Exception as e:
            notes.append(f"Gagal membuka file sumber LS/SV: {_short_err(e)}")
            notes.append(f"Cek: (1) akun yang login bisa membuka https://docs.google.com/spreadsheets/d/{LS_SV_SPREADSHEET_ID} "
                         f"di browser, (2) tab-nya bernama 'LS' dan 'SV'.")
        # kalau satu jenis tidak punya data sama sekali, jangan sentuh kolomnya
        for kind in ("LS", "SV"):
            if not counts[kind]:
                counts[kind] = None
        for n in notes:
            print(f"  [CATATAN] {n}")

    plan = build_plan(month_start, layouts, src_by_date,
                      {kind: (val or {}) for kind, val in counts.items()})
    if mcn_by_date:
        build_mcn_plan(plan, month_start, layouts, mcn_by_date)
    build_mcn_fallback(plan, month_start, layouts, mcn_by_date, src_by_date,
                       {kind: (val or {}) for kind, val in counts.items()} if src_by_date else None)
    st = plan["stats"]

    print()
    print_preview(src_by_date, mcn_by_date, st)
    try:
        pv_path = os.path.abspath(f"Pratinjau_isi_tracker_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx")
        save_preview_xlsx(src_by_date, mcn_by_date, st, pv_path)
        print(f"Pratinjau lengkap (semua creator, bisa difilter) disimpan di:\n  {pv_path}")
    except Exception as e:
        print(f"(pratinjau Excel tidak bisa disimpan: {type(e).__name__}: {e})")
    print()
    print("========== RINGKASAN YANG AKAN DITULIS ==========")
    for tab, n in st["per_tab"].items():
        print(f"  {tab:<10}: {n} creator-hari terisi")
    if src_by_date:
        print(f"  Creator baru ke {COMMUNITY_TAB} (GMV > 0): {st['added']}")
        print("  Cek angka GMV Total (Creator_List) -- file vs yang ditulis ke tracker:")
        for d_, rc in sorted(st.get("recon", {}).items()):
            gap = rc["file"] - rc["written"]
            flag = "  !! tertulis LEBIH BESAR dari file -> kemungkinan dobel, jangan di-upload dulu" if gap < -0.5 else ""
            print(f"     {d_:%d-%b-%Y}: file Rp {rc['file']:,.0f} | tertulis Rp {rc['written']:,.0f} | "
                  f"tidak tertulis Rp {max(gap, 0):,.0f} (creator belum terdaftar & di bawah batas){flag}")
    mst = st.get("mcn")
    if mst:
        if mst.get("no_tab"):
            print(f"  !! Tab '{MCN_TAB}' tidak ditemukan di tracker -- data MCN tidak ditulis.")
        else:
            print(f"  MCN: {mst['filled']} agency-hari terisi (GMV Total, LS, SV)")
            for d_, rc in sorted(mst.get("recon", {}).items()):
                gap = rc["file"] - rc["written"]
                print(f"     {d_:%d-%b-%Y}: file Rp {rc['file']:,.0f} | tertulis Rp {rc['written']:,.0f} | tidak tertulis Rp {max(gap, 0):,.0f} (agency tidak ada di tab MCN)")
            print("     (GMV agency umumnya sudah mencakup GMV creator di bawahnya -- jangan dijumlahkan dengan tab creator.)")
            for n in mcn_notes:
                print(f"     - {n}")
            if mst["loose"]:
                print(f"     Dicocokkan 'longgar': " + "; ".join(f"{a} -> {b}" for a, b in mst["loose"][:5]))
            if mst["dup"]:
                print(f"     {len(mst['dup'])} nama dobel di tab MCN -> hanya baris pertama yang diisi.")
            if mst["unmatched"]:
                top = sorted(mst["unmatched"].items(), key=lambda x: -x[1])
                print(f"     {len(top)} agency di file TIDAK ada di tab MCN (tidak ditulis), GMV terbesar:")
                for nm, g in top[:6]:
                    print(f"        - {nm}  (Rp {g:,.0f})")
    fbk = (mst or {}).get("fallback") if mst else None
    if fbk and (fbk["rows"] or fbk["in_other_tab"]):
        n_fb = sum(len(v) for v in fbk["rows"].values())
        print(f"  MCN cadangan: {n_fb} nama di tab MCN tidak ada di file Agency -> diisi dari Creator_List (folder Drive utama).")
        for d_, tot in sorted(fbk["total"].items()):
            print(f"     {d_:%d-%b-%Y}: Rp {tot:,.0f} dari {len(fbk['rows'][d_])} nama")
        if fbk["in_other_tab"]:
            print(f"     {len(fbk['in_other_tab'])} nama MCN dilewati karena sudah ada di tab creator lain (supaya tidak dobel): "
                  + "; ".join(f"{a} ({b})" for a, b in sorted(fbk['in_other_tab'])[:6]))
    if st["loose"]:
        print(f"  Dicocokkan 'longgar' (beda tanda baca/spasi, kandidat tunggal): {len(st['loose'])} nama. Contoh:")
        for a, b, t in st["loose"][:8]:
            print(f"      file: {a}   ->  tracker ({t}): {b}")
    if st["cross_tab"]:
        print(f"  {len(st['cross_tab'])} creator tercatat di lebih dari satu tab -> hanya diisi di tab prioritas "
              f"(urutan: {', '.join(ROSTER_TABS)}).")
    if st["dup_rows"]:
        print(f"  {len(st['dup_rows'])} nama dobel dalam satu tab -> hanya baris pertama yang diisi.")
    skipped = list(st["skipped_dates"]) + (mst["skipped_dates"] if mst else [])
    if skipped:
        print(f"  Tanggal di luar bulan tracker (dilewati): {sorted({d.isoformat() for d in skipped})}")
    if src_by_date and (counts["LS"] is None or counts["SV"] is None):
        print("  !! Kolom LS/SV belum terisi karena sumbernya belum terbaca. Penyebab:")
        for n in notes:
            print(f"     - {n}")
    print("==================================================")
    print("Lanjut?  y = simpan & upload ke Drive | s = simpan lokal saja (tidak upload) | n = batal")
    ans = input("> ").strip().lower()
    if ans not in ("y", "s"):
        print("Dibatalkan. Tidak ada yang diubah.")
        return

    if is_native:
        if ans != "y":
            print("  Mode simpan lokal: untuk Google Sheets asli tidak ada file hasil. Pratinjau lengkap sudah disimpan di atas; "
                  "tracker TIDAK diubah. Jalankan lagi dan pilih y untuk menulis.")
            return
        print("Membaca sel yang akan diubah (untuk backup & pengecekan rumus) ...")
        info = _native_template_info(sheets_service, plan, layouts)
        import json as _json
        bk = os.path.join(BACKUP_DIR, f"nilai_lama_sebelum_{stamp}.json")
        with open(bk, "w", encoding="utf-8") as fh:
            _json.dump({t: {f"{idx_to_col(c)}{r}": v for (r, c), v in i["old"].items()} for t, i in info.items()},
                       fh, ensure_ascii=False)
        print(f"  Backup nilai lama disimpan: {os.path.abspath(bk)} (riwayat versi Google Sheets juga bisa dipakai untuk kembali).")
        print("Menulis ke Google Sheets ...")
        _TRACKER_MTIME = None          # setelah ini isi tracker berubah -> jangan pakai hasil baca tersimpan lagi
        n_cells, shift = push_to_native(sheets_service, plan, layouts, info)
        print(f"  {n_cells} sel ditulis.")
        try:
            verify_written(None, plan, layouts, src_by_date, mcn_by_date, month_start,
                           col_sum_fn=make_native_col_sum(sheets_service, shift))
        except Exception as e:
            print(f"  (cek ulang dari Google Sheets gagal: {type(e).__name__}: {e})")
        print("  Selesai. Tracker di Google Sheets sudah diperbarui.")
        return

    print("Mengisi tracker ...")
    new_bytes = apply_plan(tracker_bytes, plan, layouts)
    try:
        ok_v = verify_written(new_bytes, plan, layouts, src_by_date, mcn_by_date, month_start)
    except Exception as e:
        ok_v = None
        print(f"  (cek ulang gagal: {type(e).__name__}: {e})")
    out_path = f"Tracker_hasil_{stamp}.xlsx"
    with open(out_path, "wb") as fh:
        fh.write(new_bytes)
    print(f"  Hasil disimpan lokal: {os.path.abspath(out_path)}")

    if ans == "y":
        print("Mengunggah ke Drive (menimpa file tracker) ...")
        upload_over(drive_service, TRACKER_FILE_ID, out_path)
        print("  Selesai. Tracker di Drive sudah diperbarui.")
    else:
        print("  Tidak diunggah (mode simpan lokal).")


def run_preview_mode():
    print("MODE LIHAT TOTAL GMV (tracker tidak disentuh)")
    drive_service, _ = get_services()
    entries = [dict(f, kind="creator") for f in list_candidate_files(drive_service)]
    try:
        entries += [dict(f, kind="mcn") for f in list_mcn_files(drive_service)]
    except Exception as e:
        print(f"  [PERINGATAN] Folder MCN tidak bisa dibuka ({_short_err(e)}).")
    if not entries:
        print("Tidak ada file di kedua folder.")
        return
    print_gmv_preview(drive_service, entries)


if __name__ == "__main__":
    if "--lihat-gmv" in sys.argv:
        sys.exit(run_preview_mode())
    if "--perbaiki-template" in sys.argv:
        sys.exit(run_fix_mode())
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nDihentikan.")
        sys.exit(1)
    except Exception as e:
        if is_net_error(e):
            print()
            print(f"!! Koneksi terputus ({type(e).__name__}). Tidak apa-apa -- jalankan skrip lagi dengan pilihan yang sama.")
            print(f"   File yang sudah diunduh/dibaca tersimpan di '{CACHE_DIR}' (tidak diunduh ulang), dan sel tracker yang sudah "
                  f"tertulis tidak ditulis ulang, jadi proses melanjutkan dari sisa yang belum selesai.")
            sys.exit(1)
        raise
