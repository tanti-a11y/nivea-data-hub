#!/usr/bin/env python3
"""
Otomasi Daily Tracker Content NIVEA: tab SV (video) dan LS (live).

Alur (sama untuk SV dan LS):
  1. Ambil semua file export harian dari folder Google Drive (atau folder lokal).
  2. Pilih rentang tanggal (per bulan atau tanggal bebas).
  3. Kalau satu Video ID / Live ID muncul di beberapa file, pakai data dari file TERBARU.
  4. ID baru: isi PID, Product Name, CAT, PRO, Tiering, lalu append ke tab tracker.
     - SV: PID & Product Name = SEMUA produk di kolom "Products", digabung dengan pemisah
       (default " | ", bisa diubah lewat config "multi_separator"). CAT & PRO tetap dari produk pertama.
       Tab "Exploded SV" (config "sv_exploded_sheet"): tabel yang sama, tapi satu baris per PID
       per Video ID. Metrik sama di setiap baris, hanya PID / Product Name / CAT / PRO yang beda.
     - LS: PID pertama dari kolom "Product ID", Product Name dari Product Bank. Duration = End - Start.
  5. PID yang belum ada CAT/PRO-nya ditanyakan, lalu otomatis ditambahkan ke Product Bank.
  6. ID yang sudah ada: metrik diperbarui dari file terbaru (tidak ada ID dobel).

Contoh:
  python nivea_tracker.py                                   # pilih SV/LS dan bulan secara interaktif
  python nivea_tracker.py --mode ls --month 2026-10         # LS satu bulan
  python nivea_tracker.py --mode sv --start 2026-10-01 --end 2026-10-07
  python nivea_tracker.py --mode all --month current --yes  # SV + LS, tanpa tanya (jadwal otomatis)
  python nivea_tracker.py --mode ls --month 2026-10 --dry-run   # cek dulu, tidak menulis apa pun
"""

import argparse
import calendar
import csv
import io
import json
import os
import re
import sys
from datetime import datetime, date, time as dtime, timedelta

import pandas as pd

MONTH_ABBR = ["", "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
              "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",  # perlu tulis untuk update Product Bank
]

GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# --------------------------------------------------------------------------
# Definisi tracker SV dan LS
# --------------------------------------------------------------------------
# "fields": kolom internal -> daftar kemungkinan nama kolom di file export (dicoba berurutan).
# Kolom metrik memakai nama kolom tracker sebagai kunci. Bisa ditimpa lewat config
# ("sv_column_map" / "ls_column_map"), misalnya {"GMV": "Video GMV (Rp)"}.

FMT_DATE = ("DATE", "dd-mmm-yyyy")
FMT_DATETIME = ("DATE_TIME", "mm/dd/yyyy h:mm")
FMT_DURATION = ("TIME", "[h]:mm:ss")
FMT_ACC = ("NUMBER", "#,##0;\\(#,##0\\);\\-")
FMT_INT = ("NUMBER", "#,##0")
FMT_PCT = ("PERCENT", "0.00%")

SV = {
    "key": "sv", "label": "SV (Video)", "unit": "video",
    "sheet_cfg": ("tracker_sheet", "SV"), "header_cfg": ("header_row", 2),
    "folder_cfg": "daily_folder_id", "map_cfg": "sv_column_map",
    "columns": ["Date", "Upload Datetime", "Week", "Video ID", "Hastag", "Creator", "Tiering", "PID",
                "Product Name", "CAT", "PRO", "GMV", "Video views", "Likes", "Comments", "Shares",
                "CTR (Video) %", "V-to-L Rate %", "V-to-L Clicks", "Video Finish Rate %",
                "CTOR (SKU Order) %"],
    "id_col": "Video ID", "first_metric": "GMV", "date_order": "dmy",
    "fields": {
        "_id": ["Video ID"],
        "_time": ["Time", "Post time", "Publish time"],
        "_creator": ["Creator name", "Creator"],
        "_products": ["Products"],
        "_video_info": ["Video Info", "Video info", "Video name", "Video title"],
        "GMV": ["Video-attributed GMV (Rp)"],
        "Video views": ["Video views", "Video Views", "Views", "Video plays"],
        "Likes": ["Likes"],
        "Comments": ["Comments"],
        "Shares": ["Shares"],
        "CTR (Video) %": ["Click-through rate (Video)"],
        "V-to-L Rate %": ["V-to-L rate"],
        "V-to-L Clicks": ["V-to-L clicks"],
        "Video Finish Rate %": ["Video Finish Rate"],
        "CTOR (SKU Order) %": ["CTOR (SKU order)"],
    },
    "formats": {"Date": FMT_DATE, "Upload Datetime": FMT_DATE, "GMV": FMT_ACC, "Video views": FMT_INT, "Likes": FMT_INT,
                "Comments": FMT_INT, "Shares": FMT_INT, "V-to-L Clicks": FMT_INT,
                "CTR (Video) %": FMT_PCT, "V-to-L Rate %": FMT_PCT, "Video Finish Rate %": FMT_PCT,
                "CTOR (SKU Order) %": FMT_PCT},
}

LS = {
    "key": "ls", "label": "LS (Live)", "unit": "live",
    "sheet_cfg": ("ls_sheet", "LS"), "header_cfg": ("ls_header_row", 1),
    "folder_cfg": "ls_folder_id", "map_cfg": "ls_column_map",
    "columns": ["Date", "Week", "Live Title", "Live ID", "Creator", "Tiering", "Start Time",
                "End Time", "Duration", "PID", "Product Name", "CAT", "PRO", "GMV", "Items Sold",
                "Refunds", "Orders", "AOV", "Avg Viewing Duration (sec)", "Likes", "Comments",
                "Shares", "Product Impressions", "Product Clicks", "Impressions", "Show GPM",
                "Products Sold", "Avg GMV per Customer", "Engagement %", "CTR %",
                "LIVE Product Viewers", "Tap-through Rate %", "Est Commission"],
    "id_col": "Live ID", "first_metric": "GMV", "date_order": "mdy",
    "fields": {
        "_id": ["Live ID", "LIVE ID", "LIVE room ID", "Live room ID", "Room ID", "LIVE session ID", "ID"],
        "_time": ["LIVE start time", "Start time", "Start Time", "Launched time", "Start"],
        "_end": ["LIVE end time", "End time", "End Time", "End"],
        "_duration": ["Duration", "LIVE duration"],
        "_creator": ["Creator name", "Creator", "Creator username", "Username", "Nickname"],
        "_title": ["LIVE title", "Live Title", "LIVE title", "LIVE name", "Live name", "Title", "LIVE room name"],
        "_pid_raw": ["Product ID", "Product IDs", "Product id", "PID", "Products"],
        "_pname": ["Product name", "Product Name", "Product names"],
        "GMV": ["Creator LIVE-attributed GMV", "GMV", "GMV (Rp)", "LIVE GMV", "Gross merchandise value", "Gross revenue"],
        "Items Sold": ["LIVE-attributed items sold", "Items sold", "Item sold"],
        "Refunds": ["Refunds", "Refund", "Refunded GMV", "Refund amount"],
        "Orders": ["LIVE-attributed orders", "Orders", "SKU orders", "Created SKU orders", "Paid orders"],
        "AOV": ["AOV", "Average order value", "Avg. order value"],
        "Avg Viewing Duration (sec)": ["Avg viewing duration", "Avg. viewing duration", "Average viewing duration",
                                       "Avg. viewing duration per view"],
        "Likes": ["Likes"],
        "Comments": ["Comments"],
        "Shares": ["Shares"],
        "Product Impressions": ["LIVE product impressions", "Product impressions"],
        "Product Clicks": ["Product clicks"],
        "Impressions": ["Impressions", "LIVE impressions", "Viewers", "Views"],
        "Show GPM": ["Avg. daily GPM", "Show GPM", "GPM", "Watch GPM"],
        "Products Sold": ["Avg. daily products sold", "Products sold", "Product sold"],
        "Avg GMV per Customer": ["Avg GMV per customer", "Avg. GMV per customer", "Average GMV per customer",
                                 "Avg. GMV per buyer"],
        "Engagement %": ["Avg. daily engagement", "Engagement rate", "Engagement", "Engagement %"],
        "CTR %": ["Avg. daily LIVE product CTR", "CTR", "Click-through rate", "Product CTR"],
        "LIVE Product Viewers": ["Avg. daily unique viewers", "LIVE product viewers", "Product viewers"],
        "Tap-through Rate %": ["Avg. daily tap-through rate", "Tap-through rate", "Tap through rate", "LIVE tap-through rate"],
        "Est Commission": ["Est. commission", "Est commission", "Estimated commission"],
    },
    "formats": {"Date": FMT_DATE, "Start Time": FMT_DATETIME, "End Time": FMT_DATETIME,
                "Duration": FMT_DURATION, "GMV": FMT_ACC, "Items Sold": FMT_INT, "Refunds": FMT_INT,
                "Orders": FMT_INT, "AOV": FMT_ACC, "Avg Viewing Duration (sec)": FMT_INT, "Likes": FMT_INT,
                "Comments": FMT_INT, "Shares": FMT_INT, "Product Impressions": FMT_INT,
                "Product Clicks": FMT_INT, "Impressions": FMT_INT, "Show GPM": FMT_INT,
                "Products Sold": FMT_INT, "Avg GMV per Customer": FMT_ACC, "Engagement %": FMT_PCT,
                "CTR %": FMT_PCT, "LIVE Product Viewers": FMT_INT, "Tap-through Rate %": FMT_PCT,
                "Est Commission": FMT_ACC},
}

SPECS = {"sv": SV, "ls": LS}


def metric_cols(spec):
    return spec["columns"][spec["columns"].index(spec["first_metric"]):]


def col_letter(n):
    """1 -> A, 27 -> AA."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# --------------------------------------------------------------------------
# Helper parsing
# --------------------------------------------------------------------------

def clean_id(value):
    """Ubah ID (PID / Video ID) jadi string angka utuh, aman dari format float/scientific."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip().lstrip("'")
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s


def clean_name(value):
    """Normalisasi nama creator untuk pencocokan (huruf kecil, tanpa @ dan spasi)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip().lstrip("@").lower()


_PID_RE = re.compile(r"\((\d{10,})\)")


def parse_first_product(products):
    """Ambil (PID, Product Name) dari produk pertama di kolom Products.

    Contoh: "NIVEA Lip Balm ... Shea Butter(1732010205287712512),NIVEA Lip Balm ...(1732...)"
            -> ("1732010205287712512", "NIVEA Lip Balm ... Shea Butter")
    """
    if products is None or (isinstance(products, float) and pd.isna(products)):
        return "", ""
    text = str(products).strip()
    m = _PID_RE.search(text)
    if not m:
        return "", text
    return m.group(1), text[:m.start()].strip().strip(",").strip()


def parse_all_products(products):
    """Semua (PID, Product Name) dari kolom Products, urut seperti di export, PID dobel dibuang.

    Contoh: "Nama A(1732...512),Nama B, 50ml(1738...456)"
            -> [("1732...512", "Nama A"), ("1738...456", "Nama B, 50ml")]
    """
    if products is None or (isinstance(products, float) and pd.isna(products)):
        return []
    text = str(products).strip()
    out, seen, prev = [], set(), 0
    for m in _PID_RE.finditer(text):
        pid = m.group(1)
        name = text[prev:m.start()].strip().lstrip(",").strip()
        prev = m.end()
        if pid in seen:
            continue
        seen.add(pid)
        out.append((pid, name))
    if not out and text:
        return [("", text)]
    return out


_HASHTAG_RE = re.compile(r"#[^\s#]+")


def extract_hashtags(text):
    """Ambil semua #hashtag dari teks Video Info (urut, tanpa dobel), dipisah spasi."""
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    seen, out = set(), []
    for tag in _HASHTAG_RE.findall(str(text)):
        tag = tag.rstrip(".,;:!?)\"'")
        if len(tag) > 1 and tag.lower() not in seen:
            seen.add(tag.lower())
            out.append(tag)
    return " ".join(out)


def products_of(rec, spec, bank):
    """Return list (PID, Product Name). SV: semua produk. LS: satu produk (PID pertama)."""
    if spec["key"] == "sv":
        return parse_all_products(rec.get("_products")) or [("", "")]
    return [product_of(rec, spec, bank)]


def join_products(prods, sep):
    """Gabungkan list produk jadi (teks PID, teks Product Name)."""
    return sep.join(p for p, _ in prods if p), sep.join(n for _, n in prods if n)


def parse_number(value):
    """Angka dari teks export: '1,234', '1.234.567', 'Rp12.345', '12.5K', dll."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip().replace("Rp", "").replace("IDR", "").replace(" ", "").replace(" ", "")
    if s in ("", "-", "--", "N/A", "nan"):
        return ""
    mult = 1
    if s[-1:].upper() in ("K", "M", "B") and re.match(r"^-?[\d.,]+[KMB]$", s, re.I):
        mult = {"K": 1e3, "M": 1e6, "B": 1e9}[s[-1].upper()]
        s = s[:-1]
    if "," in s and "." in s:
        # pemisah desimal = yang muncul paling akhir
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        s = s.replace(",", "") if all(len(p) == 3 for p in parts[1:]) else s.replace(",", ".")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[1]) == 3 and mult == 1):
        s = s.replace(".", "")
    try:
        n = float(s) * mult
    except ValueError:
        return ""
    return int(n) if n.is_integer() else n


def parse_percent(value):
    """Persen jadi pecahan (5.23% -> 0.0523) supaya cocok dengan format 0.00% di sheet."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (int, float)):
        # angka mentah dari Excel: 0.0523 sudah pecahan, 5.23 berarti persen
        return value / 100 if value > 1 else value
    s = str(value).strip().replace(" ", "").replace(" ", "")
    had_pct = s.endswith("%")
    s = s.rstrip("%").replace(",", ".")
    try:
        n = float(s)
    except ValueError:
        return ""
    return n / 100 if had_pct or n > 1 else n


_SLASH_DATE_RE = re.compile(r"^\s*(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})")


def detect_date_order(values, default="mdy"):
    """Tebak urutan tanggal untuk teks seperti 10/02/2026: 'mdy' (bulan/hari) atau 'dmy' (hari/bulan).
    Kalau ada bagian pertama > 12 berarti dmy, kalau bagian kedua > 12 berarti mdy."""
    for v in values:
        m = _SLASH_DATE_RE.match(str(v)) if isinstance(v, str) else None
        if not m:
            continue
        a, b = int(m.group(1)), int(m.group(2))
        if a > 12:
            return "dmy"
        if b > 12:
            return "mdy"
    return default


def parse_time(value, order="dmy"):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, (int, float)):  # serial date Excel
        return datetime(1899, 12, 30) + pd.Timedelta(days=float(value))
    s = str(value).strip()
    if order == "mdy":
        local = ("%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%m/%d/%Y", "%m-%d-%Y %H:%M:%S", "%m-%d-%Y %H:%M", "%m-%d-%Y")
    else:
        local = ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d-%m-%Y")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S",
                "%Y/%m/%d %H:%M", "%Y/%m/%d") + local + ("%d-%b-%Y", "%d %b %Y %H:%M", "%d %b %Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    try:
        return pd.to_datetime(s, dayfirst=(order != "mdy")).to_pydatetime()
    except (ValueError, TypeError):
        return None


def week_label(d):
    """1-7 = W1, 8-14 = W2, 15-21 = W3, 22-akhir bulan = W4. Format: 'OCT W1 (1-7)'."""
    last = calendar.monthrange(d.year, d.month)[1]
    ranges = [(1, 7), (8, 14), (15, 21), (22, last)]
    for i, (a, b) in enumerate(ranges, 1):
        if a <= d.day <= b:
            return f"{MONTH_ABBR[d.month]} W{i} ({a}-{b})"


def as_text(s):
    """Paksa jadi teks di Google Sheets (ID panjang tidak berubah jadi 1.73E+18)."""
    s = "" if s is None else str(s)
    return "'" + s if s else ""


def safe_text(s):
    s = "" if s is None else str(s)
    return "'" + s if s[:1] in ("=", "+", "-", "@") else s


# --------------------------------------------------------------------------
# Baca tabel (export harian, Product Bank, Tiering)
# --------------------------------------------------------------------------

def find_header_and_frame(raw, must_have):
    """raw = DataFrame tanpa header. Cari baris yang berisi kolom `must_have` lalu jadikan header."""
    want = must_have.strip().lower()
    for i in range(min(len(raw), 30)):
        row = [str(x).strip().lower() for x in raw.iloc[i].tolist()]
        if want in row:
            df = raw.iloc[i + 1:].copy()
            df.columns = [str(x).strip() for x in raw.iloc[i].tolist()]
            df = df.loc[:, [c for c in df.columns if c and c.lower() != "nan"]]
            return df.dropna(how="all").reset_index(drop=True)
    return None


def read_table_bytes(content, filename, must_have, sheet_name=None):
    """Baca xlsx/xls/csv dari bytes, cari sheet yang punya kolom `must_have`."""
    name = filename.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8-sig", "utf-16", "latin-1"):
            try:
                raw = pd.read_csv(io.BytesIO(content), header=None, dtype=str, encoding=enc,
                                  sep=None, engine="python")
                break
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        else:
            raise ValueError(f"Tidak bisa membaca CSV {filename}")
        sheets = {"csv": raw}
    else:
        sheets = pd.read_excel(io.BytesIO(content), header=None, dtype=object, sheet_name=None)
    order = list(sheets)
    if sheet_name and sheet_name in sheets:
        order = [sheet_name] + [s for s in order if s != sheet_name]
    for s in order:
        df = find_header_and_frame(sheets[s], must_have)
        if df is not None:
            return df
    raise ValueError(f"Kolom '{must_have}' tidak ditemukan di {filename}")


# --------------------------------------------------------------------------
# Google API
# --------------------------------------------------------------------------

def load_credentials(credentials_file, token_file):
    """Deteksi otomatis jenis file kredensial:
    - OAuth client (Desktop app, isi JSON diawali "installed"): login lewat browser sekali,
      token disimpan di token.json, jadi run berikutnya tidak perlu login lagi.
    - Service account (isi JSON ada "client_email"): pakai akun robot.
    """
    if not os.path.exists(credentials_file):
        sys.exit(f"File kredensial tidak ditemukan: {credentials_file}")
    with open(credentials_file, encoding="utf-8") as f:
        info = json.load(f)

    if "installed" in info or "web" in info:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials as UserCredentials
        from google_auth_oauthlib.flow import InstalledAppFlow

        creds = None
        if os.path.exists(token_file):
            creds = UserCredentials.from_authorized_user_file(token_file)
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
        if creds and not creds.has_scopes(SCOPES):
            creds = None  # izin lama (read-only) -> login ulang sekali
        if not creds or not creds.valid:
            print("   Browser akan terbuka untuk login Google. Pilih akun yang punya akses ke tracker.")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
        return creds

    if info.get("type") == "service_account":
        from google.oauth2.service_account import Credentials
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    sys.exit("Format file kredensial tidak dikenali. Pakai OAuth client ID (Desktop app) "
             "atau key JSON service account dari Google Cloud.")


class Google:
    def __init__(self, credentials_file, token_file):
        import gspread
        from googleapiclient.discovery import build

        creds = load_credentials(credentials_file, token_file)
        self.gc = gspread.authorize(creds)
        self.drive = build("drive", "v3", credentials=creds, cache_discovery=False)

    def list_folder(self, folder_id):
        files, token = [], None
        q = f"'{folder_id}' in parents and trashed = false"
        while True:
            resp = self.drive.files().list(
                q=q, pageToken=token, pageSize=1000,
                fields="nextPageToken, files(id, name, mimeType, modifiedTime)",
                supportsAllDrives=True, includeItemsFromAllDrives=True,
            ).execute()
            files += resp.get("files", [])
            token = resp.get("nextPageToken")
            if not token:
                break
        return files

    def file_meta(self, file_id):
        return self.drive.files().get(fileId=file_id, fields="id, name, mimeType, modifiedTime",
                                      supportsAllDrives=True).execute()

    def download(self, meta):
        """Return (bytes, filename). Google Sheet di-export jadi xlsx."""
        from googleapiclient.http import MediaIoBaseDownload

        if meta["mimeType"] == GOOGLE_SHEET_MIME:
            req = self.drive.files().export_media(fileId=meta["id"], mimeType=XLSX_MIME)
            filename = meta["name"] + ".xlsx"
        else:
            req = self.drive.files().get_media(fileId=meta["id"], supportsAllDrives=True)
            filename = meta["name"]
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        return buf.getvalue(), filename

    def upload_xlsx(self, file_id, content):
        from googleapiclient.http import MediaIoBaseUpload

        media = MediaIoBaseUpload(io.BytesIO(content), mimetype=XLSX_MIME, resumable=False)
        self.drive.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()

    def read_table(self, file_id, must_have, sheet_name=None):
        meta = self.file_meta(file_id)
        content, filename = self.download(meta)
        return read_table_bytes(content, filename, must_have, sheet_name)


# --------------------------------------------------------------------------
# Product Bank & Tiering
# --------------------------------------------------------------------------

def _cell_text(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.startswith("#") else s  # #N/A dsb dianggap kosong


class ProductBank:
    """lookup: PID -> (CAT, PRO) untuk PID yang CAT-nya sudah terisi.
    pids: semua PID yang tercatat (termasuk yang CAT/PRO masih kosong)."""

    def __init__(self, df):
        cols = {c.lower(): c for c in df.columns}
        pid_col = cols.get("product id") or cols.get("pid")
        name_col = cols.get("product name")
        self.lookup, self.pids, self.known_names, self.names = {}, set(), {}, {}
        self.cats, self.pros = [], []
        for _, r in df.iterrows():
            pid = clean_id(r[pid_col])
            if not pid:
                continue
            self.pids.add(pid)
            if name_col and pid not in self.names:
                self.names[pid] = _cell_text(r.get(name_col))
            cat, pro = _cell_text(r.get(cols.get("cat"))), _cell_text(r.get(cols.get("pro")))
            if cat and pid not in self.lookup:
                self.lookup[pid] = (cat, pro)
                name = _cell_text(r.get(name_col)) if name_col else ""
                if name:
                    self.known_names[name] = (cat, pro)
                if cat not in self.cats:
                    self.cats.append(cat)
                if pro and pro not in self.pros:
                    self.pros.append(pro)

    def get(self, pid):
        return self.lookup.get(pid, ("", ""))

    def name(self, pid):
        return self.names.get(pid, "")

    def suggest(self, name):
        """Tebak CAT/PRO dari produk dengan nama paling mirip di Product Bank (berdasarkan kata)."""
        words = _name_words(name)
        best, best_score = None, 0.0
        for known in self.known_names:
            kw = _name_words(known)
            if not words or not kw:
                continue
            score = len(words & kw) / len(words | kw)
            if score > best_score:
                best, best_score = known, score
        if best is None or best_score < 0.25:
            return None, ("", "")
        return best, self.known_names[best]


def _name_words(name):
    """Kata penting dari bagian utama nama produk (sebelum ' - ' pertama)."""
    core = str(name or "").split(" - ")[0].lower()
    return {w for w in re.findall(r"[a-z0-9&]+", core) if w not in ("nivea", "isi", "paket", "hemat", "x", "ml", "gr")}


def build_tier_lookup(df):
    cols = {c.lower(): c for c in df.columns}
    name_col = cols.get("nama creator") or cols.get("creator") or cols.get("creator name")
    tier_col = cols.get("tiering") or cols.get("tier")
    lookup = {}
    for _, r in df.iterrows():
        name = clean_name(r[name_col])
        tier = r[tier_col]
        if name and name not in lookup and tier is not None and not pd.isna(tier):
            lookup[name] = str(tier).strip()
    return lookup


def _pick(label, options, default):
    print(f"     Pilihan {label}: " + "  ".join(f"{i}.{o}" for i, o in enumerate(options, 1)))
    hint = f"Enter = {default}, " if default else ""
    ans = input(f"     {label} [{hint}nomor / ketik baru / '-' = kosongkan]: ").strip()
    if ans == "":
        return default
    if ans == "-":
        return ""
    if ans.isdigit() and 1 <= int(ans) <= len(options):
        return options[int(ans) - 1]
    return ans.upper()


def ask_cat_pro(missing, bank):
    """missing: dict PID -> {"name", "count"}. Return list entri untuk Product Bank."""
    entries = []
    print(f"\n📝 Ada {len(missing)} PID yang belum punya CAT/PRO. Isi sekarang supaya otomatis masuk Product Bank.")
    for i, (pid, info) in enumerate(sorted(missing.items(), key=lambda x: -x[1]["count"]), 1):
        print(f"\n  [{i}/{len(missing)}] PID {pid} ({info['count']} konten)")
        if info["name"]:
            print(f"     {info['name'][:150]}")
        else:
            info["name"] = input("     Nama produk belum ada. Ketik nama produk [Enter = kosongkan]: ").strip()
        similar, (s_cat, s_pro) = bank.suggest(info["name"])
        if similar:
            print(f"     Saran dari produk mirip: {s_cat} / {s_pro or '-'}  ({similar[:70]}...)")
        cat = _pick("CAT", bank.cats, s_cat)
        pro = _pick("PRO", bank.pros, s_pro)
        entries.append({"pid": pid, "name": info["name"], "cat": cat, "pro": pro})
    return entries


def write_product_bank(g, file_id, sheet_name, entries):
    """Tambah PID baru (atau isi CAT/PRO yang kosong) di Product Bank. Mendukung .xlsx dan Google Sheet."""
    import openpyxl
    from copy import copy

    meta = g.file_meta(file_id)
    if meta["mimeType"] == GOOGLE_SHEET_MIME:
        ws = g.gc.open_by_key(file_id).worksheet(sheet_name)
        values = ws.get_all_values()
        hdr = next(i for i, r in enumerate(values) if "Product ID" in r)
        idx = {h: values[hdr].index(h) for h in ("Product ID", "Product Name", "CAT", "PRO")}
        rows = {clean_id(r[idx["Product ID"]]): n for n, r in enumerate(values[hdr + 1:], start=hdr + 2)}
        new, updates = [], []
        for e in entries:
            if e["pid"] in rows:
                r = rows[e["pid"]]
                for key, col in (("cat", "CAT"), ("pro", "PRO")):
                    if e[key]:
                        updates.append({"range": gspread_a1(r, idx[col] + 1), "values": [[e[key]]]})
            else:
                row = [""] * len(values[hdr])
                row[idx["Product ID"]] = "'" + e["pid"]
                row[idx["Product Name"]] = safe_text(e["name"])
                row[idx["CAT"]], row[idx["PRO"]] = e["cat"], e["pro"]
                new.append(row)
        if updates:
            ws.batch_update(updates, value_input_option="USER_ENTERED")
        if new:
            ws.append_rows(new, value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS")
        return

    content, _ = g.download(meta)
    wb = openpyxl.load_workbook(io.BytesIO(content))
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]
    hdr, cols = None, {}
    for r in range(1, 31):
        vals = {str(ws.cell(r, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value}
        if "Product ID" in vals:
            hdr, cols = r, vals
            break
    if hdr is None:
        raise ValueError("Header 'Product ID' tidak ditemukan di Product Bank")
    c_pid, c_name, c_cat, c_pro = (cols["Product ID"], cols.get("Product Name"), cols.get("CAT"), cols.get("PRO"))
    rows, last = {}, hdr
    for r in range(hdr + 1, ws.max_row + 1):
        pid = clean_id(ws.cell(r, c_pid).value)
        if pid:
            rows[pid], last = r, r
    for e in entries:
        if e["pid"] in rows:
            r = rows[e["pid"]]
            if e["name"] and c_name and not ws.cell(r, c_name).value:
                ws.cell(r, c_name).value = e["name"]
            if e["cat"] and c_cat:
                ws.cell(r, c_cat).value = e["cat"]
            if e["pro"] and c_pro:
                ws.cell(r, c_pro).value = e["pro"]
            continue
        last += 1
        for c in range(1, ws.max_column + 1):  # ikut gaya baris di atasnya
            src = ws.cell(last - 1, c)
            if src.has_style:
                ws.cell(last, c)._style = copy(src._style)
        ws.cell(last, c_pid).value = e["pid"]
        ws.cell(last, c_pid).number_format = "@"
        if c_name:
            ws.cell(last, c_name).value = e["name"]
        if c_cat:
            ws.cell(last, c_cat).value = e["cat"] or None
        if c_pro:
            ws.cell(last, c_pro).value = e["pro"] or None
        rows[e["pid"]] = last
    buf = io.BytesIO()
    wb.save(buf)
    g.upload_xlsx(file_id, buf.getvalue())


def gspread_a1(row, col):
    import gspread.utils as gu
    return gu.rowcol_to_a1(row, col)




def parse_seconds(value):
    """Durasi jadi detik: 45, '45s', '1m 23s', '1h 2m', '00:01:23', '01:23'."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, timedelta):
        return value.total_seconds()
    if isinstance(value, dtime):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip().lower()
    if not s or s in ("-", "--"):
        return ""
    if re.fullmatch(r"\d+(:\d{1,2}){1,2}(\.\d+)?", s):
        parts = [float(p) for p in s.split(":")]
        while len(parts) < 3:
            parts.insert(0, 0)
        return parts[0] * 3600 + parts[1] * 60 + parts[2]
    units = re.findall(r"([\d.,]+)\s*(h|jam|m|min|menit|s|sec|detik)", s)
    if units:
        mult = {"h": 3600, "jam": 3600, "m": 60, "min": 60, "menit": 60, "s": 1, "sec": 1, "detik": 1}
        return sum(float(n.replace(",", ".")) * mult[u] for n, u in units)
    n = parse_number(s)
    return n


# --------------------------------------------------------------------------
# Membaca baris export -> record
# --------------------------------------------------------------------------

def _norm(s):
    s = re.sub(r"\((rp|idr|sec|s|%)\)", "", str(s).lower())
    return re.sub(r"[^a-z0-9]", "", s)


def resolve_columns(columns, spec, cfg):
    """Cocokkan nama kolom export dengan field. Return dict field -> nama kolom export."""
    fields = {k: list(v) for k, v in spec["fields"].items()}
    if spec["key"] == "sv" and cfg.get("gmv_column"):
        fields["GMV"] = [cfg["gmv_column"]]
    for k, v in (cfg.get(spec["map_cfg"]) or {}).items():
        fields[k] = [v] if isinstance(v, str) else list(v)
    exact = {str(c).strip().lower(): c for c in columns}
    normed = {}
    for c in columns:
        normed.setdefault(_norm(c), c)
    out = {}
    for field, cands in fields.items():
        for cand in cands:
            c = exact.get(cand.strip().lower()) or normed.get(_norm(cand))
            if c is not None:
                out[field] = c
                break
    return out


def collect_export_rows(export_frames, spec, cfg, start=None, end=None, warn=None):
    """Return dict id -> (start datetime, record, filename). File terbaru menang.
    start/end (date, inklusif) membatasi berdasarkan waktu; None = semua."""
    latest = {}
    for _, fname, df in sorted(export_frames, key=lambda x: x[0]):
        colmap = resolve_columns(df.columns, spec, cfg)
        if "_id" not in colmap or "_time" not in colmap:
            continue  # bukan file export untuk tracker ini
        if warn is not None:
            optional = {"_end", "_duration", "_title", "_pname"}
            warn.update(f for f in spec["fields"] if f not in colmap and f not in optional)
        order = frame_date_order(df, colmap, cfg, spec["date_order"])
        for _, r in df.iterrows():
            rec = {f: r.get(c) for f, c in colmap.items()}
            vid = clean_id(rec.get("_id"))
            t = parse_time(rec.get("_time"), order)
            if not vid or t is None:
                continue
            if (start and t.date() < start) or (end and t.date() > end):
                continue
            rec["_order"] = order
            latest[vid] = (t, rec, fname)
    return latest


def frame_date_order(df, colmap, cfg, default="mdy"):
    """Urutan tanggal untuk satu file: dari config ("date_order": "mdy"/"dmy"), atau ditebak otomatis."""
    forced = str(cfg.get("date_order", "auto")).lower()
    if forced in ("mdy", "dmy"):
        return forced
    vals = []
    for f in ("_time", "_end"):
        if f in colmap:
            vals += list(df[colmap[f]])
    return detect_date_order(vals, default=default)


def available_months(export_frames, spec, cfg):
    out = {}
    for _, _, df in export_frames:
        colmap = resolve_columns(df.columns, spec, cfg)
        col = colmap.get("_time")
        if col is None:
            continue
        order = frame_date_order(df, colmap, cfg, spec["date_order"])
        for v in df[col]:
            t = parse_time(v, order)
            if t:
                out[(t.year, t.month)] = out.get((t.year, t.month), 0) + 1
    return out


_DIGITS_RE = re.compile(r"\d{10,}")


def product_of(rec, spec, bank):
    """Return (PID, Product Name) untuk satu record."""
    if spec["key"] == "sv":
        return parse_first_product(rec.get("_products"))
    raw = rec.get("_pid_raw")
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return "", ""
    if isinstance(raw, (int, float)):
        pid = clean_id(raw)
    else:
        pid, _ = parse_first_product(raw)
        if not pid:
            m = _DIGITS_RE.search(str(raw))
            pid = m.group(0) if m else ""
    name = bank.name(pid)
    if not name:
        pn = rec.get("_pname")
        if pn is not None and not (isinstance(pn, float) and pd.isna(pn)):
            name = str(pn).strip()
    return pid, name


def metric_values(rec, spec):
    out = {}
    for col in metric_cols(spec):
        raw = rec.get(col)
        if col.endswith("%"):
            out[col] = parse_percent(raw)
        elif col.startswith("Avg Viewing Duration"):
            out[col] = parse_seconds(raw)
        else:
            out[col] = parse_number(raw)
    return out


def live_times(t, rec):
    """Return (end datetime atau None, durasi timedelta atau None)."""
    end = parse_time(rec.get("_end"), rec.get("_order", "dmy"))
    if end is not None and end >= t:
        return end, end - t
    secs = parse_seconds(rec.get("_duration"))
    if secs not in ("", None):
        d = timedelta(seconds=float(secs))
        return t + d, d
    return None, None


def build_tracker_row(spec, t, rec, bank, tiers, cfg, report, product=None):
    """product=(PID, Name): paksa satu produk saja (dipakai tab Exploded SV).
    Tanpa product: SV = semua produk digabung (CAT/PRO dari produk pertama)."""
    sep = cfg.get("multi_separator", " | ")
    prods = [product] if product is not None else products_of(rec, spec, bank)
    pid, pname = join_products(prods, sep)
    creator = "" if rec.get("_creator") is None or pd.isna(rec.get("_creator")) else str(rec.get("_creator")).strip()
    cat, pro = bank.get(prods[0][0])
    for p, n in prods:
        if p and not bank.get(p)[0]:
            report["missing_pid"].setdefault(p, {"name": n, "count": 0})["count"] += 1
    if not pid:
        report["no_pid"].append(clean_id(rec.get("_id")))
    tier = tiers.get(clean_name(creator))
    if tier is None:
        report["missing_creator"][creator] = report["missing_creator"].get(creator, 0) + 1
        tier = cfg.get("default_tier", "OPEN PLAN")

    v = {
        "Date": t.strftime("%Y-%m-%d"),
        "Week": week_label(t),
        spec["id_col"]: as_text(clean_id(rec.get("_id"))),
        "Creator": safe_text(creator),
        "Hastag": extract_hashtags(rec.get("_video_info")),
        "Tiering": tier,
        "PID": as_text(pid),
        "Product Name": safe_text(pname),
        "CAT": cat,
        "PRO": pro,
    }
    if spec["key"] == "sv":
        v["Upload Datetime"] = t.strftime("%Y-%m-%d %H:%M:%S")
    else:
        title = rec.get("_title")
        v["Live Title"] = "" if title is None or (isinstance(title, float) and pd.isna(title)) else safe_text(str(title).strip())
        end, dur = live_times(t, rec)
        v["Start Time"] = t.strftime("%Y-%m-%d %H:%M:%S")
        v["End Time"] = end.strftime("%Y-%m-%d %H:%M:%S") if end else ""
        v["Duration"] = round(dur.total_seconds() / 86400, 10) if dur is not None else ""
    v.update(metric_values(rec, spec))
    return [v.get(c, "") for c in spec["columns"]]


def find_missing_pids(latest, ids, spec, bank):
    missing = {}
    for i in ids:
        for pid, pname in products_of(latest[i][1], spec, bank):
            if pid and pid not in bank.lookup:
                missing.setdefault(pid, {"name": pname, "count": 0})["count"] += 1
    return missing


def new_report():
    return {"missing_pid": {}, "missing_creator": {}, "no_pid": []}


def _same(a, b):
    if a in ("", None) and b in ("", None):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-9 * max(1, abs(a), abs(b))
    return str(a) == str(b)


def _blank(v):
    return v is None or str(v).strip() == ""


def plan_existing_updates(spec, tracker_rows, header_row, latest_all, bank, cfg, refresh=False, exploded=False):
    """Untuk ID yang sudah ada di tracker:
    - metrik diperbarui kalau angkanya berubah di file terbaru;
    - PID / Product Name / CAT / PRO yang masih kosong dilengkapi;
    - refresh=True (hanya tab SV): PID / Product Name baris lama ditimpa dengan SEMUA produk dari export;
    - exploded=True: tab Exploded SV, satu PID per baris, PID tidak pernah diturunkan dari export."""
    sep = cfg.get("multi_separator", " | ")
    cols = spec["columns"]
    i_id = cols.index(spec["id_col"])
    i_m = cols.index(spec["first_metric"])
    mcols = metric_cols(spec)
    i_pid = cols.index("PID")
    i_h = cols.index("Hastag") if "Hastag" in cols else None
    pblock = ["PID", "Product Name", "CAT", "PRO"]
    L = col_letter
    updates, n_metric, n_prod = [], 0, 0
    for n, row in enumerate(tracker_rows, start=header_row + 1):
        row = list(row) + [""] * (len(cols) - len(row))
        vid = clean_id(row[i_id])
        rec = latest_all.get(vid, (None, None, None))[1] if vid else None
        if rec is not None:
            new = metric_values(rec, spec)
            new_vals = [new[c] for c in mcols]
            if not all(_same(o, nv) for o, nv in zip(row[i_m:], new_vals)):
                updates.append({"range": f"{L(i_m + 1)}{n}:{L(len(cols))}{n}", "values": [new_vals]})
                n_metric += 1
        if i_h is not None and rec is not None and _blank(row[i_h]):
            tags = extract_hashtags(rec.get("_video_info"))
            if tags:
                updates.append({"range": f"{L(i_h + 1)}{n}", "values": [[tags]]})
                n_prod += 1
        old = row[i_pid:i_pid + 4]
        if refresh and not exploded and rec is not None and spec["key"] == "sv":
            prods = products_of(rec, spec, bank)
            new_pid, new_name = join_products(prods, sep)
            if new_pid and (clean_id(old[0]) != new_pid or str(old[1]) != new_name):
                cat, pro = bank.get(prods[0][0])
                updates.append({"range": f"{L(i_pid + 1)}{n}:{L(i_pid + len(pblock))}{n}",
                                "values": [[as_text(new_pid), safe_text(new_name),
                                            old[2] if not _blank(old[2]) else cat,
                                            old[3] if not _blank(old[3]) else pro]]})
                n_prod += 1
                continue
        if not any(_blank(x) for x in old):
            continue
        pid, pname = clean_id(old[0]), old[1]
        if not pid and rec is not None and not exploded:
            pid, pname2 = join_products(products_of(rec, spec, bank), sep)
            pname = pname if not _blank(pname) else pname2
        if not pid:
            continue
        first_pid = pid.split(sep.strip() or sep)[0].strip()  # CAT/PRO/nama dicari dari PID pertama
        if _blank(pname):
            pname = bank.name(first_pid) if first_pid == pid else ""
        cat, pro = bank.get(first_pid)
        merged = [as_text(pid),
                  pname if not _blank(old[1]) else safe_text(pname or ""),
                  old[2] if not _blank(old[2]) else cat,
                  old[3] if not _blank(old[3]) else pro]
        if [str(x).lstrip("'") for x in merged] != [str(x) for x in old]:
            if not _blank(old[0]):
                merged[0] = as_text(clean_id(old[0]))
            updates.append({"range": f"{L(i_pid + 1)}{n}:{L(i_pid + len(pblock))}{n}", "values": [merged]})
            n_prod += 1
    return updates, n_metric, n_prod


# --------------------------------------------------------------------------
# Penulisan ke sheet
# --------------------------------------------------------------------------

def read_tracker(ws, spec, header_row):
    last = col_letter(len(spec["columns"]))
    return ws.get(f"A{header_row + 1}:{last}", value_render_option="UNFORMATTED_VALUE")


def append_rows(ws, spec, rows, header_row):
    last = col_letter(len(spec["columns"]))
    resp = ws.append_rows(rows, value_input_option="USER_ENTERED",
                          insert_data_option="INSERT_ROWS",
                          table_range=f"A{header_row}:{last}{header_row}")
    updated = resp.get("updates", {}).get("updatedRange", "")
    m = re.search(r"!A(\d+):", updated)
    if not m:
        return None, None
    start = int(m.group(1))
    end = start + len(rows) - 1
    fmts = []
    for i, c in enumerate(spec["columns"], 1):
        if c in spec["formats"]:
            typ, pattern = spec["formats"][c]
            fmts.append({"range": f"{col_letter(i)}{start}:{col_letter(i)}{end}",
                         "format": {"numberFormat": {"type": typ, "pattern": pattern}}})
    if fmts:
        ws.batch_format(fmts)
    return start, end


def write_updates(ws, updates):
    for i in range(0, len(updates), 500):
        ws.batch_update(updates[i:i + 500], value_input_option="USER_ENTERED")


# --------------------------------------------------------------------------
# Laporan
# --------------------------------------------------------------------------

def write_csv(path, header, rows):
    """Tulis CSV. Kalau file lama sedang dibuka di Excel, simpan dengan nama lain."""
    try:
        f = open(path, "w", newline="", encoding="utf-8-sig")
    except PermissionError:
        alt = path[:-4] + datetime.now().strftime("_%H%M%S") + ".csv"
        print(f"   ⚠️  {os.path.basename(path)} sedang dibuka (Excel?). Disimpan sebagai {os.path.basename(alt)}")
        path, f = alt, open(alt, "w", newline="", encoding="utf-8-sig")
    with f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    return path


def print_and_save_report(report, out_dir, label, unit):
    os.makedirs(out_dir, exist_ok=True)
    mp, mc = report["missing_pid"], report["missing_creator"]
    if mp:
        print(f"⚠️  {len(mp)} PID masih belum ada CAT/PRO:")
        for pid, info in sorted(mp.items(), key=lambda x: -x[1]["count"])[:20]:
            print(f"   - {pid} ({info['count']} {unit}): {(info['name'] or '')[:90]}")
        path = write_csv(os.path.join(out_dir, f"pid_belum_ada_{label}.csv"), ["PID", "Product Name", "Jumlah"],
                         [[pid, i["name"], i["count"]] for pid, i in mp.items()])
        print(f"   (daftar lengkap: {path})")
    else:
        print("✅ Semua PID sudah ada CAT/PRO-nya.")
    if mc:
        print(f"⚠️  {len(mc)} creator tidak ada di Creator Tiering (diisi default tier):")
        for name, n in sorted(mc.items(), key=lambda x: -x[1])[:20]:
            print(f"   - {name} ({n} {unit})")
        if len(mc) > 20:
            print(f"   ... dan {len(mc) - 20} lainnya")
        path = write_csv(os.path.join(out_dir, f"creator_belum_ada_tier_{label}.csv"), ["Creator", "Jumlah"],
                         list(mc.items()))
        print(f"   (daftar lengkap: {path})")
    if report["no_pid"]:
        print(f"⚠️  {len(report['no_pid'])} {unit} tidak punya PID.")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def parse_date_arg(s):
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    sys.exit(f"Format tanggal tidak dikenali: {s} (pakai YYYY-MM-DD, contoh 2026-10-01)")


def month_range(y, m):
    return date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1])


def choose_range(available, args):
    """Return (start, end) inklusif."""
    if args.start or args.end:
        start = parse_date_arg(args.start) if args.start else date(1900, 1, 1)
        end = parse_date_arg(args.end) if args.end else date(2999, 12, 31)
        return start, end
    if args.month:
        if args.month == "current":
            t = date.today()
            return month_range(t.year, t.month)
        y, m = args.month.split("-")
        return month_range(int(y), int(m))
    if not available:
        sys.exit("Tidak ada data dengan tanggal yang valid di file export.")
    print("Bulan yang ada di file export:")
    options = sorted(available)
    for i, (y, m) in enumerate(options, 1):
        print(f"  {i}. {calendar.month_name[m]} {y}  ({available[(y, m)]} baris)")
    print("  R. Rentang tanggal bebas")
    pick = input("Pilih nomor bulan, atau R: ").strip()
    if pick.lower() == "r":
        start = parse_date_arg(input("  Dari tanggal (YYYY-MM-DD): "))
        end = parse_date_arg(input("  Sampai tanggal (YYYY-MM-DD): "))
        return start, end
    return month_range(*options[int(pick) - 1])


def load_exports_from_local(folder):
    frames = []
    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if not name.lower().endswith((".xlsx", ".xls", ".csv")) or name.startswith("~$"):
            continue
        with open(path, "rb") as f:
            content = f.read()
        df = _read_export(content, name)
        if df is not None:
            frames.append((datetime.fromtimestamp(os.path.getmtime(path)).isoformat(), name, df))
    return frames


def _read_export(content, filename):
    for key in ("Video ID", "Live ID", "LIVE ID", "Product ID", "Start time", "Creator name"):
        try:
            return read_table_bytes(content, filename, key)
        except ValueError:
            continue
    print(f"   lewati {filename}: header tidak dikenali")
    return None


def load_exports_from_drive(g, folder_id):
    frames = []
    for meta in g.list_folder(folder_id):
        name = meta["name"]
        if meta["mimeType"] != GOOGLE_SHEET_MIME and not name.lower().endswith((".xlsx", ".xls", ".csv")):
            continue
        content, filename = g.download(meta)
        df = _read_export(content, filename)
        if df is not None:
            frames.append((meta["modifiedTime"], name, df))
            print(f"   ✓ {name} ({len(df)} baris)")
    return frames


def run_mode(spec, g, cfg, args, frames, start, end, bank, tiers, out_dir, label):
    print(f"\n========== {spec['label']} ==========")
    header_row = int(cfg.get(spec["header_cfg"][0], spec["header_cfg"][1]))
    book = g.gc.open_by_key(cfg["tracker_spreadsheet_id"])
    ws = book.worksheet(cfg.get(*spec["sheet_cfg"]))
    tracker_rows = read_tracker(ws, spec, header_row)
    i_id = spec["columns"].index(spec["id_col"])
    i_pid = spec["columns"].index("PID")
    i_cat = spec["columns"].index("CAT")
    existing = {clean_id(r[i_id]) for r in tracker_rows if len(r) > i_id and clean_id(r[i_id])}
    print(f"   {len(existing)} {spec['id_col']} sudah ada di tab {ws.title}")

    # Tab Exploded SV: tabel yang sama, satu baris per (Video ID, PID)
    exp_ws, exp_rows, exp_header = None, [], header_row
    if spec["key"] == "sv" and not args.no_exploded:
        exp_name = cfg.get("sv_exploded_sheet", "Exploded SV")
        exp_header = int(cfg.get("sv_exploded_header_row", header_row))
        try:
            exp_ws = book.worksheet(exp_name)
            exp_rows = read_tracker(exp_ws, spec, exp_header)
        except Exception as e:
            exp_ws = None
            print(f"   ⚠️  Tab '{exp_name}' tidak bisa dibuka ({e}). Tab Exploded dilewati.")
    exp_keys = {(clean_id(r[i_id]), clean_id(r[i_pid]) if len(r) > i_pid else "")
                for r in exp_rows if len(r) > i_id and clean_id(r[i_id])}
    if exp_ws is not None:
        print(f"   {len(exp_keys)} baris (Video ID + PID) sudah ada di tab {exp_ws.title}")

    warn = set()
    latest = collect_export_rows(frames, spec, cfg, start, end, warn)
    if warn:
        print(f"   ⚠️  Kolom ini tidak ketemu di file export (akan kosong): {', '.join(sorted(warn))}")
        print(f"      Kalau namanya beda, atur di config.json bagian \"{spec['map_cfg']}\".")
    new_ids = sorted((v for v in latest if v not in existing), key=lambda v: latest[v][0])
    print(f"   {len(latest)} {spec['unit']} di rentang ini, {len(latest) - len(new_ids)} sudah ada, {len(new_ids)} baru")

    latest_all = collect_export_rows(frames, spec, cfg) if not args.no_update_existing else {}
    backfill_ids = [clean_id(r[i_id]) for r in tracker_rows
                    if len(r) > i_id and clean_id(r[i_id]) in latest_all and (len(r) <= i_cat or _blank(r[i_cat]))]
    exp_backfill_ids = [clean_id(r[i_id]) for r in exp_rows
                        if len(r) > i_id and clean_id(r[i_id]) in latest_all and (len(r) <= i_cat or _blank(r[i_cat]))]

    # Baris baru untuk Exploded SV: setiap (Video ID, PID) di rentang yang belum ada di tab itu.
    # Dicek per pasangan (bukan per Video ID), jadi tab yang masih kosong juga terisi dari rentang ini.
    exp_new = []
    if exp_ws is not None:
        for vid in sorted(latest, key=lambda v: latest[v][0]):
            t, rec, _ = latest[vid]
            for prod in products_of(rec, spec, bank):
                if (vid, prod[0]) not in exp_keys:
                    exp_new.append((t, rec, prod))
        print(f"   Exploded: {len(exp_new)} baris baru (Video ID + PID)")

    # PID baru -> tanya CAT/PRO, nanti ditulis ke Product Bank (semua produk, bukan hanya yang pertama)
    pb_entries = []
    pool = dict(latest_all)
    pool.update(latest)
    check_ids = new_ids + backfill_ids + exp_backfill_ids + (list(latest) if exp_ws is not None else [])
    missing = find_missing_pids(pool, list(dict.fromkeys(check_ids)), spec, bank)
    if missing and not args.dry_run and not args.no_product_bank:
        if args.yes:
            pb_entries = [{"pid": p, "name": i["name"], "cat": "", "pro": ""} for p, i in missing.items()]
        else:
            pb_entries = ask_cat_pro(missing, bank)
        for e in pb_entries:
            if e["cat"]:
                bank.lookup[e["pid"]] = (e["cat"], e["pro"])
            if e["name"] and not bank.name(e["pid"]):
                bank.names[e["pid"]] = e["name"]

    report = new_report()
    rows = [build_tracker_row(spec, latest[v][0], latest[v][1], bank, tiers, cfg, report) for v in new_ids]
    exp_built = [build_tracker_row(spec, t, rec, bank, tiers, cfg, new_report(), product=prod)
                 for t, rec, prod in exp_new]

    updates, n_metric, n_prod = [], 0, 0
    exp_updates, en_metric, en_prod = [], 0, 0
    if not args.no_update_existing:
        updates, n_metric, n_prod = plan_existing_updates(spec, tracker_rows, header_row, latest_all, bank, cfg,
                                                          refresh=args.refresh_products)
        print(f"   {n_metric} {spec['unit']} lama metriknya berubah, {n_prod} baris lama data produk/hastag-nya bisa dilengkapi")
        if exp_ws is not None:
            exp_updates, en_metric, en_prod = plan_existing_updates(spec, exp_rows, exp_header, latest_all, bank,
                                                                    cfg, exploded=True)
            print(f"   Exploded: {en_metric} baris lama metriknya berubah, {en_prod} baris lama data produk/hastag-nya bisa dilengkapi")

    preview = write_csv(os.path.join(out_dir, f"preview_{spec['key']}_{label}.csv"), spec["columns"],
                        [[str(x).lstrip("'") for x in r] for r in rows])
    if exp_ws is not None:
        write_csv(os.path.join(out_dir, f"preview_{spec['key']}_exploded_{label}.csv"), spec["columns"],
                  [[str(x).lstrip("'") for x in r] for r in exp_built])
    print_and_save_report(report, out_dir, f"{spec['key']}_{label}", spec["unit"])

    if args.dry_run:
        print(f"(dry-run) Tidak ada yang ditulis. Preview baris baru: {preview}")
        return
    if not rows and not updates and not pb_entries and not exp_built and not exp_updates:
        print("Tidak ada perubahan.")
        return
    if not args.yes:
        parts = []
        if pb_entries:
            parts.append(f"tambah/isi {len(pb_entries)} PID di Product Bank")
        if rows:
            parts.append(f"tambah {len(rows)} baris baru")
        if updates:
            parts.append(f"perbarui {len(updates)} bagian baris lama")
        msg = f"\nLanjut ke tab {ws.title}: {', '.join(parts) or 'tidak ada perubahan'}"
        if exp_ws is not None and (exp_built or exp_updates):
            msg += f"; tab {exp_ws.title}: tambah {len(exp_built)} baris, perbarui {len(exp_updates)} bagian baris lama"
        if input(msg + "? (y/n): ").strip().lower() != "y":
            print("Dibatalkan.")
            return

    if pb_entries:
        pb_id, pb_sheet = cfg["product_bank_file_id"], cfg.get("product_bank_sheet", "Product Bank")
        try:
            write_product_bank(g, pb_id, pb_sheet, pb_entries)
            print(f"✅ {len(pb_entries)} PID ditulis ke Product Bank.")
        except Exception as e:  # biasanya tidak punya akses edit
            path = write_csv(os.path.join(out_dir, f"produk_baru_untuk_product_bank_{label}.csv"),
                             ["Product ID", "Product Name", "CAT", "PRO"],
                             [[x["pid"], x["name"], x["cat"], x["pro"]] for x in pb_entries])
            print(f"⚠️  Gagal menulis ke Product Bank ({e}).\n   Data disimpan di {path}, silakan paste manual.")
    if updates:
        write_updates(ws, updates)
        print(f"✅ {n_metric} {spec['unit']} lama diperbarui metriknya, {n_prod} baris dilengkapi data produk/hastag.")
    if rows:
        s_row, e_row = append_rows(ws, spec, rows, header_row)
        print(f"✅ {len(rows)} baris baru ditambahkan ke {ws.title}" + (f" (baris {s_row}-{e_row})." if s_row else "."))
    if exp_ws is not None:
        if exp_updates:
            write_updates(exp_ws, exp_updates)
            print(f"✅ {exp_ws.title}: {en_metric} baris lama diperbarui metriknya, {en_prod} baris dilengkapi data produk/hastag.")
        if exp_built:
            s_row, e_row = append_rows(exp_ws, spec, exp_built, exp_header)
            print(f"✅ {len(exp_built)} baris baru ditambahkan ke {exp_ws.title}" + (f" (baris {s_row}-{e_row})." if s_row else "."))


def main():
    ap = argparse.ArgumentParser(description="Isi Daily Tracker NIVEA (SV & LS) dari file export harian")
    ap.add_argument("--config", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json"))
    ap.add_argument("--mode", choices=["sv", "ls", "all"], help="sv = video, ls = live, all = keduanya")
    ap.add_argument("--month", help="YYYY-MM, atau 'current'")
    ap.add_argument("--start", help="Tanggal mulai, YYYY-MM-DD (inklusif)")
    ap.add_argument("--end", help="Tanggal akhir, YYYY-MM-DD (inklusif)")
    ap.add_argument("--dry-run", action="store_true", help="Tidak menulis apa pun, hanya preview + laporan")
    ap.add_argument("--yes", action="store_true",
                    help="Tanpa tanya apa pun (untuk jadwal otomatis). PID baru masuk Product Bank dengan CAT/PRO kosong")
    ap.add_argument("--local-dir", help="Baca file export dari folder lokal, bukan dari Google Drive")
    ap.add_argument("--no-update-existing", action="store_true",
                    help="Jangan perbarui metrik ID yang sudah ada di tracker")
    ap.add_argument("--no-product-bank", action="store_true",
                    help="Jangan menambahkan PID baru ke Product Bank")
    ap.add_argument("--no-exploded", action="store_true",
                    help="Jangan isi tab Exploded SV (satu baris per PID)")
    ap.add_argument("--refresh-products", action="store_true",
                    help="Tab SV: timpa PID & Product Name baris lama dengan SEMUA produk dari export "
                         "(default: baris lama dibiarkan, hanya baris baru yang memuat semua produk)")
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as f:
        cfg = json.load(f)
    base = os.path.dirname(os.path.abspath(args.config))
    cred_file = cfg.get("credentials_file") or cfg.get("service_account_file") or "credentials.json"
    if not os.path.isabs(cred_file):
        cred_file = os.path.join(base, cred_file)
    if not os.path.exists(cred_file):  # kompatibel dengan nama lama
        for alt in ("credentials.json", "service_account.json"):
            if os.path.exists(os.path.join(base, alt)):
                cred_file = os.path.join(base, alt)
                break
    token_file = os.path.join(base, cfg.get("token_file", "token.json"))
    out_dir = os.path.join(base, cfg.get("report_dir", "laporan"))
    os.makedirs(out_dir, exist_ok=True)

    mode = args.mode
    if not mode:
        if args.yes:
            mode = "all"
        else:
            pick = input("Mau isi tab apa?  1. SV (Video)   2. LS (Live)   3. Keduanya  : ").strip()
            mode = {"1": "sv", "2": "ls", "3": "all"}.get(pick, "sv")
    specs = [SV, LS] if mode == "all" else [SPECS[mode]]

    g = Google(cred_file, token_file)

    frames_by_mode = {}
    for spec in specs:
        print(f"1) Membaca file export {spec['label']}...")
        if args.local_dir:
            frames_by_mode[spec["key"]] = load_exports_from_local(args.local_dir)
        else:
            folder = cfg.get(spec["folder_cfg"])
            if not folder:
                sys.exit(f"'{spec['folder_cfg']}' belum diisi di config.json")
            frames_by_mode[spec["key"]] = load_exports_from_drive(g, folder)
        if not frames_by_mode[spec["key"]]:
            print(f"   Tidak ada file export {spec['label']} yang bisa dibaca.")

    available = {}
    for spec in specs:
        for k, n in available_months(frames_by_mode[spec["key"]], spec, cfg).items():
            available[k] = available.get(k, 0) + n
    start, end = choose_range(available, args)
    label = f"{start:%Y%m%d}-{end:%Y%m%d}" if (start.year > 1900 and end.year < 2999) else "custom"
    print(f"   Rentang dipilih: {start:%d %b %Y} s/d {end:%d %b %Y}")

    print("2) Membaca Product Bank & Creator Tiering...")
    bank = ProductBank(g.read_table(cfg["product_bank_file_id"], "Product ID",
                                    cfg.get("product_bank_sheet", "Product Bank")))
    tiers = build_tier_lookup(g.read_table(cfg["creator_tiering_file_id"], "Nama Creator",
                                           cfg.get("creator_tiering_sheet")))
    print(f"   {len(bank.pids)} PID ({len(bank.lookup)} sudah ada CAT/PRO), {len(tiers)} creator")

    for spec in specs:
        if frames_by_mode[spec["key"]]:
            run_mode(spec, g, cfg, args, frames_by_mode[spec["key"]], start, end, bank, tiers, out_dir, label)


if __name__ == "__main__":
    main()
