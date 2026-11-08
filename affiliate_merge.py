#!/usr/bin/env python3
"""
Gabungkan export Affiliate_Order dan Affiliate_Patner (TikTok Shop) ke satu tab Google Sheet.

Alur:
  1. Tampilkan daftar file di folder Google Drive Affiliate_Order dan Affiliate_Patner,
     kamu pilih file mana yang jadi input (boleh pilih lebih dari satu, boleh lewati salah satu).
  2. Kolom disesuaikan dengan header di tab output ("[Order] Creator+Patner").
     Kolom yang tidak ada di salah satu sumber dikosongkan.
  3. Kolom Source diisi AFFILIATE (dari Affiliate_Order) atau PATNER (dari Affiliate_Patner).
  4. Kolom Type Product diisi dari "SKU name" (kalau SKU name kosong / "Default", pakai "Product Name"):
       Single / Multipack / Bundling, ditambah " + GWP" kalau ada gift with purchase.
  5. Baris baru ditambahkan. Baris yang sudah ada (kunci: Source + Order ID + Sku Id) tidak dobel;
     kolom yang bisa berubah (status order, komisi aktual, dll) diperbarui.
  6. Laporan Type Product (SKU/produk apa jadi tipe apa) disimpan sebagai CSV supaya mudah dicek.
  7. CAT & PRO diambil dari Product Bank (dicocokkan lewat Product ID).
  8. Sebelum menulis, ditampilkan isi tiap file (jumlah baris, rentang Time Created, total Est. Commission Base)
     per Source, berapa baris dobel yang dilewati, dan posisi sheet sebelum / sesudah.
  9. Mode product: gabungkan file harian product_list_YYYYMMDD.xlsx dari folder Product Performance ke tab
     "[Product Performance] Daily November (Raw)". Tanggal diambil dari nama file; CAT/PRO dari Product Bank;
     tanggal yang sudah ada di sheet tidak dimasukkan lagi (kecuali --refresh).
 10. Setelah menulis, tab "Last Update" (data sampai tanggal berapa) dan "Update Log" (riwayat tiap run)
     di spreadsheet output diperbarui.

Contoh:
  python affiliate_merge.py                      # pilih file lewat daftar
  python affiliate_merge.py --dry-run            # cek dulu, tidak menulis apa pun
  python affiliate_merge.py --order "1-8 November" --partner "1-8 November" --yes
  python affiliate_merge.py --order-path order.csv --partner-path partner.csv --dry-run   # file lokal

config.json (sama dengan nivea_tracker.py) boleh berisi:
  "credentials_file", "token_file",
  "affiliate_order_folder_id", "affiliate_partner_folder_id",
  "merge_output_spreadsheet_id", "merge_output_sheet", "merge_header_row",
  "source_label_order" (AFFILIATE), "source_label_partner" (PATNER),
  "type_labels": {"single": "Single", "multipack": "Multipack", "bundling": "Bundling"},
  "gwp_suffix": " + GWP", "gwp_keywords": [...], "bundling_keywords": [...],
  "type_overrides": {"teks SKU name / Product Name persis": "Bundling"},
  "merge_product_bank_file_id", "merge_product_bank_sheet" ("Product Bank"),
  "merge_last_update_sheet" ("Last Update"), "merge_log_sheet" ("Update Log"),
  "product_folder_id", "product_output_sheet" ("[Product Performance] Daily November (Raw)"), "product_header_row" (3, otomatis terdeteksi)
"""

import argparse
import csv
import io
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta

import pandas as pd

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",  # sama dengan nivea_tracker.py supaya token.json bisa dipakai bersama
]
GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

DEFAULTS = {
    "affiliate_order_folder_id": "1iUzLAYExrm4dNJYH5Y4iU18O2ZAQHKFc",
    "affiliate_partner_folder_id": "1vhvcJX8yaSXUoADQeOhb6QI2_xYWVoau",
    "merge_output_spreadsheet_id": "1hrwzM1pnVZ5l-yYQ3cKnS480XJJUKW5X1bBB9I-T3Hg",
    "merge_output_sheet": "[Order] Creator+Patner",
    "merge_header_row": 1,
    "merge_product_bank_file_id": "1N7CKyWKEg1DKKgJHNWMvK_PXfHpAGw-9",
    "merge_product_bank_sheet": "Product Bank",
    "merge_last_update_sheet": "Last Update",
    "merge_log_sheet": "Update Log",
    "product_folder_id": "1uJc9bov06GNPXk9-jXMF-Grl93RD14C1",
    "product_output_sheet": "[Product Performance] Daily November (Raw)",
    "product_header_row": 3,
    "product_dataset_label": "PRODUCT PERFORMANCE",
    "source_label_order": "AFFILIATE",
    "source_label_partner": "PATNER",
    "date_order": "dmy",
    "type_labels": {"single": "Single", "multipack": "Multipack", "bundling": "Bundling"},
    "gwp_suffix": " + GWP",
    # kata penanda gift with purchase (dicari di SKU name / Product Name, huruf kecil)
    "gwp_keywords": ["gwp", "gift", "hadiah", "gratis", "bonus", "free gift"],
    # kata penanda bundling (gabungan beberapa produk berbeda)
    "bundling_keywords": ["package", "paket kombo", "paket hemat kombo", "kombo", "bundle", "bundling", "mix",
                          "regime", "kit", "set", "day & night", "face & body", "lotion & serum", "body & deo",
                          "buy 1 get 1", "buy 2 get 1", "b1g1", "face & deo", "face & creme", "face & crème", "sun & body",
                          "roll on & spray"],
    # kata penanda multipack (produk yang sama lebih dari 1 pcs)
    "multipack_keywords": ["twinpack", "twin pack", "triplepack", "triple pack", "doublepack", "double pack",
                           "multipack", "multi pack"],
    "type_overrides": {},
}

# Urutan kolom sumber -> kolom output. Kunci = nama kolom output, nilai = kemungkinan nama kolom di export.
FIELD_ALIASES = {
    "Time Created": ["Time Created"],
    "Order ID": ["Order ID"],
    "Product ID": ["Product ID"],
    "Product Name": ["Product Name"],
    "Sku Id": ["Sku Id", "SKU ID"],
    "SKU name": ["SKU name", "SKU Name"],
    "Seller SKU": ["Seller SKU"],
    "Price": ["Price"],
    "Payment Amount": ["Payment Amount"],
    "Quantity": ["Quantity"],
    "Payment method": ["Payment method"],
    "Order Status": ["Order Status"],
    "Creator Username": ["Creator Username"],
    "Content Type": ["Content Type"],
    "Carousel": ["Carousel"],
    "Content ID": ["Content ID"],
    "commission model": ["commission model", "Commission model"],
    "Standard commission rate": ["Standard commission rate", "Commission Rate (%)", "Commission Rate"],
    "Est. Commission Base": ["Est. Commission Base"],
    "Est. standard commission payment": ["Est. standard commission payment", "Est. Commission Payment"],
    "Actual Commission Base": ["Actual Commission Base"],
    "Actual Commission Payment": ["Actual Commission Payment"],
    "Partner campaign ID": ["Partner campaign ID"],
    "Affiliate Partner ID": ["Affiliate Partner ID"],
    "Affiliate Partner Name": ["Affiliate Partner Name"],
    "Affiliate Partner Commission Rate": ["Affiliate Partner Commission Rate"],
    "Est. Commission for Affiliate Partner": ["Est. Commission for Affiliate Partner"],
    "Actual Commission for Affiliate Partner": ["Actual Commission for Affiliate Partner"],
}

ID_COLS = {"Order ID", "Product ID", "Sku Id", "Content ID", "Partner campaign ID", "Affiliate Partner ID"}
TEXT_COLS = {"Product Name", "SKU name", "Seller SKU", "Payment method", "Order Status", "Creator Username",
             "Content Type", "Carousel", "commission model", "Affiliate Partner Name", "Source", "Type Product",
             "CAT", "PRO"}
NUM_COLS = {"Price", "Payment Amount", "Quantity", "Est. Commission Base", "Est. standard commission payment",
            "Actual Commission Base", "Actual Commission Payment", "Est. Commission for Affiliate Partner",
            "Actual Commission for Affiliate Partner"}
# Kolom yang bisa berubah di export berikutnya (status order, komisi aktual, dll)
MUTABLE_COLS = ["Order Status", "Payment Amount", "Est. Commission Base", "Est. standard commission payment",
                "Actual Commission Base", "Actual Commission Payment", "Est. Commission for Affiliate Partner",
                "Actual Commission for Affiliate Partner", "Standard commission rate",
                "Affiliate Partner Commission Rate", "CAT", "PRO"]
KEY_COLS = ["Source", "Order ID", "Sku Id"]


# --------------------------------------------------------------------------
# Helper umum
# --------------------------------------------------------------------------

def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def is_blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def clean_id(value):
    if is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip().lstrip("'")
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s


def as_text(s):
    s = "" if s is None else str(s)
    return "'" + s if s else ""


def safe_text(s):
    s = "" if s is None else str(s)
    return "'" + s if s[:1] in ("=", "+", "-", "@") else s


def to_number(v):
    """Angka murni jadi int/float; selain itu (mis. '20%') dibiarkan sebagai teks."""
    if is_blank(v):
        return ""
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else v
    s = str(v).strip()
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d+\.\d+", s):
        f = float(s)
        return int(f) if f.is_integer() else f
    return s


def parse_time(value, order="dmy"):
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value
    s = str(value).strip()
    local = ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y") if order == "dmy" else \
            ("%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%m/%d/%Y")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d") + local:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def _norm(s):
    s = re.sub(r"\((rp|idr|sec|s|%)\)", "", str(s).lower())
    return re.sub(r"[^a-z0-9]", "", s)


def deep_cfg(cfg):
    out = json.loads(json.dumps(DEFAULTS))
    for k, v in cfg.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k].update(v)
        else:
            out[k] = v
    return out


# --------------------------------------------------------------------------
# Type Product
# --------------------------------------------------------------------------

_SIZE_THEN_JOIN = re.compile(r"\d\s?(?:ml|gr|g)\b\s*[&+]\s*\S", re.I)      # "50ml & EB Facial Foam"
_PLUS_JOIN = re.compile(r"\s\+\s|[A-Za-z]{2}\+[A-Za-z]{2}")                 # "Intensive + Night", "Protect+Night"


def type_source_text(sku_name, product_name):
    """SKU name adalah acuan; kalau kosong atau 'Default', pakai Product Name."""
    sku = "" if is_blank(sku_name) else str(sku_name).strip()
    if sku and sku.lower() != "default":
        return sku
    return "" if is_blank(product_name) else str(product_name).strip()


def classify_type(sku_name, product_name, cfg):
    """Return (Type Product, teks yang dipakai). Single / Multipack / Bundling [+ GWP]."""
    text = type_source_text(sku_name, product_name)
    labels = cfg["type_labels"]
    low = " ".join(text.lower().split())
    over = {" ".join(k.lower().split()): v for k, v in cfg["type_overrides"].items()}
    if low in over:
        return over[low], text

    gwp = any(re.search(r"(?<![a-z])" + re.escape(k.lower()), low) for k in cfg["gwp_keywords"])

    isi = re.search(r"\bisi\s*(\d+)", low)
    if isi:                                           # [ISI 1] = single, [ISI 2..] = multipack
        base = "single" if int(isi.group(1)) <= 1 else "multipack"
    elif re.search(r"\[\s*1\s*pcs\s*\]|\bsingle\b", low):
        base = "single"
    elif (_SIZE_THEN_JOIN.search(low) or _PLUS_JOIN.search(low)
          or any(re.search(r"(?<![a-z])" + re.escape(k) + r"(?![a-z])", low) for k in cfg["bundling_keywords"])):
        base = "bundling"
    elif (any(k in low for k in cfg["multipack_keywords"])
          or re.search(r"\bx\s?[2-9]\b|\b[2-9]\s?x\b", low)):
        base = "multipack"
    else:
        base = "single"
    label = labels[base] + (cfg["gwp_suffix"] if gwp else "")
    return label, text


# --------------------------------------------------------------------------
# Baca sumber
# --------------------------------------------------------------------------

def frame_from_rows(rows, filename, key="order id"):
    """rows = list of list. Cari baris header yang berisi `key` (ada baris catatan di atas file Partner)."""
    for i in range(min(len(rows), 30)):
        cells = [str(x).strip() for x in rows[i]]
        if key in [c.lower() for c in cells]:
            width = len(cells)
            body = [list(r)[:width] + [""] * (width - len(r)) for r in rows[i + 1:]
                    if any(str(x).strip() for x in r)]
            keep = [j for j, c in enumerate(cells) if c]
            df = pd.DataFrame([[r[j] for j in keep] for r in body], columns=[cells[j] for j in keep], dtype=object)
            return df
    raise ValueError(f"Kolom '{key}' tidak ditemukan di {filename}")


def read_source_bytes(content, filename, key="order id"):
    name = filename.lower()
    if name.endswith(".csv"):
        text = None
        for enc in ("utf-8-sig", "utf-16", "latin-1"):
            try:
                text = content.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError(f"Tidak bisa membaca CSV {filename}")
        try:
            dialect = csv.Sniffer().sniff(text[:5000], delimiters=",;\t")
            delim = dialect.delimiter
        except csv.Error:
            delim = ","
        rows = list(csv.reader(io.StringIO(text), delimiter=delim))
        return frame_from_rows(rows, filename, key)
    sheets = pd.read_excel(io.BytesIO(content), header=None, dtype=object, sheet_name=None)
    for _, raw in sheets.items():
        try:
            return frame_from_rows(raw.fillna("").values.tolist(), filename, key)
        except ValueError:
            continue
    raise ValueError(f"Kolom '{key}' tidak ditemukan di {filename}")


def resolve(columns, aliases):
    exact = {str(c).strip().lower(): c for c in columns}
    normed = {}
    for c in columns:
        normed.setdefault(_norm(c), c)
    for a in aliases:
        c = exact.get(a.strip().lower()) or normed.get(_norm(a))
        if c is not None:
            return c
    return None


def build_records(df, source_label, cfg, bank=None):
    """DataFrame export -> list dict nilai mentah per kolom output (belum diberi format)."""
    colmap = {out: resolve(df.columns, al) for out, al in FIELD_ALIASES.items()}
    missing = [k for k in ("Order ID", "Time Created") if colmap[k] is None]
    if missing:
        raise ValueError(f"Kolom wajib tidak ketemu: {', '.join(missing)}")
    records = []
    for _, r in df.iterrows():
        rec = {out: (r[c] if c is not None else "") for out, c in colmap.items()}
        if is_blank(rec["Order ID"]):
            continue
        rec["Source"] = source_label
        rec["Type Product"], rec["_type_text"] = classify_type(rec["SKU name"], rec["Product Name"], cfg)
        pid = clean_id(rec["Product ID"])
        rec["CAT"], rec["PRO"] = (bank or {}).get(pid, ("", ""))
        rec["_pid"] = pid
        records.append(rec)
    return records, [k for k, c in colmap.items() if c is None]


def to_cell(col, rec, cfg):
    v = rec.get(col, "")
    if col == "Time Created":
        t = parse_time(v, cfg["date_order"])
        return t.strftime("%Y-%m-%d %H:%M:%S") if t else ("" if is_blank(v) else str(v))
    if col in ID_COLS:
        return as_text(clean_id(v))
    if col in NUM_COLS:
        return to_number(v)
    if col in TEXT_COLS:
        return "" if is_blank(v) else safe_text(str(v).strip())
    return "" if is_blank(v) else to_number(v)  # kolom lain (mis. rate "20%")


def record_key(rec):
    return (rec["Source"], clean_id(rec["Order ID"]), clean_id(rec.get("Sku Id")))


# --------------------------------------------------------------------------
# Google API
# --------------------------------------------------------------------------

def load_credentials(credentials_file, token_file):
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
            creds = None
        if not creds or not creds.valid:
            print("   Browser akan terbuka untuk login Google.")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
        return creds
    if info.get("type") == "service_account":
        from google.oauth2.service_account import Credentials
        return Credentials.from_service_account_info(info, scopes=SCOPES)
    sys.exit("Format file kredensial tidak dikenali (pakai OAuth client Desktop app atau service account).")


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
                fields="nextPageToken, files(id, name, mimeType, modifiedTime, size)",
                supportsAllDrives=True, includeItemsFromAllDrives=True,
            ).execute()
            files += resp.get("files", [])
            token = resp.get("nextPageToken")
            if not token:
                break
        ok = [f for f in files if f["mimeType"] == GOOGLE_SHEET_MIME
              or f["name"].lower().endswith((".csv", ".xlsx", ".xls"))]
        return sorted(ok, key=lambda f: f["modifiedTime"], reverse=True)

    def read_file(self, meta):
        """Return DataFrame. Google Sheet dibaca lewat Sheets API (ID panjang tetap utuh sebagai teks)."""
        from googleapiclient.http import MediaIoBaseDownload

        if meta["mimeType"] == GOOGLE_SHEET_MIME:
            book = self.gc.open_by_key(meta["id"])
            last_err = None
            for ws in book.worksheets():
                try:
                    return frame_from_rows(ws.get_all_values(), meta["name"])
                except ValueError as e:
                    last_err = e
            raise last_err
        req = self.drive.files().get_media(fileId=meta["id"], supportsAllDrives=True)
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        return read_source_bytes(buf.getvalue(), meta["name"])


def bank_from_frame(df):
    cols = {str(c).strip().lower(): c for c in df.columns}
    pid_c, cat_c, pro_c = cols.get("product id") or cols.get("pid"), cols.get("cat"), cols.get("pro")
    if pid_c is None or cat_c is None:
        raise ValueError("Product Bank harus punya kolom Product ID dan CAT")
    bank = {}
    for _, r in df.iterrows():
        pid = clean_id(r[pid_c])
        cat = "" if is_blank(r[cat_c]) or str(r[cat_c]).startswith("#") else str(r[cat_c]).strip()
        pro = "" if pro_c is None or is_blank(r[pro_c]) or str(r[pro_c]).startswith("#") else str(r[pro_c]).strip()
        if pid and cat and pid not in bank:
            bank[pid] = (cat, pro)
    return bank


def read_product_bank(g, file_id, sheet_name):
    """Product Bank bisa berupa Google Sheet atau file .xlsx di Drive. Return dict PID -> (CAT, PRO)."""
    from googleapiclient.http import MediaIoBaseDownload

    meta = g.drive.files().get(fileId=file_id, fields="id, name, mimeType", supportsAllDrives=True).execute()
    if meta["mimeType"] == GOOGLE_SHEET_MIME:
        book = g.gc.open_by_key(file_id)
        order = [w for w in book.worksheets() if w.title == sheet_name] + \
                [w for w in book.worksheets() if w.title != sheet_name]
        for ws in order:
            try:
                return bank_from_frame(frame_from_rows(ws.get_all_values(), meta["name"], "product id"))
            except ValueError:
                continue
        raise ValueError(f"Kolom 'Product ID' tidak ketemu di Product Bank {meta['name']}")
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, g.drive.files().get_media(fileId=file_id, supportsAllDrives=True))
    done = False
    while not done:
        _, done = dl.next_chunk()
    return bank_from_frame(read_source_bytes(buf.getvalue(), meta["name"], "product id"))


# --------------------------------------------------------------------------
# Pilih file
# --------------------------------------------------------------------------

def pick_files(label, files, spec, interactive):
    """spec = teks dari --order/--partner (potongan nama atau ID, pisah koma); None = tanya/lewati."""
    if spec:
        chosen = []
        for part in [p.strip() for p in spec.split(",") if p.strip()]:
            hits = [f for f in files if f["id"] == part or part.lower() in f["name"].lower()]
            if not hits:
                sys.exit(f"File '{part}' tidak ditemukan di folder {label}.")
            chosen += hits
        return chosen
    if not interactive:
        return []
    if not files:
        print(f"   Folder {label} kosong.")
        return []
    print(f"\nFile di folder {label}:")
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f['name']}   (diubah {f['modifiedTime'][:16].replace('T', ' ')} UTC)")
    ans = input(f"Pilih nomor file {label} (pisah koma, Enter = lewati): ").strip()
    chosen = []
    for part in [p for p in re.split(r"[,\s]+", ans) if p]:
        if part.isdigit() and 1 <= int(part) <= len(files):
            chosen.append(files[int(part) - 1])
    return chosen


# --------------------------------------------------------------------------
# Tulis ke sheet
# --------------------------------------------------------------------------

def _numval(v):
    s = "" if v is None else str(v).lstrip("'").strip()
    try:
        return float(s[:-1]) / 100 if s.endswith("%") else float(s)
    except ValueError:
        return s


def same_value(a, b):
    x, y = _numval(a), _numval(b)
    if isinstance(x, float) and isinstance(y, float):
        return abs(x - y) < 1e-9
    return x == y


def write_csv(path, header, rows):
    try:
        f = open(path, "w", newline="", encoding="utf-8-sig")
    except PermissionError:
        path = path[:-4] + datetime.now().strftime("_%H%M%S") + ".csv"
        f = open(path, "w", newline="", encoding="utf-8-sig")
    with f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    return path


def type_review_csv(records, out_dir, label):
    counts = {}
    for r in records:
        k = (r["Type Product"], r["_type_text"])
        counts[k] = counts.get(k, 0) + 1
    rows = [[t, txt, n] for (t, txt), n in sorted(counts.items(), key=lambda x: (x[0][0], -x[1]))]
    return write_csv(os.path.join(out_dir, f"type_product_review_{label}.csv"),
                     ["Type Product", "Teks acuan (SKU name / Product Name)", "Jumlah baris"], rows)


# --------------------------------------------------------------------------
# Statistik & laporan
# --------------------------------------------------------------------------

def to_dt(v):
    """Time Created dari sheet: serial angka, teks ISO, atau dd/mm/yyyy."""
    if is_blank(v):
        return None
    if isinstance(v, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=float(v))
    return parse_time(v, "dmy")


def fmt_dt(t):
    return t.strftime("%d/%m/%Y %H:%M:%S") if t else "-"


def fmt_n(n):
    return f"{n:,.0f}"


def num_sum(values):
    return sum(v for v in (to_number(x) for x in values) if isinstance(v, (int, float)) and not isinstance(v, bool))


def rec_stats(records, cfg):
    times = [t for t in (parse_time(r["Time Created"], cfg["date_order"]) for r in records) if t]
    return {"rows": len(records), "orders": len({clean_id(r["Order ID"]) for r in records}),
            "tmin": min(times) if times else None, "tmax": max(times) if times else None,
            "est": num_sum(r.get("Est. Commission Base") for r in records)}


def sheet_stats(existing, index):
    """existing: key -> (nomor baris, nilai baris). Return Source -> statistik isi sheet saat ini."""
    out = {}
    i_t, i_e = index.get("time created"), index.get("est. commission base")
    for (src, _, _), (_, row) in existing.items():
        d = out.setdefault(src, {"rows": 0, "tmin": None, "tmax": None, "est": 0})
        d["rows"] += 1
        t = to_dt(row[i_t]) if i_t is not None else None
        if t:
            d["tmin"] = t if d["tmin"] is None else min(d["tmin"], t)
            d["tmax"] = t if d["tmax"] is None else max(d["tmax"], t)
        if i_e is not None:
            d["est"] += num_sum([row[i_e]])
    return out


def print_sheet_status(stats, labels):
    print("\nPosisi sheet output saat ini:")
    for lab in labels:
        d = stats.get(lab)
        if not d:
            print(f"   {lab:<10} belum ada data")
        else:
            print(f"   {lab:<10} {fmt_n(d['rows'])} baris, data {fmt_dt(d['tmin'])} s/d {fmt_dt(d['tmax'])}  "
                  f"(update terakhir sampai {d['tmax'].strftime('%d/%m/%Y') if d['tmax'] else '-'})")


def print_file_report(file_stats):
    print("\n=== Isi file yang dibaca (sebelum masuk ke output) ===")
    for lab, name, st in file_stats:
        print(f"   {lab:<10} {name}")
        print(f"              {fmt_n(st['rows'])} baris, {fmt_n(st['orders'])} order unik")
        print(f"              Time Created : {fmt_dt(st['tmin'])} s/d {fmt_dt(st['tmax'])}")
        print(f"              Est. Commission Base (jumlah): {fmt_n(st['est'])}")


# --------------------------------------------------------------------------
# Sheet output
# --------------------------------------------------------------------------

DEFAULT_HEADER = ["Source", "Time Created", "Order ID", "Product ID", "Product Name", "CAT", "PRO", "Sku Id",
                  "SKU name", "Type Product", "Seller SKU", "Price", "Payment Amount", "Quantity", "Payment method",
                  "Order Status", "Creator Username", "Content Type", "Carousel", "Content ID", "commission model",
                  "Standard commission rate", "Est. Commission Base", "Est. standard commission payment",
                  "Actual Commission Base", "Actual Commission Payment", "Partner campaign ID",
                  "Affiliate Partner ID", "Affiliate Partner Name", "Affiliate Partner Commission Rate",
                  "Est. Commission for Affiliate Partner", "Actual Commission for Affiliate Partner"]


def load_existing(ws, header_row):
    header = [h.strip() for h in ws.row_values(header_row)]
    if not header:
        sys.exit("Header tab output kosong.")
    index = {h.lower(): i for i, h in enumerate(header) if h}
    for k in KEY_COLS:
        if k.lower() not in index:
            sys.exit(f"Kolom kunci '{k}' tidak ada di header tab output.")
    width = len(header)
    rows = ws.get(f"A{header_row + 1}:{col_letter(width)}", value_render_option="UNFORMATTED_VALUE")
    existing = {}
    for n, row in enumerate(rows, start=header_row + 1):
        row = list(row) + [""] * (width - len(row))
        key = (str(row[index["source"]]).strip(), clean_id(row[index["order id"]]), clean_id(row[index["sku id"]]))
        if key[1]:
            existing[key] = (n, row)
    return header, index, existing


def plan_sync(header, index, existing, records, cfg, args):
    """Tentukan baris baru / baris lama yang berubah. Dobel (antar file maupun yang sudah ada di sheet) dilewati."""
    ordered = sorted(records, key=lambda r: r.get("_fmod", ""))  # file terbaru diproses terakhir -> menang
    uniq, dup_batch = {}, []
    for rec in ordered:
        k = record_key(rec)
        if k in uniq:
            dup_batch.append(uniq[k])
        uniq[k] = rec

    def build(rec):
        return [to_cell(h, rec, cfg) if h else "" for h in header]

    def tkey(kv):
        return parse_time(kv[1]["Time Created"], cfg["date_order"]) or datetime.min

    new_recs, new_rows, already, updates, n_changed = [], [], [], [], 0
    for key, rec in sorted(uniq.items(), key=tkey):
        if key not in existing:
            new_recs.append(rec)
            new_rows.append(build(rec))
            continue
        already.append(rec)
        if args.no_update_existing:
            continue
        n, old = existing[key]
        built = build(rec)
        changed = False
        for col in MUTABLE_COLS:
            i = index.get(col.lower())
            if i is None:
                continue
            if col in ("CAT", "PRO") and is_blank(built[i]):
                continue  # jangan hapus CAT/PRO yang sudah ada kalau Product Bank tidak punya
            if not same_value(old[i], built[i]):
                updates.append({"range": f"{col_letter(i + 1)}{n}", "values": [[built[i]]]})
                changed = True
        n_changed += changed
    cross = {}
    for (src, oid, sku) in uniq:
        cross.setdefault((oid, sku), set()).add(src)
    cross_n = sum(1 for v in cross.values() if len(v) > 1)
    return {"uniq": uniq, "dup_batch": dup_batch, "new_recs": new_recs, "new_rows": new_rows, "already": already,
            "updates": updates, "n_changed": n_changed, "cross": cross_n}


def print_plan_report(plan, before, labels, cfg):
    print("\n=== Yang akan masuk ke output ===")
    for title, recs in (("Dobel antar file yang kamu pilih (tanggal/order sama), dilewati", plan["dup_batch"]),
                        ("Sudah ada di sheet output, tidak dimasukkan lagi", plan["already"])):
        st = rec_stats(recs, cfg)
        rng = f", Time Created {fmt_dt(st['tmin'])} s/d {fmt_dt(st['tmax'])}" if recs else ""
        print(f"   {title}: {fmt_n(len(recs))} baris{rng}")
    if plan["cross"]:
        print(f"   ⚠️  {fmt_n(plan['cross'])} pasangan Order ID + Sku Id muncul di dua Source sekaligus "
              f"(AFFILIATE dan PATNER). Keduanya tetap masuk; cek kalau itu tidak wajar.")
    print(f"   Baris lama yang akan diperbarui (status/komisi berubah): {fmt_n(plan['n_changed'])}")
    after = {}
    for lab in labels:
        b = before.get(lab, {"rows": 0, "tmax": None, "est": 0})
        new = [r for r in plan["new_recs"] if r["Source"] == lab]
        st = rec_stats(new, cfg)
        tmax = max([t for t in (b["tmax"], st["tmax"]) if t], default=None)
        after[lab] = {"rows": b["rows"] + len(new), "tmax": tmax, "est": b["est"] + st["est"],
                      "tmin": b.get("tmin") if b.get("tmin") else st["tmin"]}
        print(f"\n   {lab}")
        print(f"      Sheet sekarang : {fmt_n(b['rows']):>10} baris | Est. Commission Base {fmt_n(b['est'])}")
        print(f"      Masuk (baru)   : {fmt_n(len(new)):>10} baris | Est. Commission Base +{fmt_n(st['est'])}")
        print(f"      Sheet sesudah  : {fmt_n(after[lab]['rows']):>10} baris | Est. Commission Base {fmt_n(after[lab]['est'])}")
        print(f"      Data sampai    : {fmt_dt(tmax)}")
    return after


def with_retry(fn, *a, **k):
    """Google membatasi ~60 permintaan tulis per menit. Kalau kena batas (429) atau error sementara, tunggu lalu ulangi."""
    for attempt in range(8):
        try:
            return fn(*a, **k)
        except Exception as e:
            code = getattr(e, "code", None) or getattr(getattr(e, "response", None), "status_code", None)
            if code in (429, 500, 502, 503) and attempt < 7:
                wait = min(70, 20 * (attempt + 1))
                print(f"   ⏳ Kena batas kuota Google ({code}), menunggu {wait} detik lalu lanjut ...")
                time.sleep(wait)
                continue
            raise


def merge_cell_updates(updates):
    """Gabungkan update sel tunggal yang berurutan dalam satu kolom jadi satu rentang (jauh lebih sedikit permintaan)."""
    cols, out = {}, []
    for u in updates:
        m = re.fullmatch(r"([A-Z]+)(\d+)", u["range"])
        if m:
            cols.setdefault(m.group(1), []).append((int(m.group(2)), u["values"][0][0]))
        else:
            out.append(u)
    for c, items in cols.items():
        items.sort(key=lambda x: x[0])
        start = prev = None
        vals = []
        for r, v in items:
            if start is not None and r == prev + 1:
                vals.append([v])
                prev = r
                continue
            if start is not None:
                out.append({"range": f"{c}{start}:{c}{prev}", "values": vals})
            start = prev = r
            vals = [[v]]
        if start is not None:
            out.append({"range": f"{c}{start}:{c}{prev}", "values": vals})
    return out


def get_or_create(book, title, rows=1000, cols=12):
    try:
        return book.worksheet(title)
    except Exception:
        return book.add_worksheet(title=title, rows=rows, cols=cols)


LAST_HEAD = ["Source", "Data mulai", "Data sampai", "Tanggal terakhir", "Total baris", "Σ (ukuran)", "Ukuran Σ",
             "Terakhir diupdate", "File terakhir dimasukkan"]
LOG_HEAD = ["Waktu run", "Source", "File", "Baris dibaca", "Data dari", "Data sampai", "Σ file", "Baris masuk",
            "Σ masuk", "Baris dilewati (dobel)", "Baris lama diperbarui (total run)"]


def upsert_last_update(book, cfg, rows):
    """Tab 'Last Update': satu baris per dataset (AFFILIATE, PATNER, PRODUCT PERFORMANCE). Baris dataset lain tidak disentuh."""
    ws = get_or_create(book, cfg["merge_last_update_sheet"], 50, len(LAST_HEAD))
    col = ws.col_values(1)
    updates = []
    if not col or col[0] != LAST_HEAD[0]:
        updates.append({"range": f"A1:{col_letter(len(LAST_HEAD))}1", "values": [LAST_HEAD]})
    pos = {v: i + 1 for i, v in enumerate(col) if v}
    nxt = max(len(col), 1) + 1
    for r in rows:
        n = pos.get(r[0])
        if n is None:
            n, nxt = nxt, nxt + 1
            pos[r[0]] = n
        updates.append({"range": f"A{n}:{col_letter(len(LAST_HEAD))}{n}", "values": [r]})
    ws.batch_update(updates, value_input_option="USER_ENTERED")


def append_log(book, cfg, entries, n_updated, fmt=None):
    fmt = fmt or fmt_dt
    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    log = get_or_create(book, cfg["merge_log_sheet"], 2000, len(LOG_HEAD))
    if not log.get("A1"):
        log.batch_update([{"range": f"A1:{col_letter(len(LOG_HEAD))}1", "values": [LOG_HEAD]}],
                         value_input_option="USER_ENTERED")
    rows = [[as_text(now), e["source"], safe_text(e["file"]), e["rows"], as_text(fmt(e["tmin"])), as_text(fmt(e["tmax"])),
             round(e["est"]), e["masuk"], round(e["masuk_est"]), e["rows"] - e["masuk"], n_updated] for e in entries]
    log.append_rows(rows, value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS")


def last_row(label, a, files, metric_name, fmt=None):
    fmt = fmt or fmt_dt
    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    return [label, as_text(fmt(a.get("tmin"))) if a else "", as_text(fmt(a["tmax"])) if a else "",
            as_text(a["tmax"].strftime("%d/%m/%Y")) if a and a["tmax"] else "", a["rows"] if a else 0,
            round(a["est"]) if a else 0, metric_name, as_text(now), safe_text("; ".join(files))]


def write_update_info(book, cfg, labels, after, file_entries, plan):
    """Tab 'Last Update' (ringkas) dan 'Update Log' (riwayat, ditambah tiap run) untuk data Order."""
    rows = [last_row(lab, after.get(lab),
                     [e["file"] for e in file_entries if e["source"] == lab and e["masuk"] > 0],
                     "Est. Commission Base") for lab in labels]
    upsert_last_update(book, cfg, rows)
    append_log(book, cfg, file_entries, plan["n_changed"])


def write_to_sheet(ws, book, header, index, plan, after, file_entries, labels, cfg):
    last = col_letter(len(header))
    header_row = int(cfg["merge_header_row"])
    updates = merge_cell_updates(plan["updates"])
    n_req = (len(updates) + 999) // 1000
    if n_req > 1:
        print(f"   Menulis perubahan baris lama dalam {n_req} permintaan (diberi jeda supaya tidak kena batas kuota Google) ...")
    for i in range(0, len(updates), 1000):
        with_retry(ws.batch_update, updates[i:i + 1000], value_input_option="USER_ENTERED")
        if n_req > 1:
            time.sleep(2)
    if updates:
        print(f"✅ {plan['n_changed']} baris lama diperbarui ({len(updates)} sel).")
    rows = plan["new_rows"]
    for i in range(0, len(rows), 5000):
        chunk = rows[i:i + 5000]
        resp = with_retry(ws.append_rows, chunk, value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS",
                          table_range=f"A{header_row}:{last}{header_row}")
        m = re.search(r"!A(\d+):", resp.get("updates", {}).get("updatedRange", ""))
        if m and "time created" in index:
            s = int(m.group(1))
            c = col_letter(index["time created"] + 1)
            with_retry(ws.batch_format, [{"range": f"{c}{s}:{c}{s + len(chunk) - 1}",
                                          "format": {"numberFormat": {"type": "DATE_TIME", "pattern": "dd/mm/yyyy hh:mm:ss"}}}])
        time.sleep(2)
    if rows:
        print(f"✅ {len(rows)} baris baru ditambahkan ke {ws.title}.")
    try:
        write_update_info(book, cfg, labels, after, file_entries, plan)
        print(f"✅ Tab '{cfg['merge_last_update_sheet']}' dan '{cfg['merge_log_sheet']}' diperbarui.")
    except Exception as e:  # data utama sudah masuk; info update tidak boleh menggagalkan run
        print(f"⚠️  Data sudah masuk, tapi tab info update gagal ditulis: {e}")


# --------------------------------------------------------------------------
# Product Performance Daily (Raw): gabungan file harian product_list_YYYYMMDD.xlsx
# --------------------------------------------------------------------------

_FILE_DATE_RE = re.compile(r"(20\d{2})(\d{2})(\d{2})")
_ANALYSIS_RE = re.compile(r"(\d{2}/\d{2}/\d{4})\s*\\?~\s*(\d{2}/\d{2}/\d{4})")
_NUMERIC_TEXT_RE = re.compile(r"^(?:Rp|IDR)?\s*-?[\d][\d.,]*$", re.I)


def date_from_name(name):
    m = _FILE_DATE_RE.search(name)
    if not m:
        return None
    try:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def parse_number_text(s):
    """'Rp45.540.460' -> 45540460, '1,234' -> 1234, '12.5' -> 12.5. Return None kalau bukan angka murni."""
    s = s.strip().replace("Rp", "").replace("IDR", "").replace(" ", "")
    if not s or not re.fullmatch(r"-?[\d.,]+", s):
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        s = s.replace(",", "") if all(len(p) == 3 for p in parts[1:]) else s.replace(",", ".")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[1]) == 3):
        s = s.replace(".", "")
    try:
        n = float(s)
    except ValueError:
        return None
    return int(n) if n.is_integer() else n


def convert_product_cell(v):
    """Nilai sel file product_list -> nilai untuk sheet. Angka (termasuk 'Rp1.234') jadi angka, '5.88%' tetap, teks tetap."""
    if is_blank(v):
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else v
    s = str(v).strip()
    if re.fullmatch(r"-+", s) or s.upper() in ("N/A", "NAN"):
        return ""
    if s.endswith("%"):
        return s
    if _NUMERIC_TEXT_RE.match(s):
        n = parse_number_text(s)
        if n is not None:
            return n
    return safe_text(s)


def read_product_file(content, filename):
    """Return dict header, group (di-forward-fill), body (list baris), analysis (teks tanggal analisis)."""
    sheets = pd.read_excel(io.BytesIO(content), header=None, dtype=object, sheet_name=None)
    for _, raw in sheets.items():
        rows = raw.values.tolist()
        for i in range(min(len(rows), 30)):
            cells = ["" if is_blank(x) else str(x).strip() for x in rows[i]]
            low = [c.lower() for c in cells]
            if "product id" in low and "product name" in low:
                groups, cur = [], ""
                prev = rows[i - 1] if i > 0 else []
                for j in range(len(cells)):
                    gv = prev[j] if j < len(prev) else ""
                    cur = "" if (j == 0 and is_blank(gv)) else (str(gv).strip() if not is_blank(gv) else cur)
                    groups.append(cur)
                analysis = ""
                for r in rows[:i]:
                    for x in r:
                        if not is_blank(x) and _ANALYSIS_RE.search(str(x)):
                            analysis = str(x)
                body = [r for r in rows[i + 1:] if any(not is_blank(x) for x in r)]
                return {"header": cells, "group": groups, "body": body, "analysis": analysis}
    raise ValueError(f"Header 'Product ID' tidak ditemukan di {filename}")


def column_keys(groups, headers):
    """Kunci unik per kolom: (grup, nama, urutan kemunculan). Nama kolom yang sama di grup berbeda tidak tertukar."""
    seen, keys = {}, []
    for g_, h_ in zip(groups, headers):
        k = (_norm(g_), _norm(h_))
        seen[k] = seen.get(k, 0) + 1
        keys.append((k, seen[k]))
    return keys


def product_file_date(name, pf):
    """Tanggal dari nama file (acuan utama). Tanggal 'Analysis date' di dalam file dipakai sebagai pengecekan."""
    d = date_from_name(name)
    m = _ANALYSIS_RE.search(pf["analysis"]) if pf["analysis"] else None
    a_start = a_end = None
    if m:
        a_start = datetime.strptime(m.group(1), "%d/%m/%Y")
        a_end = datetime.strptime(m.group(2), "%d/%m/%Y")
    note = ""
    if d is None:
        if a_start and a_start == a_end:
            d, note = a_start, "tanggal diambil dari 'Analysis date' di dalam file (nama file tidak ada tanggal)"
        else:
            raise ValueError(f"{name}: tanggal tidak ketemu di nama file dan Analysis date bukan satu hari")
    elif a_start and (a_start != d or a_end != d):
        note = (f"⚠️ tanggal di nama file {d:%d/%m/%Y} beda dengan Analysis date di dalam file "
                f"{m.group(1)}~{m.group(2)}; dipakai tanggal nama file")
    return d, note


def build_product_records(pf, fdate, out_groups, out_header, bank, fname, fmod, cfg):
    """Susun baris sesuai urutan kolom tab output. Return (records, kolom output yang tidak ketemu di file)."""
    fkeys = column_keys(pf["group"], pf["header"])
    fidx = {k: j for j, k in enumerate(fkeys)}
    okeys = column_keys(out_groups, out_header)
    low = [h.strip().lower() for h in out_header]
    special = {"tanggal", "product name", "product id", "cat", "pro"}
    src = []
    missing = []
    for j, (k, h) in enumerate(zip(okeys, low)):
        if h in special:
            src.append(None)
            continue
        f = fidx.get(k)
        if f is None:  # kolom di luar grup: cari tanpa grup (nama + urutan)
            f = next((jj for kk, jj in fidx.items() if kk[0][1] == k[0][1] and kk[1] == k[1]), None)
        src.append(f)
        if f is None and out_header[j].strip():
            missing.append(out_header[j])
    i_name = pf["header"].index(next(h for h in pf["header"] if h.lower() == "product name"))
    i_pid = pf["header"].index(next(h for h in pf["header"] if h.lower() == "product id"))
    gmv_all = fidx.get((("all", "gmv"), 1))
    gmv_aff = fidx.get((("affiliate", "attributedgmv"), 1))
    recs = []
    for r in pf["body"]:
        pid = clean_id(r[i_pid]) if i_pid < len(r) else ""
        if not pid:
            continue
        cat, pro = bank.get(pid, ("", ""))
        row = []
        for j, h in enumerate(low):
            if h == "tanggal":
                row.append(fdate.strftime("%Y-%m-%d"))
            elif h == "product name":
                row.append(safe_text("" if is_blank(r[i_name]) else str(r[i_name]).strip()))
            elif h == "product id":
                row.append(as_text(pid))
            elif h == "cat":
                row.append(safe_text(cat))
            elif h == "pro":
                row.append(safe_text(pro))
            elif src[j] is None or src[j] >= len(r):
                row.append("")
            else:
                row.append(convert_product_cell(r[src[j]]))
        recs.append({"date": fdate, "pid": pid, "row": row, "cat": cat, "name": str(r[i_name]), "ord": len(recs),
                     "gmv": convert_num_only(r, gmv_all), "aff": convert_num_only(r, gmv_aff),
                     "_file": fname, "_fmod": fmod})
    return recs, missing


def convert_num_only(r, f):
    if f is None or f >= len(r) or is_blank(r[f]):
        return 0
    v = convert_product_cell(r[f])
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def pstats(records):
    ds = [r["date"] for r in records]
    return {"rows": len(records), "tmin": min(ds) if ds else None, "tmax": max(ds) if ds else None,
            "est": sum(r["gmv"] for r in records), "aff": sum(r["aff"] for r in records),
            "dates": sorted({d.date() for d in ds})}


def fmt_date(t):
    return t.strftime("%d/%m/%Y") if t else "-"


def detect_header_row(ws, fallback, max_rows=15):
    """Cari baris header tab (yang berisi 'Tanggal' dan 'Product ID'), jadi tetap jalan kalau kamu menambah/mengurangi baris di atasnya."""
    rows = ws.get(f"A1:{col_letter(200)}{max_rows}")
    for n, row in enumerate(rows, start=1):
        low = [str(x).strip().lower() for x in row]
        if "tanggal" in low and "product id" in low:
            return n
    return fallback


def load_product_sheet(ws, header_row):
    """Header (baris grup + baris kolom) dan baris yang sudah ada: key (tanggal, PID) -> (nomor baris, GMV)."""
    groups = ws.row_values(header_row - 1) if header_row > 1 else []
    header = [h.strip() for h in ws.row_values(header_row)]
    if not header:
        sys.exit("Header tab Product Performance (Raw) kosong.")
    groups = [(groups[j].strip() if j < len(groups) else "") for j in range(len(header))]
    cur, g2 = "", []
    for gv in groups:
        cur = gv or cur
        g2.append(cur)
    low = [h.lower() for h in header]
    for need in ("tanggal", "product id"):
        if need not in low:
            sys.exit(f"Kolom '{need}' tidak ada di header tab Product Performance (Raw).")
    i_d, i_p = low.index("tanggal"), low.index("product id")
    okeys = column_keys(g2, header)
    i_g = next((j for j, (k, n) in enumerate(okeys) if k == ("all", "gmv") and n == 1), None)
    need_cols = [i_d, i_p] + ([i_g] if i_g is not None else [])
    last = col_letter(max(need_cols) + 1)
    rows = ws.get(f"A{header_row + 1}:{last}", value_render_option="UNFORMATTED_VALUE")
    existing = {}
    for n, row in enumerate(rows, start=header_row + 1):
        row = list(row) + [""] * (max(need_cols) + 1 - len(row))
        d = to_dt(row[i_d])
        pid = clean_id(row[i_p])
        if d and pid:
            existing[(d.date(), pid)] = (n, num_sum([row[i_g]]) if i_g is not None else 0)
    return g2, header, existing


def sheet_pstats(existing):
    ds = [k[0] for k in existing]
    return {"rows": len(existing), "tmin": datetime.combine(min(ds), datetime.min.time()) if ds else None,
            "tmax": datetime.combine(max(ds), datetime.min.time()) if ds else None,
            "est": sum(v[1] for v in existing.values()), "dates": sorted(set(ds))}


def list_product_files(g, folder_id):
    files = g.list_folder(folder_id)
    for f in files:
        f["_date"] = date_from_name(f["name"])
    return sorted(files, key=lambda f: (f["_date"] or datetime.min, f["name"]))


def pick_product_files(files, spec, existing_dates, interactive):
    """spec: potongan nama (mis. 20261105), pisah koma. Interaktif: nomor, 'b' = semua yang belum masuk, 'a' = semua."""
    if spec:
        chosen = []
        for part in [p.strip() for p in spec.split(",") if p.strip()]:
            hits = [f for f in files if f["id"] == part or part.lower() in f["name"].lower()]
            if not hits:
                sys.exit(f"File '{part}' tidak ditemukan di folder Product Performance.")
            chosen += hits
        return chosen
    if not interactive:
        return []
    if not files:
        print("   Folder Product Performance kosong.")
        return []
    print("\nFile di folder Product Performance:")
    for i, f in enumerate(files, 1):
        d = f["_date"]
        flag = "" if d is None else ("   (tanggal ini sudah ada di sheet)" if d.date() in existing_dates else "   (BARU)")
        print(f"  {i}. {f['name']}{flag}")
    ans = input("Pilih nomor file (pisah koma / rentang 1-3), 'b' = semua yang BARU, 'a' = semua, Enter = lewati: ").strip().lower()
    if ans == "a":
        return files
    if ans == "b":
        return [f for f in files if f["_date"] is not None and f["_date"].date() not in existing_dates]
    chosen = []
    for part in [p for p in re.split(r"[,\s]+", ans) if p]:
        m = re.fullmatch(r"(\d+)-(\d+)", part)
        if m:
            chosen += [files[k - 1] for k in range(int(m.group(1)), int(m.group(2)) + 1) if 1 <= k <= len(files)]
        elif part.isdigit() and 1 <= int(part) <= len(files):
            chosen.append(files[int(part) - 1])
    return chosen


def run_product(g, cfg, args, bank, out_dir):
    label = cfg["product_dataset_label"]
    sheet_name = args.product_sheet or cfg["product_output_sheet"]
    header_row = int(cfg["product_header_row"])
    book = ws = None
    out_groups, out_header, existing = None, None, {}
    if g is not None:
        book = g.gc.open_by_key(cfg["merge_output_spreadsheet_id"])
        ws = book.worksheet(sheet_name)
        detected = detect_header_row(ws, header_row)
        if detected != header_row:
            print(f"   Header tab '{sheet_name}' terdeteksi di baris {detected} (data mulai baris {detected + 1}).")
        header_row = detected
        out_groups, out_header, existing = load_product_sheet(ws, header_row)
    before = sheet_pstats(existing)
    print(f"\nPosisi tab '{sheet_name}':")
    if before["rows"]:
        print(f"   {fmt_n(before['rows'])} baris, {len(before['dates'])} tanggal, {fmt_date(before['tmin'])} s/d "
              f"{fmt_date(before['tmax'])} (update terakhir sampai {fmt_date(before['tmax'])}), jumlah GMV {fmt_n(before['est'])}")
    else:
        print("   belum ada data")

    # pilih & baca file
    frames = []  # (nama, dict file, tanggal, waktu modifikasi)
    if args.product_path:
        for path in [p.strip() for p in args.product_path.split(",") if p.strip()]:
            with open(path, "rb") as f:
                frames.append((os.path.basename(path), f.read(), datetime.fromtimestamp(os.path.getmtime(path)).isoformat()))
    else:
        files = list_product_files(g, cfg["product_folder_id"])
        interactive = not args.product
        for meta in pick_product_files(files, args.product, set(before.get("dates", [])), interactive):
            print(f"   membaca {meta['name']} ...")
            from googleapiclient.http import MediaIoBaseDownload
            buf = io.BytesIO()
            dl = MediaIoBaseDownload(buf, g.drive.files().get_media(fileId=meta["id"], supportsAllDrives=True))
            done = False
            while not done:
                _, done = dl.next_chunk()
            frames.append((meta["name"], buf.getvalue(), meta["modifiedTime"]))
    if not frames:
        print("Tidak ada file Product Performance yang dipilih.")
        return

    if out_header is None:  # tes tanpa Google: pakai header dari file pertama
        pf0 = read_product_file(frames[0][1], frames[0][0])
        out_groups = ["", "", "", "", ""] + pf0["group"][2:]
        out_header = ["Tanggal", "Product Name", "Product ID", "CAT", "PRO"] + [h for h in pf0["header"] if h.lower() not in ("product name", "product id")]
        out_groups = ["", "", "", "", ""] + [gp for gp, h in zip(pf0["group"], pf0["header"]) if h.lower() not in ("product name", "product id")]

    records, file_stats = [], []
    for name, content, fmod in frames:
        pf = read_product_file(content, name)
        fdate, note = product_file_date(name, pf)
        recs, missing = build_product_records(pf, fdate, out_groups, out_header, bank, name, fmod, cfg)
        if note:
            print(f"   {name}: {note}")
        if missing:
            print(f"   ⚠️  {name}: {len(missing)} kolom di tab output tidak ada di file (dikosongkan): "
                  f"{', '.join(missing[:6])}{' ...' if len(missing) > 6 else ''}")
        file_stats.append((name, fdate, pstats(recs)))
        records += recs

    print("\n=== Isi file yang dibaca (sebelum masuk ke output) ===")
    for name, fdate, st in file_stats:
        print(f"   {name}")
        print(f"      Tanggal : {fmt_date(fdate)}  |  {fmt_n(st['rows'])} produk")
        print(f"      Jumlah GMV (All): {fmt_n(st['est'])}  |  Jumlah Affiliate Attributed GMV: {fmt_n(st['aff'])}")

    # buang dobel: tanggal yang sama di beberapa file -> hanya file terbaru yang dipakai
    by_date = {}
    for r in records:
        by_date.setdefault(r["date"].date(), {}).setdefault((r["_file"], r["_fmod"]), []).append(r)
    kept, dropped_files = [], []
    for d, per_file in sorted(by_date.items()):
        winner = max(per_file, key=lambda k: (k[1], k[0]))
        for k, rs in per_file.items():
            if k == winner:
                kept += rs
            else:
                dropped_files.append((d, k[0], len(rs), winner[0]))
    # dobel baris (Tanggal, PID) di dalam satu file: yang terakhir menang
    uniq = {}
    for r in kept:
        uniq[(r["date"].date(), r["pid"])] = r
    ordered = sorted(uniq.items(), key=lambda kv: (kv[0][0], kv[1]["ord"]))  # urutan produk seperti di file
    new_recs = [r for k, r in ordered if k not in existing]
    old_recs = [r for k, r in ordered if k in existing]
    refreshed = old_recs if args.refresh else []

    print("\n=== Yang akan masuk ke output ===")
    for d, nm, n, win in dropped_files:
        print(f"   Tanggal {d:%d/%m/%Y} ada di beberapa file: {nm} ({fmt_n(n)} baris) dilewati, dipakai {win}")
    print(f"   Sudah ada di sheet (Tanggal + Product ID sama): {fmt_n(len(old_recs))} baris "
          f"{'-> DIPERBARUI (--refresh)' if args.refresh else '-> dilewati, tidak dimasukkan lagi'}")
    st_new = pstats(new_recs)
    after = {"rows": before["rows"] + len(new_recs), "est": before["est"] + st_new["est"],
             "tmin": min([t for t in (before["tmin"], st_new["tmin"]) if t], default=None),
             "tmax": max([t for t in (before["tmax"], st_new["tmax"]) if t], default=None)}
    print(f"\n   {label}")
    print(f"      Sheet sekarang : {fmt_n(before['rows']):>10} baris | Jumlah GMV {fmt_n(before['est'])}")
    print(f"      Masuk (baru)   : {fmt_n(len(new_recs)):>10} baris | Jumlah GMV +{fmt_n(st_new['est'])}"
          f"  (tanggal: {', '.join(d.strftime('%d/%m') for d in st_new['dates']) or '-'})")
    print(f"      Sheet sesudah  : {fmt_n(after['rows']):>10} baris | Jumlah GMV {fmt_n(after['est'])}")
    print(f"      Data sampai    : {fmt_date(after['tmax'])}")

    noc = {}
    for r in new_recs + refreshed:
        if not r["cat"]:
            noc.setdefault(r["pid"], [r["name"], 0])[1] += 1
    if noc:
        path = write_csv(os.path.join(out_dir, f"pid_belum_ada_cat_pro_product_{datetime.now():%Y%m%d_%H%M%S}.csv"),
                         ["Product ID", "Product Name", "Jumlah baris"], [[p, n, c] for p, (n, c) in noc.items()])
        print(f"   ⚠️  {len(noc)} Product ID belum ada CAT/PRO di Product Bank (dikosongkan): {path}")

    preview = write_csv(os.path.join(out_dir, f"preview_product_{datetime.now():%Y%m%d_%H%M%S}.csv"), out_header,
                        [[str(x).lstrip("'") for x in r["row"]] for r in new_recs])
    if args.dry_run:
        print(f"\n(dry-run) Tidak ada yang ditulis. Preview baris baru: {preview}")
        return
    if not new_recs and not refreshed:
        print("\nTidak ada perubahan.")
        return
    if not args.yes:
        if input(f"\nLanjut ke tab {ws.title}: tambah {fmt_n(len(new_recs))} baris"
                 f"{', perbarui ' + fmt_n(len(refreshed)) + ' baris lama' if refreshed else ''}? (y/n): ").strip().lower() != "y":
            print("Dibatalkan.")
            return

    width = len(out_header)
    last = col_letter(width)
    if refreshed:
        ups = [{"range": f"A{existing[(r['date'].date(), r['pid'])][0]}:{last}{existing[(r['date'].date(), r['pid'])][0]}",
                "values": [r["row"]]} for r in refreshed]
        for i in range(0, len(ups), 200):
            with_retry(ws.batch_update, ups[i:i + 200], value_input_option="USER_ENTERED")
            time.sleep(2)
        print(f"✅ {len(refreshed)} baris lama diperbarui.")
    i_date = [h.lower() for h in out_header].index("tanggal")
    for i in range(0, len(new_recs), 1000):
        chunk = [r["row"] for r in new_recs[i:i + 1000]]
        resp = with_retry(ws.append_rows, chunk, value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS",
                          table_range=f"A{header_row}:{last}{header_row}")
        m = re.search(r"!A(\d+):", resp.get("updates", {}).get("updatedRange", ""))
        if m:
            s0 = int(m.group(1))
            c = col_letter(i_date + 1)
            with_retry(ws.batch_format, [{"range": f"{c}{s0}:{c}{s0 + len(chunk) - 1}",
                                          "format": {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"}}}])
        time.sleep(2)
    if new_recs:
        print(f"✅ {len(new_recs)} baris baru ditambahkan ke {ws.title}.")
    try:
        entries = []
        for name, fdate, st in file_stats:
            mine = [r for r in new_recs if r["_file"] == name]
            entries.append({"source": label, "file": name, "rows": st["rows"], "tmin": st["tmin"], "tmax": st["tmax"],
                            "est": st["est"], "masuk": len(mine), "masuk_est": sum(r["gmv"] for r in mine)})
        upsert_last_update(book, cfg, [last_row(label, after, [e["file"] for e in entries if e["masuk"] > 0],
                                                "GMV (All)", fmt_date)])
        append_log(book, cfg, entries, len(refreshed), fmt_date)
        print(f"✅ Tab '{cfg['merge_last_update_sheet']}' dan '{cfg['merge_log_sheet']}' diperbarui.")
    except Exception as e:
        print(f"⚠️  Data sudah masuk, tapi tab info update gagal ditulis: {e}")


# --------------------------------------------------------------------------
# Order (Affiliate + Patner)
# --------------------------------------------------------------------------

def run_orders(g, cfg, args, bank, out_dir, labels):
    local = bool(args.order_path or args.partner_path)
    book = ws = None
    header, index, existing = list(DEFAULT_HEADER), {h.lower(): i for i, h in enumerate(DEFAULT_HEADER)}, {}
    if g is not None:
        book = g.gc.open_by_key(cfg["merge_output_spreadsheet_id"])
        ws = book.worksheet(cfg["merge_output_sheet"])
        header, index, existing = load_existing(ws, int(cfg["merge_header_row"]))
    before = sheet_stats(existing, index)
    print_sheet_status(before, labels)

    frames = []  # (label, nama file, DataFrame, waktu modifikasi)
    if local:
        for label, path in ((labels[0], args.order_path), (labels[1], args.partner_path)):
            if path:
                with open(path, "rb") as f:
                    mod = datetime.fromtimestamp(os.path.getmtime(path)).isoformat()
                    frames.append((label, os.path.basename(path), read_source_bytes(f.read(), path), mod))
    else:
        interactive = not (args.order or args.partner)
        for label, folder, spec, nm in ((labels[0], cfg["affiliate_order_folder_id"], args.order, "Affiliate_Order"),
                                        (labels[1], cfg["affiliate_partner_folder_id"], args.partner, "Affiliate_Patner")):
            for meta in pick_files(nm, g.list_folder(folder), spec, interactive):
                print(f"   membaca {meta['name']} ...")
                frames.append((label, meta["name"], g.read_file(meta), meta["modifiedTime"]))
    if not frames:
        print("Tidak ada file Order yang dipilih.")
        return

    records, file_stats = [], []
    for label, name, df, mod in frames:
        recs, absent = build_records(df, label, cfg, bank)
        for r in recs:
            r["_file"], r["_fmod"] = name, mod
        if absent:
            print(f"   ({name}: kolom tidak ada di file, dikosongkan: {', '.join(absent)})")
        file_stats.append((label, name, rec_stats(recs, cfg)))
        records += recs
    print_file_report(file_stats)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    print(f"\n   Laporan Type Product: {type_review_csv(records, out_dir, stamp)}")
    summary = {}
    for r in records:
        summary[r["Type Product"]] = summary.get(r["Type Product"], 0) + 1
    print("   Type Product: " + ", ".join(f"{k} {fmt_n(v)}" for k, v in sorted(summary.items())))
    nocat = {}
    for r in records:
        if r["_pid"] and not r["CAT"]:
            nocat.setdefault(r["_pid"], [str(r["Product Name"]), 0])[1] += 1
    if nocat:
        path = write_csv(os.path.join(out_dir, f"pid_belum_ada_cat_pro_{stamp}.csv"), ["Product ID", "Product Name", "Jumlah baris"],
                         [[p, n, c] for p, (n, c) in sorted(nocat.items(), key=lambda x: -x[1][1])])
        print(f"   ⚠️  {len(nocat)} Product ID belum ada CAT/PRO di Product Bank (CAT/PRO dikosongkan): {path}")

    plan = plan_sync(header, index, existing, records, cfg, args)
    after = print_plan_report(plan, before, labels, cfg)

    new_by_file = {}
    for r in plan["new_recs"]:
        d = new_by_file.setdefault((r["Source"], r["_file"]), [0, 0])
        d[0] += 1
        d[1] += num_sum([r.get("Est. Commission Base")])
    file_entries = []
    for label, name, st in file_stats:
        m = new_by_file.get((label, name), [0, 0])
        file_entries.append({"source": label, "file": name, "rows": st["rows"], "tmin": st["tmin"], "tmax": st["tmax"],
                             "est": st["est"], "masuk": m[0], "masuk_est": m[1]})

    preview = write_csv(os.path.join(out_dir, f"preview_merge_{stamp}.csv"), header,
                        [[str(x).lstrip("'") for x in r] for r in plan["new_rows"]])
    if args.dry_run:
        print(f"\n(dry-run) Tidak ada yang ditulis. Preview baris baru: {preview}")
        return
    if not plan["new_rows"] and not plan["updates"]:
        print("\nTidak ada perubahan.")
        return
    if not args.yes:
        if input(f"\nLanjut ke tab {ws.title}: tambah {fmt_n(len(plan['new_rows']))} baris, "
                 f"perbarui {fmt_n(plan['n_changed'])} baris lama? (y/n): ").strip().lower() != "y":
            print("Dibatalkan.")
            return
    write_to_sheet(ws, book, header, index, plan, after, file_entries, labels, cfg)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Gabungkan Affiliate_Order + Affiliate_Patner, dan file harian Product Performance")
    ap.add_argument("--config", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json"))
    ap.add_argument("--mode", choices=["orders", "product", "all"],
                    help="orders = Affiliate_Order + Patner, product = Product Performance Daily (Raw), all = keduanya")
    ap.add_argument("--order", help="File Affiliate_Order: potongan nama atau ID (pisah koma kalau lebih dari satu)")
    ap.add_argument("--partner", help="File Affiliate_Patner: potongan nama atau ID (pisah koma)")
    ap.add_argument("--product", help="File Product Performance harian: potongan nama, mis. 20261105,20261106")
    ap.add_argument("--product-sheet", help="Nama tab output Product Performance (default dari config)")
    ap.add_argument("--refresh", action="store_true",
                    help="Product Performance: baris (Tanggal + Product ID) yang sudah ada ditimpa dengan data file terbaru")
    ap.add_argument("--order-path", help="Pakai file lokal (csv/xlsx) untuk Affiliate_Order, bukan dari Drive")
    ap.add_argument("--partner-path", help="Pakai file lokal (csv/xlsx) untuk Affiliate_Patner, bukan dari Drive")
    ap.add_argument("--product-path", help="Pakai file Product Performance lokal (xlsx, pisah koma), bukan dari Drive")
    ap.add_argument("--product-bank-path", help="Pakai file Product Bank lokal (xlsx/csv), bukan dari Drive")
    ap.add_argument("--no-update-existing", action="store_true",
                    help="Order: baris yang sudah ada di sheet dibiarkan apa adanya (status/komisi tidak diperbarui)")
    ap.add_argument("--dry-run", action="store_true", help="Tidak menulis apa pun, hanya laporan + preview")
    ap.add_argument("--yes", action="store_true", help="Tanpa tanya konfirmasi sebelum menulis")
    args = ap.parse_args()

    cfg = {}
    if os.path.exists(args.config):
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
    cfg = deep_cfg(cfg)
    base = os.path.dirname(os.path.abspath(args.config))
    out_dir = os.path.join(base, cfg.get("report_dir", "laporan"))
    os.makedirs(out_dir, exist_ok=True)
    labels = [cfg["source_label_order"], cfg["source_label_partner"]]

    mode = args.mode
    if not mode:
        if args.order or args.partner or args.order_path or args.partner_path:
            mode = "orders"
        elif args.product or args.product_path:
            mode = "product"
        else:
            pick = input("Mau isi apa?  1. Order (Affiliate + Patner)   2. Product Performance Daily (Raw)   3. Keduanya  : ").strip()
            mode = {"1": "orders", "2": "product", "3": "all"}.get(pick, "orders")

    local_only = args.dry_run and args.product_bank_path and (args.order_path or args.partner_path or args.product_path)
    g = None
    if not local_only:
        cred_file = cfg.get("credentials_file") or cfg.get("service_account_file") or "credentials.json"
        if not os.path.isabs(cred_file):
            cred_file = os.path.join(base, cred_file)
        if not os.path.exists(cred_file):  # sama seperti nivea_tracker.py: coba nama lama
            for alt in ("credentials.json", "service_account.json"):
                if os.path.exists(os.path.join(base, alt)):
                    cred_file = os.path.join(base, alt)
                    break
        g = Google(cred_file, os.path.join(base, cfg.get("token_file", "token.json")))

    if args.product_bank_path:
        with open(args.product_bank_path, "rb") as f:
            bank = bank_from_frame(read_source_bytes(f.read(), args.product_bank_path, "product id"))
    else:
        bank = read_product_bank(g, cfg["merge_product_bank_file_id"], cfg["merge_product_bank_sheet"])
    print(f"Product Bank: {fmt_n(len(bank))} Product ID punya CAT/PRO")

    if mode in ("orders", "all"):
        print("\n========== ORDER (Affiliate + Patner) ==========")
        run_orders(g, cfg, args, bank, out_dir, labels)
    if mode in ("product", "all"):
        print("\n========== PRODUCT PERFORMANCE DAILY (RAW) ==========")
        run_product(g, cfg, args, bank, out_dir)


if __name__ == "__main__":
    main()
