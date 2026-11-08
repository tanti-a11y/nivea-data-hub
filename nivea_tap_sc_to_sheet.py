#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gabungkan TAP (xlsx) + SC (csv) Nivea  ->  tulis ke Google Sheet "Raw Data TAP&SC"

Alur: BACA file -> HITUNG -> VALIDASI angka -> (opsional) TULIS ke Google Sheet.
Default = DRY-RUN (hanya baca + validasi + simpan preview). Tambah --write untuk benar-benar menulis.

Contoh:
  # 1) cek dulu tanpa menyentuh Google Sheet
  python nivea_tap_sc_to_sheet.py --tap TAP_GMV_Nivea.xlsx --sc affiliate_orders.csv --offline

  # 2) cek + bandingkan dengan isi sheet (butuh login OAuth), belum menulis
  python nivea_tap_sc_to_sheet.py --tap TAP_GMV_Nivea.xlsx --sc affiliate_orders.csv

  # 3) tulis ke sheet (append di bawah data terakhir), hanya jika semua validasi lolos
  python nivea_tap_sc_to_sheet.py --tap TAP_GMV_Nivea.xlsx --sc affiliate_orders.csv --write

Kolom sheet (A..T):
  A Jenis Creator(rumus) | B Creator Username | C Cek | D Product Name | E Cek | F CAT(rumus) | G PRO(rumus)
  H Product ID | I Date | J WEEK(rumus) | K GMV LS | L QYT LS | M GMV SL | N QYT SL | O GMV SV | P QYT SV
  Q Total GMV(rumus) | R LS+SV(rumus) | S Kuantitas Sold(rumus) | T SOURCE (TAP / SC)
Script hanya mengisi nilai B, D, H, I, K-P, T. Kolom lain diisi RUMUS yang sama persis dengan baris lama.

Aturan hitung:
  TAP  GMV LS = Creator LIVE-attributed GMV
       GMV SV = Affiliate video-attributed GMV
       GMV SL = Revenue (Showcase) + Link GMV
       avg sold = Creator-attributed items sold / Creator-attributed orders
       QYT LS = avg * Creator LIVE-attributed orders
       QYT SV = avg * Creator video-attributed orders
       QYT SL = Creator-attributed items sold - (QYT LS + QYT SV)
       (hasil pecahan dibulatkan dengan metode "largest remainder" supaya jumlah LS+SV+SL
        tetap persis sama dengan items sold)
  SC   GMV = Est. Commission Base, dipisah per Content Type:
       Livestream -> LS | Video -> SV | selain itu (Showcase, External Traffic Program, dst) -> SL
       QYT = Quantity
       Tanggal = Time Created (dd/mm/yyyy)

Instalasi:  pip install pandas openpyxl google-api-python-client google-auth-oauthlib
Login    :  taruh credentials.json (OAuth Desktop client) di folder yang sama; token.json dibuat otomatis.
            (Service Account diblokir org policy, jadi pakai OAuth Desktop.)
"""
import subprocess
import sys


def _ensure(pkgs):
    """Pasang otomatis library yang belum ada (satu file, tanpa requirements.txt)."""
    import importlib
    miss = [p for mod, p in pkgs if importlib.util.find_spec(mod) is None]
    if miss:
        raise RuntimeError("Library belum terpasang: " + ", ".join(miss))


import importlib.util  # noqa: E402
_ensure([("pandas", "pandas"), ("openpyxl", "openpyxl")])

CLIENT_CONFIG = {}  # dikosongkan: kredensial TIDAK boleh ditanam di kode (diambil dari Streamlit secrets oleh hub/auth.py)

import argparse
import math
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ----------------------------------------------------------------------------------------
# KONFIGURASI
# ----------------------------------------------------------------------------------------
SPREADSHEET_ID = "1rpOcHUqiN16ZVdFydtDLbJH50z3sWP7L5ml1kl2sJxA"
SHEET_NAME = "Raw Data TAP&SC"
HEADER_ROW = 15
FIRST_DATA_ROW = 16
LOOKUP_CREATOR_SHEET = "Sheet5"   # A: creator, B: kategori (dipakai rumus kolom A)
LOOKUP_PRODUCT_SHEET = "Sheet9"   # A: product id, C: CAT, D: PRO (dipakai rumus kolom F & G)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CHUNK_ROWS = 4000

EXPECTED_HEADER = ["Jenis Creator", "Creator Username", "Cek", "Product Name", "Cek", "CAT", "PRO",
                   "Product ID", "Date", "WEEK", "GMV LS", "QYT LS", "GMV SL", "QYT SL", "GMV SV",
                   "QYT SV", "Total GMV", "LS+SV", "Kuantitas Sold", "SOURCE"]

TAP_COLS = ["Date", "Creator name", "Product ID", "Product name", "Creator-attributed GMV",
            "Affiliate video-attributed GMV", "Creator LIVE-attributed GMV", "Revenue (Showcase)",
            "Link GMV", "Creator-attributed orders", "Creator LIVE-attributed orders",
            "Creator video-attributed orders", "Creator-attributed items sold"]
SC_COLS = ["Order ID", "Product ID", "Product Name", "SKU ID", "Quantity", "Order Status", "Currency",
           "Creator Username", "Content Type", "Est. Commission Base", "Time Created"]
SC_LS_TYPES = {"livestream"}
SC_SV_TYPES = {"video"}
SC_KNOWN_TYPES = {"livestream", "video", "showcase", "external traffic program"}

WEEK_FORMULA = (
    '=UPPER(TEXT(I{r},"MMM"))&" W"&\nIF(DAY(I{r})<=6,1,\nIF(DAY(I{r})<=13,2,\nIF(DAY(I{r})<=20,3,\n'
    'IF(DAY(I{r})<=27,4,5))))&\n" ("&\nIF(DAY(I{r})<=6,1,\nIF(DAY(I{r})<=13,7,\nIF(DAY(I{r})<=20,14,\n'
    'IF(DAY(I{r})<=27,21,28))))&\n"-"&\nIF(DAY(I{r})<=6,6,\nIF(DAY(I{r})<=13,13,\nIF(DAY(I{r})<=20,20,\n'
    'IF(DAY(I{r})<=27,27,DAY(EOMONTH(I{r},0))))))&\n")"'
)


# ----------------------------------------------------------------------------------------
# Laporan validasi
# ----------------------------------------------------------------------------------------
class Report:
    def __init__(self):
        self.lines, self.errors, self.warnings = [], [], []

    def info(self, msg=""):
        self.lines.append(msg)
        print(msg)

    def ok(self, msg):
        self.info(f"  [OK]   {msg}")

    def warn(self, msg):
        self.warnings.append(msg)
        self.info(f"  [WARN] {msg}")

    def err(self, msg):
        self.errors.append(msg)
        self.info(f"  [FAIL] {msg}")

    def check(self, cond, ok_msg, fail_msg):
        (self.ok if cond else self.err)(ok_msg if cond else fail_msg)
        return cond


# ----------------------------------------------------------------------------------------
# Helper parsing
# ----------------------------------------------------------------------------------------
_RP = re.compile(r"^-?\s*Rp\s*[\d.,]+$")


def parse_rp(series, name, rep):
    """'Rp6.780.001' -> 6780001 (titik = ribuan, koma = desimal). Gagal keras jika ada format aneh."""
    s = series.astype(str).str.strip()
    bad = ~s.str.match(_RP)
    if bad.any():
        rep.err(f"{name}: {int(bad.sum())} sel bukan format 'Rp...' (contoh: {s[bad].head(3).tolist()})")
        return None
    neg = s.str.startswith("-")
    num = (s.str.replace("-", "", regex=False).str.replace("Rp", "", regex=False).str.strip()
           .str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float))
    return num.where(~neg, -num)


def parse_int(series, name, rep):
    try:
        out = pd.to_numeric(series.astype(str).str.replace(",", "", regex=False).str.strip(), errors="raise")
    except Exception as e:  # noqa
        rep.err(f"{name}: ada nilai non-angka ({e})")
        return None
    return out.astype(float)


def clean_str(series):
    return series.fillna("").astype(str).str.strip().replace({"nan": "", "--": ""})


def fmt_idr(x):
    return f"{x:,.0f}".replace(",", ".")


def qty_split(items, orders, ls_orders, sv_orders):
    """Pecah items sold -> (LS, SV, SL) integer, jumlah = items (largest remainder)."""
    items = int(round(items))
    if orders <= 0 or items <= 0:
        return 0, 0, max(items, 0), False
    avg = items / orders
    raw = [round(avg * ls_orders, 6), round(avg * sv_orders, 6)]
    raw.append(round(max(items - raw[0] - raw[1], 0.0), 6))
    base = [int(math.floor(v)) for v in raw]
    need = items - sum(base)
    fractional = any(abs(v - round(v)) > 1e-9 for v in raw)
    if need > 0:
        order = sorted(range(3), key=lambda i: (-(raw[i] - base[i]), i))
        for i in order[:need]:
            base[i] += 1
    return base[0], base[1], base[2], fractional


# ----------------------------------------------------------------------------------------
# TAP
# ----------------------------------------------------------------------------------------
def load_tap(path, rep):
    rep.info(f"\n=== TAP: {Path(path).name} ===")
    raw = pd.read_excel(path, dtype=str)
    missing = [c for c in TAP_COLS if c not in raw.columns]
    if missing:
        rep.err(f"Kolom TAP tidak ditemukan: {missing}")
        return None, None
    df = raw[TAP_COLS].copy()
    rep.info(f"  baris dibaca: {len(df):,}")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    if not rep.check(df["Date"].notna().all(), "semua tanggal valid", "ada tanggal TAP yang tidak valid"):
        return None, None
    for c in ["Creator-attributed GMV", "Affiliate video-attributed GMV", "Creator LIVE-attributed GMV",
              "Revenue (Showcase)", "Link GMV"]:
        df[c] = parse_rp(df[c], c, rep)
    for c in ["Creator-attributed orders", "Creator LIVE-attributed orders", "Creator video-attributed orders",
              "Creator-attributed items sold"]:
        df[c] = parse_int(df[c], c, rep)
    if rep.errors:
        return None, None
    df["Creator name"] = clean_str(df["Creator name"])
    df["Product ID"] = clean_str(df["Product ID"])
    df["Product name"] = clean_str(df["Product name"])
    rep.check((df["Creator name"] != "").all(), "creator tidak ada yang kosong", "ada creator kosong di TAP")
    rep.check((df["Product ID"] != "").all(), "product id tidak ada yang kosong", "ada product id kosong di TAP")

    # hanya ambil baris dengan Creator-attributed GMV bukan 0
    zero = df["Creator-attributed GMV"] == 0
    if zero.any():
        rep.info(f"  baris Creator-attributed GMV = 0 dibuang: {int(zero.sum()):,} "
                 f"(items sold di baris itu: {fmt_idr(df.loc[zero, 'Creator-attributed items sold'].sum())})")
        df = df[~zero].copy()
    else:
        rep.info("  tidak ada baris dengan Creator-attributed GMV = 0")
    rep.info(f"  baris dipakai: {len(df):,}")

    df["gmv_ls"] = df["Creator LIVE-attributed GMV"]
    df["gmv_sv"] = df["Affiliate video-attributed GMV"]
    df["gmv_sl"] = df["Revenue (Showcase)"] + df["Link GMV"]
    mism = (df["Creator-attributed GMV"] - (df.gmv_ls + df.gmv_sv + df.gmv_sl)).abs() > 0.5
    rep.check(not mism.any(), "Creator-attributed GMV == LS + SV + SL di setiap baris (rumus GMV konsisten)",
              f"{int(mism.sum())} baris: Creator-attributed GMV != LS+SV+SL (cek definisi kolom)")
    rep.check((df[["gmv_ls", "gmv_sv", "gmv_sl"]] >= 0).all().all(), "tidak ada GMV negatif", "ada GMV negatif")

    ls_q, sv_q, sl_q, frac = [], [], [], []
    for it, od, lo, vo in zip(df["Creator-attributed items sold"], df["Creator-attributed orders"],
                              df["Creator LIVE-attributed orders"], df["Creator video-attributed orders"]):
        a, b, c, f = qty_split(it, od, lo, vo)
        ls_q.append(a); sv_q.append(b); sl_q.append(c); frac.append(f)
    df["qty_ls"], df["qty_sv"], df["qty_sl"], df["_frac"] = ls_q, sv_q, sl_q, frac
    rep.check((df.qty_ls + df.qty_sv + df.qty_sl == df["Creator-attributed items sold"]).all(),
              "QYT LS+SV+SL == Creator-attributed items sold di semua baris",
              "ada baris yang QYT LS+SV+SL != items sold")
    if sum(frac):
        rep.warn(f"{int(sum(frac))} baris punya pecahan saat avg.sold x orders (dibulatkan largest-remainder, total tetap sama)")
    zero_ord = (df["Creator-attributed orders"] <= 0) & (df["Creator-attributed items sold"] > 0)
    if zero_ord.any():
        rep.warn(f"{int(zero_ord.sum())} baris orders=0 tapi items sold>0 (avg.sold tak terdefinisi, semua dimasukkan ke SL)")
    over = (df["Creator LIVE-attributed orders"] + df["Creator video-attributed orders"]) > df["Creator-attributed orders"]
    if over.any():
        rep.warn(f"{int(over.sum())} baris LIVE orders + video orders > total orders")

    # agregasi per creator-product-date (campaign berbeda -> dijumlah)
    key = ["Creator name", "Product ID", "Date"]
    dup = int(df.duplicated(key).sum())
    if dup:
        rep.warn(f"{dup} baris TAP punya creator+product+tanggal yang sama (beda campaign) -> dijumlahkan jadi 1 baris")
    agg = df.groupby(key, as_index=False).agg(
        name=("Product name", "first"), gmv_ls=("gmv_ls", "sum"), gmv_sl=("gmv_sl", "sum"), gmv_sv=("gmv_sv", "sum"),
        qty_ls=("qty_ls", "sum"), qty_sl=("qty_sl", "sum"), qty_sv=("qty_sv", "sum"))
    out = pd.DataFrame({
        "creator": agg["Creator name"], "product_name": agg["name"], "product_id": agg["Product ID"],
        "date": agg["Date"].dt.strftime("%Y-%m-%d"),
        "gmv_ls": agg.gmv_ls, "qty_ls": agg.qty_ls, "gmv_sl": agg.gmv_sl, "qty_sl": agg.qty_sl,
        "gmv_sv": agg.gmv_sv, "qty_sv": agg.qty_sv, "source": "TAP"})

    # rekonsiliasi output vs sumber (jalur hitung terpisah)
    rep.info("  -- rekonsiliasi TAP (output vs file mentah) --")
    pairs = [("GMV LS", out.gmv_ls.sum(), df["Creator LIVE-attributed GMV"].sum()),
             ("GMV SV", out.gmv_sv.sum(), df["Affiliate video-attributed GMV"].sum()),
             ("GMV SL", out.gmv_sl.sum(), (df["Revenue (Showcase)"] + df["Link GMV"]).sum()),
             ("Total GMV", (out.gmv_ls + out.gmv_sl + out.gmv_sv).sum(), df["Creator-attributed GMV"].sum()),
             ("Total QYT", (out.qty_ls + out.qty_sl + out.qty_sv).sum(), df["Creator-attributed items sold"].sum())]
    for n, a, b in pairs:
        rep.check(abs(a - b) < 0.5, f"{n}: output {fmt_idr(a)} == sumber {fmt_idr(b)}",
                  f"{n}: output {fmt_idr(a)} != sumber {fmt_idr(b)}")
    per_day_out = (out.assign(t=out.gmv_ls + out.gmv_sl + out.gmv_sv).groupby("date")["t"].sum())
    per_day_raw = df.groupby(df["Date"].dt.strftime("%Y-%m-%d"))["Creator-attributed GMV"].sum()
    rep.check(((per_day_out - per_day_raw).abs() < 0.5).all(), "Total GMV per tanggal cocok dengan sumber",
              "Total GMV per tanggal TAP tidak cocok")
    rep.info(f"  baris output: {len(out):,} | tanggal {out.date.min()} s/d {out.date.max()} | creator unik {out.creator.nunique():,}")
    rep.info("  GMV per tanggal: " + ", ".join(f"{d[5:]}={fmt_idr(v)}" for d, v in per_day_out.items()))
    return out, {"gmv": out[["gmv_ls", "gmv_sl", "gmv_sv"]].sum().sum(), "qty": out[["qty_ls", "qty_sl", "qty_sv"]].sum().sum()}


# ----------------------------------------------------------------------------------------
# SC
# ----------------------------------------------------------------------------------------
def load_sc(path, rep, exclude_status=()):
    rep.info(f"\n=== SC: {Path(path).name} ===")
    raw = pd.read_csv(path, dtype=str, low_memory=False)
    missing = [c for c in SC_COLS if c not in raw.columns]
    if missing:
        rep.err(f"Kolom SC tidak ditemukan: {missing}")
        return None, None
    df = raw[SC_COLS].copy()
    rep.info(f"  baris dibaca: {len(df):,}")
    base = pd.to_numeric(df["Est. Commission Base"], errors="coerce")
    qty = pd.to_numeric(df["Quantity"], errors="coerce")
    rep.check(base.notna().all(), "Est. Commission Base semua angka", f"{int(base.isna().sum())} baris Est. Commission Base bukan angka")
    rep.check(qty.notna().all(), "Quantity semua angka", f"{int(qty.isna().sum())} baris Quantity bukan angka")
    dts = pd.to_datetime(df["Time Created"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    rep.check(dts.notna().all(), "Time Created semua terbaca (dd/mm/yyyy hh:mm:ss)", f"{int(dts.isna().sum())} baris Time Created tidak terbaca")
    if rep.errors:
        return None, None
    df["base"], df["qty"], df["date"] = base, qty, dts.dt.strftime("%Y-%m-%d")
    df["Creator Username"] = clean_str(df["Creator Username"])
    df["Product ID"] = clean_str(df["Product ID"])
    df["Product Name"] = clean_str(df["Product Name"])
    rep.check((df["Currency"].str.upper() == "IDR").all(), "semua mata uang IDR", "ada currency selain IDR")
    rep.check((df["Product ID"] != "").all(), "product id tidak ada yang kosong", "ada product id kosong di SC")
    dup = int(df.duplicated(["Order ID", "SKU ID"]).sum())
    rep.check(dup == 0, "tidak ada duplikat Order ID + SKU ID", f"{dup} baris duplikat Order ID + SKU ID (file ter-export dobel?)")
    blank = int((df["Creator Username"] == "").sum())
    if blank:
        rep.warn(f"{blank} baris tanpa Creator Username -> ditulis dengan creator kosong (sama seperti baris COMMUNITY lama)")
    ct = df["Content Type"].fillna("").str.strip().str.lower()
    unknown = sorted(set(ct) - SC_KNOWN_TYPES)
    if unknown:
        rep.warn(f"Content Type baru terdeteksi: {unknown} -> dihitung sebagai SL")
    if (qty < 0).any() or (base < 0).any():
        rep.warn("ada Quantity / Est. Commission Base negatif")
    rep.info("  Order Status (Est. Commission Base): " + ", ".join(
        f"{k}={fmt_idr(v)}" for k, v in df.groupby("Order Status")["base"].sum().items()))
    if exclude_status:
        n0 = len(df)
        df = df[~df["Order Status"].isin(exclude_status)]
        rep.warn(f"status {list(exclude_status)} dibuang: {n0 - len(df):,} baris")
    else:
        rep.info("  (semua Order Status ikut dihitung; pakai --sc-exclude-status 'Ineligible' dst. jika perlu dibuang)")

    ct = df["Content Type"].fillna("").str.strip().str.lower()
    df["bucket"] = ct.map(lambda x: "LS" if x in SC_LS_TYPES else ("SV" if x in SC_SV_TYPES else "SL"))
    key = ["Creator Username", "Product ID", "date"]
    g = df.groupby(key + ["bucket"], as_index=False).agg(base=("base", "sum"), qty=("qty", "sum"))
    names = df.groupby("Product ID")["Product Name"].first()
    wide = g.pivot_table(index=key, columns="bucket", values=["base", "qty"], aggfunc="sum").reset_index()
    wide.columns = [a if not b else f"{a}_{b}" for a, b in wide.columns]
    for c in ["base_LS", "base_SV", "base_SL", "qty_LS", "qty_SV", "qty_SL"]:
        if c not in wide:
            wide[c] = float("nan")
    out = pd.DataFrame({
        "creator": wide["Creator Username"], "product_name": wide["Product ID"].map(names), "product_id": wide["Product ID"],
        "date": wide["date"], "gmv_ls": wide.base_LS, "qty_ls": wide.qty_LS, "gmv_sl": wide.base_SL, "qty_sl": wide.qty_SL,
        "gmv_sv": wide.base_SV, "qty_sv": wide.qty_SV, "source": "SC"})

    rep.info("  -- rekonsiliasi SC (output vs file mentah) --")
    for n, a, b in [("GMV LS", out.gmv_ls.sum(), df.loc[df.bucket == "LS", "base"].sum()),
                    ("GMV SV", out.gmv_sv.sum(), df.loc[df.bucket == "SV", "base"].sum()),
                    ("GMV SL", out.gmv_sl.sum(), df.loc[df.bucket == "SL", "base"].sum()),
                    ("Total GMV", out[["gmv_ls", "gmv_sl", "gmv_sv"]].sum().sum(), df["base"].sum()),
                    ("Total QYT", out[["qty_ls", "qty_sl", "qty_sv"]].sum().sum(), df["qty"].sum())]:
        rep.check(abs(a - b) < 0.5, f"{n}: output {fmt_idr(a)} == sumber {fmt_idr(b)}",
                  f"{n}: output {fmt_idr(a)} != sumber {fmt_idr(b)}")
    tot_out = out.assign(t=out[["gmv_ls", "gmv_sl", "gmv_sv"]].sum(axis=1)).groupby("date")["t"].sum()
    tot_raw = df.groupby("date")["base"].sum()
    rep.check(((tot_out - tot_raw).abs() < 0.5).all(), "Total GMV per tanggal cocok dengan sumber", "Total GMV per tanggal SC tidak cocok")
    rep.info(f"  baris output: {len(out):,} | tanggal {out.date.min()} s/d {out.date.max()} | creator unik {out.creator.nunique():,}")
    rep.info("  GMV per tanggal: " + ", ".join(f"{d[5:]}={fmt_idr(v)}" for d, v in tot_out.items()))
    return out, {"gmv": out[["gmv_ls", "gmv_sl", "gmv_sv"]].sum().sum(), "qty": out[["qty_ls", "qty_sl", "qty_sv"]].sum().sum()}


# ----------------------------------------------------------------------------------------
# Build baris sheet
# ----------------------------------------------------------------------------------------
def as_text(v):
    """Cegah USER_ENTERED mengubah teks jadi angka/rumus (mis. product id 19 digit, username '=...')."""
    v = "" if v is None else str(v)
    if not v:
        return ""
    try:
        float(v)
        return "'" + v
    except ValueError:
        return "'" + v if v[0] in "=+-@'" else v


def num_or_blank(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    return int(v) if float(v).is_integer() else float(v)


def build_rows(df, start_row):
    rows = []
    for i, r in enumerate(df.itertuples(index=False)):
        n = start_row + i
        rows.append([
            f'=XLOOKUP(B{n},{LOOKUP_CREATOR_SHEET}!$A:$A,{LOOKUP_CREATOR_SHEET}!$B:$B,"OPEN PLAN")',
            as_text(r.creator), "", as_text(r.product_name), "",
            f"=XLOOKUP(H{n},{LOOKUP_PRODUCT_SHEET}!$A:$A,{LOOKUP_PRODUCT_SHEET}!$C:$C)",
            f"=XLOOKUP(H{n},{LOOKUP_PRODUCT_SHEET}!$A:$A,{LOOKUP_PRODUCT_SHEET}!$D:$D)",
            "'" + str(r.product_id), r.date, WEEK_FORMULA.format(r=n),
            num_or_blank(r.gmv_ls), num_or_blank(r.qty_ls), num_or_blank(r.gmv_sl), num_or_blank(r.qty_sl),
            num_or_blank(r.gmv_sv), num_or_blank(r.qty_sv),
            f"=SUM(K{n},M{n},O{n})", f"=SUM(K{n},O{n})", f"=SUM(L{n},N{n},P{n})", r.source])
    return rows


# ----------------------------------------------------------------------------------------
# Google Sheets
# ----------------------------------------------------------------------------------------
def get_service(creds_file="credentials.json", token_file="token.json"):
    _ensure([("googleapiclient", "google-api-python-client"), ("google_auth_oauthlib", "google-auth-oauthlib")])
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    creds = None
    if Path(token_file).exists():
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if Path(creds_file).exists():
                flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
            else:  # pakai credentials yang ditanam di script
                flow = InstalledAppFlow.from_client_config(CLIENT_CONFIG, SCOPES)
            creds = flow.run_local_server(port=0)
        Path(token_file).write_text(creds.to_json())
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def q(sheet):
    return "'" + sheet.replace("'", "''") + "'"


def read_sheet_state(svc, rep):
    vals = svc.spreadsheets().values()
    hdr = vals.get(spreadsheetId=SPREADSHEET_ID, range=f"{q(SHEET_NAME)}!A{HEADER_ROW}:T{HEADER_ROW}").execute().get("values", [[]])[0]
    rep.check(hdr == EXPECTED_HEADER, "header sheet (baris 15) sesuai struktur", f"header sheet beda dari yang diharapkan: {hdr}")
    i_col = vals.get(spreadsheetId=SPREADSHEET_ID, range=f"{q(SHEET_NAME)}!I{FIRST_DATA_ROW}:I",
                     valueRenderOption="FORMATTED_VALUE").execute().get("values", [])
    t_col = vals.get(spreadsheetId=SPREADSHEET_ID, range=f"{q(SHEET_NAME)}!T{FIRST_DATA_ROW}:T").execute().get("values", [])
    n = max(len(i_col), len(t_col))
    last = FIRST_DATA_ROW + n - 1
    existing = set()
    for k in range(n):
        d = i_col[k][0] if k < len(i_col) and i_col[k] else ""
        s = t_col[k][0] if k < len(t_col) and t_col[k] else ""
        if d and s:
            existing.add((d, s))
    rep.info(f"  sheet saat ini: baris data {FIRST_DATA_ROW}-{last} ({n:,} baris); kombinasi tanggal+source: {len(existing)}")
    prod = vals.get(spreadsheetId=SPREADSHEET_ID, range=f"{q(LOOKUP_PRODUCT_SHEET)}!A:A").execute().get("values", [])
    cre = vals.get(spreadsheetId=SPREADSHEET_ID, range=f"{q(LOOKUP_CREATOR_SHEET)}!A:A").execute().get("values", [])
    return {"last_row": last, "existing": existing,
            "products": {r[0].strip() for r in prod if r}, "creators": {r[0].strip() for r in cre if r}}


def check_against_sheet(df, state, rep, allow_overlap):
    rep.info("\n=== Cek terhadap isi Google Sheet ===")
    new_pairs = set(zip(df.date, df.source))
    overlap = sorted(new_pairs & state["existing"])
    if overlap:
        msg = f"{len(overlap)} kombinasi tanggal+source SUDAH ADA di sheet (mis. {overlap[:3]}) -> risiko data dobel"
        rep.warn(msg + " (dilanjutkan karena --allow-overlap)") if allow_overlap else rep.err(msg + ". Hapus baris lama dulu atau pakai --allow-overlap.")
    else:
        rep.ok(f"tidak ada tumpang tindih tanggal+source dengan data yang sudah ada ({len(new_pairs)} kombinasi baru)")
    miss = df[~df.product_id.isin(state["products"])]
    if len(miss):
        gm = miss[["gmv_ls", "gmv_sl", "gmv_sv"]].sum().sum()
        ex = miss.groupby("product_id")[["gmv_ls", "gmv_sl", "gmv_sv"]].sum().sum(axis=1).sort_values(ascending=False).head(5)
        rep.warn(f"{miss.product_id.nunique()} product ID belum ada di {LOOKUP_PRODUCT_SHEET} (CAT/PRO jadi #N/A), "
                 f"GMV terdampak {fmt_idr(gm)}. Teratas: {ex.round(0).astype(int).to_dict()}")
    else:
        rep.ok(f"semua product ID ada di {LOOKUP_PRODUCT_SHEET} (CAT & PRO akan terisi)")
    new_cre = df[(df.creator != "") & ~df.creator.isin(state["creators"])]
    if len(new_cre):
        rep.info(f"  info: {new_cre.creator.nunique():,} creator tidak ada di {LOOKUP_CREATOR_SHEET} -> Jenis Creator = 'OPEN PLAN' "
                 f"(baris TAP di antaranya: {int((new_cre.source == 'TAP').sum())})")
        tap_new = new_cre[new_cre.source == "TAP"].creator.unique()[:10]
        if len(tap_new):
            rep.warn(f"creator TAP yang belum terdaftar di {LOOKUP_CREATOR_SHEET} (jadi OPEN PLAN): {list(tap_new)}")


def write_to_sheet(svc, df, state, rep):
    rep.info("\n=== Menulis ke Google Sheet ===")
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields="sheets.properties").execute()
    props = next(s["properties"] for s in meta["sheets"] if s["properties"]["title"] == SHEET_NAME)
    sheet_id, grid_rows = props["sheetId"], props["gridProperties"]["rowCount"]
    start = state["last_row"] + 1
    end = start + len(df) - 1
    reqs = []
    if end > grid_rows:
        reqs.append({"appendDimension": {"sheetId": sheet_id, "dimension": "ROWS", "length": end - grid_rows + 100}})
    reqs.append({"copyPaste": {
        "source": {"sheetId": sheet_id, "startRowIndex": state["last_row"] - 1, "endRowIndex": state["last_row"],
                   "startColumnIndex": 0, "endColumnIndex": 20},
        "destination": {"sheetId": sheet_id, "startRowIndex": start - 1, "endRowIndex": end,
                        "startColumnIndex": 0, "endColumnIndex": 20},
        "pasteType": "PASTE_FORMAT"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": reqs}).execute()
    rows = build_rows(df, start)
    for i in range(0, len(rows), CHUNK_ROWS):
        part = rows[i:i + CHUNK_ROWS]
        a, b = start + i, start + i + len(part) - 1
        svc.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=f"{q(SHEET_NAME)}!A{a}:T{b}",
            valueInputOption="USER_ENTERED", body={"values": part}).execute()
        rep.info(f"  tertulis baris {a}-{b}")
    return start, end


def verify_written(svc, df, start, end, rep):
    """Baca balik hasil rumus Sheets (Q, R, S) dan cocokkan dengan angka dari file."""
    rep.info("\n=== Verifikasi setelah tulis ===")
    got = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"{q(SHEET_NAME)}!A{start}:T{end}",
        valueRenderOption="UNFORMATTED_VALUE").execute().get("values", [])
    rep.check(len(got) == len(df), f"jumlah baris terbaca {len(got):,} == {len(df):,}", f"baris terbaca {len(got):,} != {len(df):,}")
    g = pd.DataFrame([r + [""] * (20 - len(r)) for r in got])
    for col, name, exp in [(16, "Total GMV (kolom Q)", (df.gmv_ls.fillna(0) + df.gmv_sl.fillna(0) + df.gmv_sv.fillna(0))),
                           (17, "LS+SV (kolom R)", (df.gmv_ls.fillna(0) + df.gmv_sv.fillna(0))),
                           (18, "Kuantitas (kolom S)", (df.qty_ls.fillna(0) + df.qty_sl.fillna(0) + df.qty_sv.fillna(0)))]:
        s = pd.to_numeric(g[col], errors="coerce")
        rep.check(s.notna().all() and abs(s.sum() - exp.sum()) < 0.5,
                  f"{name}: sheet {fmt_idr(s.sum())} == file {fmt_idr(exp.sum())}",
                  f"{name}: sheet {fmt_idr(s.sum())} != file {fmt_idr(exp.sum())} (atau ada sel error)")
    err_cells = g[[0, 5, 6, 9]].astype(str).apply(lambda c: c.str.startswith("#")).sum()
    for col, nm in zip([0, 5, 6, 9], ["Jenis Creator", "CAT", "PRO", "WEEK"]):
        if err_cells[col]:
            rep.warn(f"kolom {nm}: {int(err_cells[col])} sel error (#N/A dsb.)")


# ----------------------------------------------------------------------------------------
# Pilih file lewat jendela (tkinter)
# ----------------------------------------------------------------------------------------
def pick_files(tap, sc):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    if not tap:
        tap = filedialog.askopenfilename(title="Pilih file TAP (.xlsx)", filetypes=[("Excel", "*.xlsx *.xls"), ("Semua", "*.*")])
    if tap and not sc:
        sc = filedialog.askopenfilename(title="Pilih file SC affiliate_orders (.csv)",
                                        initialdir=str(Path(tap).parent), filetypes=[("CSV", "*.csv"), ("Semua", "*.*")])
    root.destroy()
    return tap, sc


# ----------------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Gabung TAP + SC Nivea -> Google Sheet 'Raw Data TAP&SC'")
    ap.add_argument("--tap", help="file TAP .xlsx (kosong = pilih lewat jendela)")
    ap.add_argument("--sc", help="file SC affiliate_orders .csv (kosong = pilih lewat jendela)")
    ap.add_argument("--write", action="store_true", help="benar-benar tulis ke Google Sheet (default: dry-run)")
    ap.add_argument("--offline", action="store_true", help="jangan konek ke Google (validasi file saja)")
    ap.add_argument("--allow-overlap", action="store_true", help="izinkan tulis walau tanggal+source sudah ada di sheet")
    ap.add_argument("--sc-exclude-status", nargs="*", default=[], help="Order Status SC yang dibuang, mis. Ineligible")
    ap.add_argument("--creds", default="credentials.json")
    ap.add_argument("--token", default="token.json")
    ap.add_argument("--outdir", default=".", help="folder untuk preview CSV + laporan validasi")
    a = ap.parse_args()

    interactive = not (a.tap and a.sc)
    if interactive:
        a.tap, a.sc = pick_files(a.tap, a.sc)
        if not a.tap or not a.sc:
            print("Pemilihan file dibatalkan.")
            return 1

    rep = Report()
    rep.info(f"Nivea TAP + SC  |  {datetime.now():%Y-%m-%d %H:%M}")
    tap, _ = load_tap(a.tap, rep)
    sc, _ = load_sc(a.sc, rep, tuple(a.sc_exclude_status))
    if tap is None or sc is None:
        rep.err("file gagal dibaca/divalidasi -> berhenti")
        return finish(rep, a, None, 1)

    rep.info("\n=== Gabungan ===")
    rep.info(f"  TAP {tap.date.min()}..{tap.date.max()} | SC {sc.date.min()}..{sc.date.max()}")
    if (tap.date.min(), tap.date.max()) != (sc.date.min(), sc.date.max()):
        rep.warn("rentang tanggal TAP dan SC berbeda")
    df = pd.concat([tap, sc], ignore_index=True)
    df = df.sort_values(["source", "date", "creator", "product_id"], ascending=[False, True, True, True]).reset_index(drop=True)
    tot = df[["gmv_ls", "gmv_sl", "gmv_sv"]].sum()
    rep.info(f"  total baris: {len(df):,} (TAP {len(tap):,} + SC {len(sc):,})")
    rep.info(f"  GMV LS {fmt_idr(tot.gmv_ls)} | SL {fmt_idr(tot.gmv_sl)} | SV {fmt_idr(tot.gmv_sv)} | Total {fmt_idr(tot.sum())}")
    rep.check(len(df) == len(tap) + len(sc), "jumlah baris gabungan benar", "jumlah baris gabungan tidak cocok")

    state = None
    if not a.offline:
        try:
            svc = get_service(a.creds, a.token)
            state = read_sheet_state(svc, rep)
            check_against_sheet(df, state, rep, a.allow_overlap)
        except Exception as e:  # noqa
            rep.err(f"gagal akses Google Sheet: {e}")
    else:
        rep.info("  (mode --offline: pengecekan terhadap isi sheet dilewati)")

    if interactive and not a.write and not a.offline and state is not None:
        if rep.errors:
            print("\nAda [FAIL] -> tidak bisa menulis. Perbaiki dulu.")
        else:
            ans = input(f"\nTulis {len(df):,} baris ke Google Sheet sekarang? ketik 'y' lalu Enter (selain itu = batal): ")
            a.write = ans.strip().lower() in ("y", "ya", "yes")

    if a.write:
        if a.offline or state is None:
            rep.err("--write butuh koneksi ke Google Sheet (jangan pakai --offline)")
        elif rep.errors:
            rep.err("ADA validasi gagal -> TIDAK menulis ke sheet")
        else:
            start, end = write_to_sheet(svc, df, state, rep)
            verify_written(svc, df, start, end, rep)
    else:
        rep.info("\nDRY-RUN: belum ada yang ditulis ke sheet. Jalankan lagi dengan --write jika semua [OK].")
    return finish(rep, a, df, 1 if rep.errors else 0)


def finish(rep, a, df, code):
    out = Path(a.outdir)
    out.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if df is not None:
        df.to_csv(out / f"preview_rows_{stamp}.csv", index=False)
    rep.info(f"\nRINGKASAN: {len(rep.errors)} gagal, {len(rep.warnings)} peringatan")
    (out / f"laporan_validasi_{stamp}.txt").write_text("\n".join(rep.lines), encoding="utf-8")
    return code


if __name__ == "__main__":
    sys.exit(main())
