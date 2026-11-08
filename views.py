import calendar
import datetime as dt
import os

import pandas as pd
import streamlit as st

from . import config as cfgmod
from . import drive as drv
from . import services as svc
from .errors import friendly
from . import theme
from .widgets import guided_run, show_result


def quick_box(key, label, run, ready=True, msg="Pilih file dulu."):
    """Tombol cek cepat: hasil ditampilkan, tidak ada yang ditulis."""
    if st.button(f"{label}", key=f"qb_{key}", disabled=not ready):
        with st.spinner("Membaca file terpilih saja…"):
            st.session_state[f"q_{key}"] = run()
    if not ready:
        st.caption(msg)
    if st.session_state.get(f"q_{key}") is not None:
        show_result(st.session_state[f"q_{key}"], "Cek cepat")
    st.divider()


# ------------------------------------------------------------------ Beranda
def home(ctx):
    theme.hero("Nivea Data Hub", "Satu tempat untuk semua pekerjaan data Nivea. Pilih menu di kiri, lalu ikuti dua langkah: "
               "Cek dulu, kemudian Tulis. Google Sheet tidak berubah sebelum kamu menekan tombol tulis.")
    theme.cards([
        ("01", "Tracker Video & Live", "Isi tab SV, Exploded SV, dan LS dari export harian."),
        ("02", "GMV Harian Creator & MCN", "Isi tracker GMV dari file Creator_List dan Agency list."),
        ("03", "Order Affiliate + Patner", "Gabungkan Affiliate_Order dan Patner, tanpa data dobel."),
        ("04", "Product Performance Harian", "Gabungkan file product_list harian ke tab Raw."),
        ("05", "TAP + SC", "Unggah file TAP dan SC, hasilnya masuk ke Raw Data TAP&SC."),
    ])
    st.divider()
    st.subheader("Cek koneksi")
    if ctx.email:
        st.info(f"Aplikasi masuk ke Google sebagai: **{ctx.email}**. Semua folder & spreadsheet harus di-share "
                f"(Editor) ke email ini.")
    if st.button("Tes akses ke semua folder & sheet"):
        with st.spinner("Mengecek…"):
            for name, ok, note in svc.check_access(ctx):
                (st.success if ok else st.error)(f"{name}: {note}")


# ------------------------------------------------------------------ Tracker SV / LS
def tracker(ctx):
    theme.header("Tracker Video & Live (SV / LS)")
    st.write("Mengisi tab **SV** (Video), **Exploded SV**, dan **LS** (Live) dari file export harian di Google Drive.")
    mode_lbl = st.radio("Mau isi apa?", ["SV (Video)", "LS (Live)", "Keduanya"], horizontal=True)
    mode = {"SV (Video)": "sv", "LS (Live)": "ls", "Keduanya": "all"}[mode_lbl]
    today = dt.date.today()
    first = today.replace(day=1)
    BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
             "September", "Oktober", "November", "Desember"]
    cara = st.radio("Periode", ["Pilih bulan & tahun", "Atur tanggal sendiri"], horizontal=True)
    if cara == "Pilih bulan & tahun":
        a, b = st.columns(2)
        bln = a.selectbox("Bulan", BULAN, index=today.month - 1)
        years = list(range(today.year - 3, today.year + 2))
        thn = b.selectbox("Tahun", years, index=years.index(today.year))
        m = BULAN.index(bln) + 1
        start = dt.date(thn, m, 1)
        end = dt.date(thn, m, calendar.monthrange(thn, m)[1])
    else:
        a, b = st.columns(2)
        start = a.date_input("Dari tanggal", first)
        end = b.date_input("Sampai tanggal", today)
    st.caption(f"Periode: {start:%d %b %Y} s/d {end:%d %b %Y}")
    with st.expander("Opsi lanjutan (boleh dibiarkan)"):
        upd = st.checkbox("Perbarui angka metrik untuk ID yang sudah ada di tracker", True)
        pb = st.checkbox("Tambahkan PID baru ke Product Bank (CAT/PRO dikosongkan)", True)
        exp = st.checkbox("Isi juga tab Exploded SV (satu baris per produk)", True)
        ref = st.checkbox("Timpa PID & Product Name baris lama dengan SEMUA produk", False)
    sig = (mode, start, end, upd, pb, exp, ref)
    st.subheader("Cek cepat (opsional)")
    st.caption("Pilih file export yang mau dilihat saja; hanya itu yang dibaca. Tidak membuka tracker.")
    qmode = "ls" if mode == "ls" else "sv"
    if mode == "all":
        qmode = {"SV": "sv", "LS": "ls"}[st.radio("Cek cepat untuk", ["SV", "LS"], horizontal=True)]
    qfolder = ctx.cfg["ls_folder_id"] if qmode == "ls" else ctx.cfg["daily_folder_id"]
    qf = _folder_picker(ctx, f"trk_{qmode}", qfolder, f"export {qmode.upper()}")
    quick_box("tracker", "Cek cepat: jumlah baris & GMV file terpilih",
              lambda: svc.quick_tracker(ctx, qmode, qf, start, end), ready=bool(qf))
    guided_run("tracker", sig, lambda w: svc.run_tracker(ctx, mode, start, end, w, upd, pb, exp, ref),
               ready=start <= end, not_ready_msg="Tanggal mulai harus sebelum tanggal akhir.")


# ------------------------------------------------------------------ Order
def _folder_picker(ctx, key, folder_id, title, name_contains=None):
    ck = f"files_{key}"
    if st.button(f"Muat daftar file — {title}", key=f"load_{key}"):
        with st.spinner("Membaca folder Drive…"):
            drive, _ = ctx.services()
            st.session_state[ck] = drv.list_folder(drive, folder_id, name_contains)
    files = st.session_state.get(ck)
    if files is None:
        st.caption("Klik tombol di atas untuk menampilkan file di folder.")
        return []
    if not files:
        st.warning("Folder kosong atau belum di-share ke akun aplikasi.")
        return []
    by_label = {drv.label(f): f["id"] for f in files}
    pick = st.multiselect(f"Pilih file {title}", list(by_label), key=f"pick_{key}")
    return [by_label[p] for p in pick]


def orders(ctx):
    theme.header("Order Affiliate + Patner")
    st.write("Menggabungkan file Affiliate_Order dan Affiliate_Patner ke tab **[Order] Creator+Patner**. "
             "Baris yang sudah ada / dobel antar file otomatis tidak dimasukkan lagi. "
             "Laporan pengecekan menampilkan jumlah baris, rentang Time Created, dan total Est. Commission Base.")
    c1, c2 = st.columns(2)
    with c1:
        o = _folder_picker(ctx, "order", ctx.cfg["affiliate_order_folder_id"], "Affiliate_Order")
    with c2:
        p = _folder_picker(ctx, "partner", ctx.cfg["affiliate_partner_folder_id"], "Affiliate_Patner")
    upd = st.checkbox("Perbarui status/komisi baris yang sudah ada di sheet", True)
    quick_box("orders", "Cek cepat: jumlah baris, Time Created & Est. Commission Base file terpilih",
              lambda: svc.quick_orders(ctx, o, p), ready=bool(o or p))
    guided_run("orders", (tuple(o), tuple(p), upd), lambda w: svc.run_orders(ctx, o, p, w, upd),
               ready=bool(o or p), not_ready_msg="Pilih minimal satu file Order atau Patner.")


# ------------------------------------------------------------------ Product
def product(ctx):
    theme.header("Product Performance Harian")
    st.write("Menggabungkan file harian `product_list_YYYYMMDD` ke tab **[Product Performance] Daily November (Raw)**. "
             "Tanggal diambil dari nama file; data tanggal yang sudah ada tidak dobel.")
    f = _folder_picker(ctx, "product", ctx.cfg["product_folder_id"], "Product Performance", "product_list")
    ref = st.checkbox("Timpa baris (Tanggal + Product ID) yang sudah ada dengan data file terbaru", False)
    quick_box("product", "Cek cepat: jumlah produk & GMV file terpilih",
              lambda: svc.quick_product(ctx, f), ready=bool(f))
    guided_run("product", (tuple(f), ref), lambda w: svc.run_product(ctx, f, w, ref),
               ready=bool(f), not_ready_msg="Pilih minimal satu file.")


# ------------------------------------------------------------------ TAP + SC
def tapsc(ctx):
    theme.header("TAP + SC")
    st.write("Unggah file TAP (.xlsx) dan SC affiliate_orders (.csv). Hasilnya ditambahkan ke tab **Raw Data TAP&SC**.")
    c1, c2 = st.columns(2)
    tap = c1.file_uploader("File TAP (.xlsx)", type=["xlsx", "xls"])
    sc = c2.file_uploader("File SC affiliate_orders (.csv)", type=["csv"])
    excl = st.text_input("Order Status SC yang dibuang (pisah koma, kosongkan jika tidak ada)", "")
    overlap = st.checkbox("Izinkan tulis walau tanggal+source sudah ada di sheet (berisiko dobel)", False)
    up = os.path.join(ctx.workdir, "uploads")
    os.makedirs(up, exist_ok=True)

    def save(f):
        p = os.path.join(up, os.path.basename(f.name))
        with open(p, "wb") as fh:
            fh.write(f.getbuffer())
        return p

    statuses = tuple(s.strip() for s in excl.split(",") if s.strip())

    def run(w):
        res = svc.run_tapsc(ctx, save(tap), save(sc), w, overlap, statuses)
        st.session_state["tapsc_preview"] = res.extra.get("preview_csv")
        return res

    sig = (tap.name if tap else None, tap.size if tap else 0, sc.name if sc else None, sc.size if sc else 0, statuses, overlap)
    guided_run("tapsc", sig, run, ready=bool(tap and sc), not_ready_msg="Unggah kedua file dulu.")
    pv = st.session_state.get("tapsc_preview")
    if pv and os.path.exists(pv):
        with st.expander("Pratinjau 50 baris pertama yang akan ditulis"):
            st.dataframe(pd.read_csv(pv, nrows=50))


# ------------------------------------------------------------------ GMV Daily
def gmv(ctx):
    theme.header("GMV Harian Creator & MCN")
    st.write("Mengisi tracker GMV (tab Celeb, MCN, TOP VIP, RISING, KLUB VIP, Community) dari file "
             "`Creator_List_*` dan file Agency list (MCN) di Drive.")
    if st.button("Muat daftar file di Drive"):
        with st.spinner("Membaca folder Drive…"):
            try:
                st.session_state["gmv_entries"] = svc.list_gmv_entries(ctx)
            except Exception as e:  # noqa
                st.error(friendly(str(e)))
    entries = st.session_state.get("gmv_entries")
    if not entries:
        st.caption("Klik tombol di atas untuk menampilkan file.")
        return
    df = pd.DataFrame({
        "Pakai": False,
        "Jenis": ["Creator" if e["kind"] == "creator" else "MCN" for e in entries],
        "Tanggal": [e["date"].strftime("%d %b %Y") if e["date"] else "⚠ tidak dikenali" for e in entries],
        "Nama file": [e["name"] for e in entries],
    })
    edited = st.data_editor(df, hide_index=True, use_container_width=True, disabled=["Jenis", "Tanggal", "Nama file"],
                            key="gmv_tbl")
    chosen = tuple(e["id"] for e, use in zip(entries, edited["Pakai"]) if use)
    st.caption(f"{len(chosen)} file dipilih. Jika ada beberapa export untuk tanggal & laporan yang sama, "
               f"hanya yang terbaru dipakai (otomatis, tercatat di laporan).")
    if st.button("Cek cepat: hanya total GMV file terpilih", disabled=not chosen):
        with st.spinner("Membaca file terpilih saja…"):
            st.session_state["gmv_quick"] = svc.quick_gmv(ctx, entries, chosen)
    if st.session_state.get("gmv_quick") is not None:
        show_result(st.session_state["gmv_quick"], "Cek cepat")
    st.divider()
    st.caption("Pengecekan lengkap di bawah juga membaca seluruh tracker, jadi lebih lama. "
               "Pakai Cek cepat dulu untuk memastikan angka file benar.")
    guided_run("gmv", chosen, lambda w: svc.run_gmv(ctx, entries, chosen, w),
               ready=bool(chosen), not_ready_msg="Centang minimal satu file di tabel.")


# ------------------------------------------------------------------ Pengaturan
def settings(ctx):
    theme.header("Pengaturan")
    st.write("ID folder / spreadsheet yang dipakai. Ubah di sini berlaku untuk semua pengguna **sampai server restart**. "
             "Untuk perubahan permanen, isi bagian `[config]` di Streamlit Secrets (lihat PANDUAN_DEPLOY.md).")
    new = {}
    for k, lbl in cfgmod.LABELS.items():
        new[k] = st.text_input(lbl, ctx.cfg.get(k, ""), key=f"cfg_{k}")
    if st.button("Terapkan"):
        from . import auth
        ctx.cfg.update({k: v.strip() for k, v in new.items()})
        auth.write_runtime_files(ctx.workdir, ctx.secrets, ctx.cfg)
        st.success("Tersimpan untuk sesi server ini.")
