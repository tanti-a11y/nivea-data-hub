"""Jembatan antara tampilan (Streamlit) dan engine lama. Tanpa kode Streamlit -> bisa dites sendiri."""
import glob
import importlib
import os
from contextlib import contextmanager

from . import drive as drv
from .runner import RunResult, run_main


def _engine(name):
    return importlib.import_module(f"engines.{name}")


def _chdir_wrap(workdir, fn):
    def inner():
        old = os.getcwd()
        os.chdir(workdir)
        try:
            return fn()
        finally:
            os.chdir(old)
    return inner


def _latest(workdir, pattern):
    files = glob.glob(os.path.join(workdir, pattern), recursive=True)
    return max(files, key=os.path.getmtime) if files else None


# ---------------------------------------------------------------- Tracker SV / LS
def run_tracker(ctx, mode, start, end, write, update_existing=True, add_product_bank=True,
                exploded=True, refresh_products=False):
    eng = _engine("nivea_tracker")
    argv = ["--config", ctx.config_path, "--mode", mode, "--start", f"{start:%Y-%m-%d}", "--end", f"{end:%Y-%m-%d}", "--yes"]
    if not write:
        argv.append("--dry-run")
    if not update_existing:
        argv.append("--no-update-existing")
    if not add_product_bank:
        argv.append("--no-product-bank")
    if not exploded:
        argv.append("--no-exploded")
    if refresh_products:
        argv.append("--refresh-products")
    return run_main(eng.main, argv)


# ---------------------------------------------------------------- Order Affiliate + Patner
def run_orders(ctx, order_ids, partner_ids, write, update_existing=True):
    eng = _engine("affiliate_merge")
    argv = ["--config", ctx.config_path, "--mode", "orders", "--yes"]
    if order_ids:
        argv += ["--order", ",".join(order_ids)]
    if partner_ids:
        argv += ["--partner", ",".join(partner_ids)]
    if not write:
        argv.append("--dry-run")
    if not update_existing:
        argv.append("--no-update-existing")
    return run_main(eng.main, argv)


# ---------------------------------------------------------------- Product Performance Daily
def run_product(ctx, file_ids, write, refresh=False):
    eng = _engine("affiliate_merge")
    argv = ["--config", ctx.config_path, "--mode", "product", "--yes", "--product", ",".join(file_ids)]
    if not write:
        argv.append("--dry-run")
    if refresh:
        argv.append("--refresh")
    return run_main(eng.main, argv)


# ---------------------------------------------------------------- TAP + SC
def run_tapsc(ctx, tap_path, sc_path, write, allow_overlap=False, exclude_status=()):
    eng = _engine("nivea_tap_sc_to_sheet")
    outdir = os.path.join(ctx.workdir, "tapsc")
    os.makedirs(outdir, exist_ok=True)
    argv = ["--tap", tap_path, "--sc", sc_path, "--outdir", outdir]
    if write:
        argv.append("--write")
    if allow_overlap:
        argv.append("--allow-overlap")
    if exclude_status:
        argv += ["--sc-exclude-status", *exclude_status]
    _, sheets = ctx.services()
    res = run_main(eng.main, argv, patches={(eng, "get_service"): (lambda *a, **k: sheets)})
    res.extra["preview_csv"] = _latest(outdir, "preview_rows_*.csv")
    res.extra["report_txt"] = _latest(outdir, "laporan_validasi_*.txt")
    return res


# ---------------------------------------------------------------- GMV Daily (Creator_List + MCN)
def list_gmv_entries(ctx):
    eng = _engine("sync_nivea_tracker")
    drive, _ = ctx.services()
    entries = [dict(f, kind="creator") for f in eng.list_candidate_files(drive)]
    try:
        entries += [dict(f, kind="mcn") for f in eng.list_mcn_files(drive)]
    except Exception:
        pass
    entries = eng._sorted_entries(entries)
    for e in entries:
        e["date"] = eng.entry_date(e)
    return entries


def run_gmv(ctx, entries, chosen_ids, write):
    eng = _engine("sync_nivea_tracker")
    drive, sheets = ctx.services()
    chosen = [e for e in entries if e["id"] in set(chosen_ids)]
    patches = {
        (eng, "get_services"): (lambda: (drive, sheets)),
        (eng, "prompt_user_selection"): (lambda ents, on_preview=None: chosen),
        (eng, "BACKUP_DIR"): os.path.join(ctx.workdir, "backup_tracker"),
        (eng, "CACHE_DIR"): os.path.join(ctx.workdir, "cache_tracker"),
    }
    res = run_main(_chdir_wrap(ctx.workdir, eng.main), [], answers=["y" if write else "n"], patches=patches)
    res.extra["preview_xlsx"] = _latest(ctx.workdir, "Pratinjau_isi_tracker_*.xlsx")
    return res


def quick_gmv(ctx, entries, chosen_ids):
    """Cek cepat: hanya total GMV tiap file terpilih (tracker tidak dibaca/disentuh)."""
    eng = _engine("sync_nivea_tracker")
    drive, _ = ctx.services()
    chosen = [e for e in entries if e["id"] in set(chosen_ids)]
    patches = {(eng, "CACHE_DIR"): os.path.join(ctx.workdir, "cache_tracker")}
    return run_main(lambda: eng.print_gmv_preview(drive, chosen), [], patches=patches)


# ---------------------------------------------------------------- Cek cepat (tanpa membuka sheet output)
def _cred_paths(ctx):
    cred = "service_account.json" if os.path.exists(os.path.join(ctx.workdir, "service_account.json")) else "credentials.json"
    return os.path.join(ctx.workdir, cred), os.path.join(ctx.workdir, "token.json")


def quick_orders(ctx, order_ids, partner_ids):
    """Hanya baca file terpilih: jumlah baris, rentang Time Created, total Est. Commission Base."""
    eng = _engine("affiliate_merge")

    def go():
        cfg = eng.deep_cfg(dict(ctx.cfg))
        labels = [cfg["source_label_order"], cfg["source_label_partner"]]
        g = eng.Google(*_cred_paths(ctx))
        stats = []
        for label, folder, ids in ((labels[0], cfg["affiliate_order_folder_id"], order_ids),
                                   (labels[1], cfg["affiliate_partner_folder_id"], partner_ids)):
            if not ids:
                continue
            metas = {m["id"]: m for m in g.list_folder(folder)}
            for i in ids:
                m = metas[i]
                print(f"   membaca {m['name']} ...")
                recs, _ = eng.build_records(g.read_file(m), label, cfg, {})
                stats.append((label, m["name"], eng.rec_stats(recs, cfg)))
        eng.print_file_report(stats)
        print("\n(cek cepat) Belum membandingkan dengan isi sheet, jadi belum tahu berapa yang dobel/baru.")
    return run_main(go, [])


def quick_product(ctx, file_ids):
    eng = _engine("affiliate_merge")

    def go():
        cfg = eng.deep_cfg(dict(ctx.cfg))
        g = eng.Google(*_cred_paths(ctx))
        metas = {m["id"]: m for m in g.list_folder(cfg["product_folder_id"])}
        from googleapiclient.http import MediaIoBaseDownload
        import io
        out_groups = out_header = None
        for i in file_ids:
            m = metas[i]
            print(f"   membaca {m['name']} ...")
            buf = io.BytesIO()
            dl = MediaIoBaseDownload(buf, g.drive.files().get_media(fileId=m["id"], supportsAllDrives=True))
            done = False
            while not done:
                _, done = dl.next_chunk()
            pf = eng.read_product_file(buf.getvalue(), m["name"])
            if out_header is None:
                keep = [(gp, h) for gp, h in zip(pf["group"], pf["header"]) if h.lower() not in ("product name", "product id")]
                out_header = ["Tanggal", "Product Name", "Product ID", "CAT", "PRO"] + [h for _, h in keep]
                out_groups = ["", "", "", "", ""] + [gp for gp, _ in keep]
            fdate, note = eng.product_file_date(m["name"], pf)
            recs, _ = eng.build_product_records(pf, fdate, out_groups, out_header, {}, m["name"], m["modifiedTime"], cfg)
            st = eng.pstats(recs)
            print(f"   {m['name']}")
            print(f"      Tanggal : {eng.fmt_date(fdate)}  |  {eng.fmt_n(st['rows'])} produk")
            print(f"      Jumlah GMV (All): {eng.fmt_n(st['est'])}  |  Jumlah Affiliate Attributed GMV: {eng.fmt_n(st['aff'])}")
        print("\n(cek cepat) Belum membandingkan dengan isi sheet.")
    return run_main(go, [])


def quick_tracker(ctx, mode, file_ids, start, end):
    """Baca file export terpilih saja: jumlah baris dalam periode (dan total GMV bila ada)."""
    eng = _engine("nivea_tracker")

    def go():
        spec = eng.SPECS[mode]
        cfg = dict(ctx.cfg)
        g = eng.Google(*_cred_paths(ctx))
        folder = cfg[spec["folder_cfg"]]
        metas = {m["id"]: m for m in g.list_folder(folder)}
        tot_rows = 0
        tot_gmv = 0.0
        for i in file_ids:
            m = metas[i]
            content, fname = g.download(m)
            df = eng._read_export(content, fname)
            if df is None:
                print(f"   {m['name']}: bukan file export {spec['label']}")
                continue
            colmap = eng.resolve_columns(df.columns, spec, cfg)
            if "_id" not in colmap or "_time" not in colmap:
                print(f"   {m['name']}: kolom ID/waktu tidak dikenali")
                continue
            order = eng.frame_date_order(df, colmap, cfg, spec["date_order"])
            rows, gmv, ids = 0, 0.0, set()
            for _, r in df.iterrows():
                t = eng.parse_time(r.get(colmap["_time"]), order)
                if t is None or t.date() < start or t.date() > end:
                    continue
                vid = eng.clean_id(r.get(colmap["_id"]))
                if not vid or vid in ids:
                    continue
                ids.add(vid)
                rows += 1
                if "GMV" in colmap:
                    v = eng.parse_number(r.get(colmap["GMV"]))
                    gmv += v if isinstance(v, (int, float)) else 0.0
            tot_rows += rows
            tot_gmv += gmv
            print(f"   {m['name']}: {rows:,} {spec['id_col']} unik di periode" + (f" | GMV Rp {gmv:,.0f}" if "GMV" in colmap else ""))
        print(f"\nTOTAL: {tot_rows:,} baris" + (f" | GMV Rp {tot_gmv:,.0f}" if tot_gmv else "")
              + "\n(cek cepat) ID yang sama di beberapa file bisa terhitung dua kali di total ini; Cek dulu (lengkap) menyaringnya.")
    return run_main(go, [])


# ---------------------------------------------------------------- Cek koneksi
def check_access(ctx):
    """Kembalikan daftar (nama, ok, keterangan)."""
    out = []
    try:
        drive, sheets = ctx.services()
        # uji token betul-betul (bukan sekadar membaca secrets)
        who = drive.about().get(fields="user(emailAddress)").execute().get("user", {}).get("emailAddress", "")
        out.append(("Login Google", True, who or ctx.email or "OK"))
    except Exception as e:  # noqa
        msg = str(e)
        hint = ""
        if "invalid_grant" in msg or "refresh" in msg.lower() or "invalid_client" in msg:
            hint = " -> refresh_token salah/kedaluwarsa/belum diisi. Buat ulang dengan buat_refresh_token.py."
        elif "has not been used" in msg or "accessNotConfigured" in msg or "SERVICE_DISABLED" in msg:
            hint = " -> Google Drive API / Sheets API belum di-Enable di Google Cloud."
        return [("Login Google", False, msg[:300] + hint)]
    cfg = ctx.cfg
    checks = [
        ("Folder export SV (Video)", "daily_folder_id", "folder"),
        ("Folder export LS (Live)", "ls_folder_id", "folder"),
        ("Folder Affiliate_Order", "affiliate_order_folder_id", "folder"),
        ("Folder Affiliate_Patner", "affiliate_partner_folder_id", "folder"),
        ("Folder Product Performance", "product_folder_id", "folder"),
        ("Spreadsheet tracker SV/LS", "tracker_spreadsheet_id", "file"),
        ("Spreadsheet output Order/Product", "merge_output_spreadsheet_id", "file"),
        ("Product Bank", "product_bank_file_id", "file"),
        ("Creator Tiering", "creator_tiering_file_id", "file"),
    ]
    for name, key, _ in checks:
        try:
            m = drive.files().get(fileId=cfg[key], fields="name", supportsAllDrives=True).execute()
            out.append((name, True, m.get("name", "")))
        except Exception as e:  # noqa
            out.append((name, False, "tidak bisa dibuka: " + str(e)[:200]))
    return out
