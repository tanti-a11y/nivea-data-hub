# Nivea Data Hub — Panduan Pasang Online (Streamlit Community Cloud)

## Isi aplikasi
| Menu | Fungsi (dari skrip lamamu) |
|---|---|
| 🎬 Tracker Video & Live | `nivea_tracker.py` — isi tab SV, Exploded SV, LS |
| 📈 GMV Harian Creator & MCN | `sync_nivea_tracker.py` — isi tracker GMV dari Creator_List & Agency list |
| 🧾 Order Affiliate + Patner | `affiliate_merge.py` mode orders |
| 📦 Product Performance Harian | `affiliate_merge.py` mode product |
| 🎯 TAP + SC | `nivea_tap_sc_to_sheet.py` |

Semua menu memakai alur sama: **Cek dulu** (tidak menulis) → baca laporan → centang → **Tulis**.

## 1. Siapkan akun Google untuk aplikasi
Pakai **Service Account** yang sudah kamu punya (`service_account.json`).
Buka semua folder Drive & spreadsheet yang dipakai → **Share** → tambahkan email service account
(`...@...iam.gserviceaccount.com`) sebagai **Editor**. Daftar yang perlu di-share ada di menu *Beranda → Tes akses*.

Kalau organisasi memblokir key Service Account untuk fitur GMV: pakai cara B di `secrets.toml.example`
(jalankan `buat_refresh_token.py` di komputermu). Catatan: jika OAuth consent screen berstatus *Testing*,
refresh token kedaluwarsa tiap 7 hari — atur ke **Internal** (khusus akun organisasi) agar tidak kedaluwarsa.

## 2. Taruh kode di GitHub (repo **Private**)
Unggah isi folder ini. File rahasia (`secrets.toml`, `*.json` kredensial) sudah dikecualikan oleh `.gitignore`.

## 3. Deploy
1. https://share.streamlit.io → **New app** → pilih repo → *Main file path*: `app.py`.
2. **Advanced settings → Secrets**: tempel isi `.streamlit/secrets.toml.example` yang sudah kamu isi
   (kata sandi aplikasi + kredensial Google).
3. **Deploy**. Bagikan link-nya ke tim; mereka cukup memasukkan kata sandi.

## Login Google (hanya email @warnaemasindonesia.com)
Aplikasi memeriksa email akun Google yang login: harus berakhiran `@warnaemasindonesia.com` dan sudah terverifikasi.
Akun lain (mis. Gmail) ditolak, walaupun punya link-nya.

1. Google Cloud Console → **APIs & Services → Credentials → Create credentials → OAuth client ID**.
   Jenis: **Web application** (bukan Desktop; ini terpisah dari client untuk Drive).
2. **Authorized redirect URIs**, tambahkan:
   - `http://localhost:8501/oauth2callback` (untuk uji di komputer)
   - `https://NAMA-APP.streamlit.app/oauth2callback` (setelah deploy)
3. **OAuth consent screen** → User type **Internal** (membatasi ke akun organisasi, lapis pengaman kedua).
4. Isi bagian `[auth]` di secrets (`client_id`, `client_secret`, `redirect_uri`, dan `cookie_secret` berupa teks acak panjang).
5. Restart app. Halaman depan sekarang menampilkan tombol **Masuk dengan Google**.

Catatan: perlu Streamlit 1.42 atau lebih baru (sudah ada di requirements.txt). Tanpa blok `[auth]`, aplikasi memakai `app_password` sebagai cadangan.

## Hal penting
- ⚠️ **Putar (rotate) client secret OAuth lama** di Google Cloud: ia sempat tertanam di `nivea_tap_sc_to_sheet.py`
  (sudah dihapus dari versi ini). `token.json` / `credentials.json` / `service_account.json` jangan pernah di-upload ke GitHub.
- Server Streamlit Cloud hanya menyimpan file sementara. Backup nilai lama GMV/laporan ada di tampilan
  (tombol *Unduh laporan*) dan riwayat versi Google Sheets.
- Proses besar (47 ribu baris Order, tracker GMV) bisa memakan beberapa menit — jangan tutup halaman.
- Error kuota Google (429): tunggu 1 menit dan klik lagi; aman diulang.
- Dashboard HTML lama (`Dashboard_Harian_Nivea.html`) tidak ikut karena berupa file statis; bisa dijadikan menu terpisah kalau mau.

## Menjalankan di komputer (opsional)
```
pip install -r requirements.txt
mkdir .streamlit   (isi secrets.toml dari contoh)
py -m streamlit run app.py
```
