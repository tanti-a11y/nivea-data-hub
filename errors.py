"""Ubah pesan error teknis jadi kalimat yang dimengerti pengguna awam."""
import re

RULES = [
    (r"429|Quota exceeded|Write requests",
     "Google membatasi jumlah penulisan per menit. Tunggu sekitar 1 menit lalu klik tombolnya lagi. "
     "Aman diulang: data yang sudah masuk tidak akan dobel."),
    (r"\[403\]|PERMISSION_DENIED|insufficient|does not have permission|caller does not have",
     "Akses ditolak. Pastikan folder / spreadsheet-nya sudah di-share (sebagai Editor) ke akun yang dipakai aplikasi ini. "
     "Emailnya ada di halaman Beranda."),
    (r"\[404\]|not found|File not found|Requested entity was not found",
     "File / folder / sheet tidak ditemukan. Cek ID-nya di halaman Pengaturan, dan pastikan sudah di-share ke akun aplikasi."),
    (r"Kolom '.+' tidak ada",
     "Ada kolom yang tidak ketemu di header sheet. Cek apakah nama kolom atau baris header di sheet berubah."),
    (r"invalid_grant|Token has been expired|refresh",
     "Izin login Google sudah kedaluwarsa. Minta admin membuat ulang refresh token (lihat PANDUAN_DEPLOY.md)."),
    (r"timed out|Connection|ConnectionReset|RemoteDisconnected|ssl",
     "Koneksi ke Google terputus. Coba klik lagi - file yang sudah terbaca tidak diulang dari awal."),
    (r"Kredensial Google belum diisi",
     "Kredensial Google belum dipasang di Streamlit secrets. Lihat PANDUAN_DEPLOY.md langkah 3."),
    (r"ModuleNotFoundError|Library belum terpasang",
     "Ada library yang belum terpasang di server. Pastikan requirements.txt ikut ter-deploy."),
]


def friendly(msg):
    for pat, text in RULES:
        if re.search(pat, msg or "", re.I):
            return text
    return "Terjadi kesalahan. Buka 'Lihat detail laporan' untuk pesan lengkapnya."
