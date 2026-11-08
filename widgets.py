"""Komponen tampilan yang dipakai semua halaman: alur 'Cek dulu -> Tulis'."""
import streamlit as st

from .errors import friendly


def show_result(res, title):
    if res.ok:
        st.success(f"{title}: selesai tanpa error.")
    else:
        st.error(f"{title}: gagal.")
        st.warning(friendly(res.error + "\n" + res.text[-1500:]))
        if res.error:
            st.caption(f"Pesan teknis: {res.error}")
    with st.expander("Lihat detail laporan", expanded=True):
        st.text_area("laporan", res.text or "(kosong)", height=420, label_visibility="collapsed")
    st.download_button("Unduh laporan (.txt)", (res.text or "").encode("utf-8"),
                       file_name="laporan.txt", key=f"dl_{id(res)}")


def guided_run(key, sig, run, write_label="Tulis ke Google Sheet sekarang", ready=True, not_ready_msg=""):
    """run(write: bool) -> RunResult.  sig = tanda pilihan saat ini; kalau pilihan berubah, cek ulang wajib."""
    S = st.session_state.setdefault(f"guided_{key}", {})
    st.markdown("#### Langkah 1 — Cek dulu")
    st.caption("Hanya membaca & menghitung. **Belum ada yang ditulis** ke Google Sheet.")
    if not ready:
        st.info(not_ready_msg)
    if st.button("Cek dulu (aman)", key=f"chk_{key}", disabled=not ready, type="primary"):
        with st.spinner("Sedang membaca & memeriksa data… (bisa 1–3 menit untuk data besar)"):
            S.clear()
            S.update(dry=run(False), sig=sig, wrote=None)

    dry = S.get("dry")
    if dry is None:
        return
    if S.get("sig") != sig:
        st.warning("Pilihanmu berubah sejak pengecekan terakhir. Klik **Cek dulu** lagi sebelum menulis.")
        return
    show_result(dry, "Pengecekan")
    if not dry.ok:
        return

    st.markdown("#### Langkah 2 — Tulis ke Google Sheet")
    st.caption("Baca laporan di atas. Kalau angkanya sudah sesuai, centang lalu klik tulis.")
    ok = st.checkbox("Saya sudah membaca laporan dan datanya sudah benar", key=f"ok_{key}")
    if st.button(write_label, key=f"wr_{key}", disabled=not ok or S.get("wrote") is not None):
        with st.spinner("Menulis ke Google Sheet… jangan tutup halaman ini."):
            S["wrote"] = run(True)
    if S.get("wrote") is not None:
        show_result(S["wrote"], "Penulisan")
        if S["wrote"].ok:
            st.balloons()
        st.caption("Untuk menulis lagi, klik **Cek dulu** ulang.")
