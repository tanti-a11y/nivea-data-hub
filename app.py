import streamlit as st

from hub import theme, views
from hub.access import is_allowed
from hub.ctx import make_ctx

st.set_page_config(page_title="Nivea Data Hub", page_icon=theme.page_icon(), layout="wide")
theme.inject()


@st.cache_resource
def get_ctx():
    return make_ctx(st.secrets)


def _user():
    return getattr(st, "user", None) or getattr(st, "experimental_user", None)


def gate():
    """Kembalikan email pengguna (atau '' bila tanpa login). Hentikan halaman bila belum boleh masuk."""
    domain = st.secrets.get("allowed_domain", "warnaemasindonesia.com")
    extra = list(st.secrets.get("allowed_emails", []))

    # 1) Login Google (disarankan): hanya email @domain perusahaan
    if "auth" in st.secrets:
        u = _user()
        if not getattr(u, "is_logged_in", False):
            theme.login_card(f"Masuk dengan akun Google kantor (@{domain}).")
            _, mid, _ = st.columns([1, 1.2, 1])
            with mid:
                st.button("Masuk dengan Google", type="primary", use_container_width=True, on_click=st.login)
            st.stop()
        email = getattr(u, "email", "") or ""
        if not is_allowed(email, getattr(u, "email_verified", None), domain, extra):
            theme.login_card(f"Akun <b>{email}</b> tidak punya akses. Gunakan email @{domain}.")
            _, mid, _ = st.columns([1, 1.2, 1])
            with mid:
                st.button("Keluar dan ganti akun", use_container_width=True, on_click=st.logout)
            st.stop()
        return email

    # 2) Cadangan: kata sandi bersama
    pw = st.secrets.get("app_password")
    if pw and not st.session_state.get("auth_ok"):
        theme.login_card("Masukkan kata sandi tim untuk melanjutkan.")
        _, mid, _ = st.columns([1, 1.2, 1])
        with mid:
            typed = st.text_input("Kata sandi", type="password", label_visibility="collapsed", placeholder="Kata sandi")
            if typed:
                if typed == pw:
                    st.session_state["auth_ok"] = True
                    st.rerun()
                st.error("Kata sandi salah.")
        st.stop()
    return ""


email = gate()
try:
    ctx = get_ctx()
except Exception as e:  # noqa
    theme.login_card("Aplikasi belum siap: kredensial Google belum dipasang di Streamlit Secrets.")
    st.caption(str(e))
    st.stop()

MENU = [
    ("Beranda", "beranda", views.home),
    ("Tracker Video & Live", "tracker", views.tracker),
    ("GMV Harian & MCN", "gmv", views.gmv),
    ("Order Affiliate + Patner", "order", views.orders),
    ("Product Performance", "product", views.product),
    ("TAP + SC", "tapsc", views.tapsc),
    ("Pengaturan", "pengaturan", views.settings),
]


def _make(fn):
    def page():
        fn(ctx)
    return page


pages = [st.Page(_make(fn), title=title, url_path=slug, default=(i == 0)) for i, (title, slug, fn) in enumerate(MENU)]
nav = st.navigation(pages, position="hidden")
with st.sidebar:
    theme.sidebar_brand()
    current = getattr(nav, "title", None)
    for pg in pages:
        if pg.title == current:   # menu aktif: kotak putih (digambar sendiri, tidak bergantung versi Streamlit)
            st.markdown(f'<div class="nv-active">{pg.title}</div>', unsafe_allow_html=True)
        else:
            st.page_link(pg, label=pg.title)
    if email:
        st.markdown(f'<div class="nv-foot">Masuk sebagai<br><b>{email}</b></div>', unsafe_allow_html=True)
        st.button("Keluar", on_click=st.logout)
    theme.sidebar_footer()
nav.run()
