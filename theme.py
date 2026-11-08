"""Tampilan: warna biru Nivea, logo, menu samping berbentuk daftar."""
import base64
import os

import streamlit as st

ASSETS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
LOGO = os.path.join(ASSETS, "logo.png")
NAVY = "#222F87"
NAVY_DARK = "#0B2A6F"


def logo_b64():
    try:
        with open(LOGO, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except OSError:
        return ""


def page_icon():
    try:
        from PIL import Image
        return Image.open(LOGO)
    except Exception:  # noqa
        return "🧴"


CSS = """
<style>
:root { --nv:#222F87; --nvd:#0B2A6F; --bg:#F1F4FB; --line:#DCE4F5; }
html, body, .stApp { background: var(--bg); }
#MainMenu, footer { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }
.block-container { padding-top: 2.2rem; padding-bottom: 4rem; max-width: 1180px; }

/* ---------- sidebar ---------- */
section[data-testid="stSidebar"] { background: linear-gradient(180deg, #0B2A6F 0%, #102F7C 55%, #0B2A6F 100%); border-right: 0; }
section[data-testid="stSidebar"] > div { padding-top: 1.2rem; }
section[data-testid="stSidebar"] * { color: #D3DEF6; }
.nv-brand { display:flex; align-items:center; gap:12px; padding: 4px 6px 18px 6px; }
.nv-brand img { width:46px; height:46px; border-radius:50%; box-shadow: 0 0 0 3px rgba(255,255,255,.18); }
.nv-brand .t { font-weight:700; font-size:1.08rem; color:#fff !important; line-height:1.15; }
.nv-brand .s { font-size:.72rem; color:#9FB3E6 !important; letter-spacing:.04em; }
.nv-sep { height:1px; background:rgba(255,255,255,.12); margin: 2px 0 14px 0; }
/* menu = daftar tautan (st.page_link), tanpa bulatan */
section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] { gap: .35rem; }
section[data-testid="stSidebar"] [data-testid="stElementContainer"], section[data-testid="stSidebar"] .element-container { margin:0; }
.nv-active { background:#FFFFFF; color:#0B2A6F; font-weight:700; font-size:1rem; padding: 14px 18px; border-radius:14px;
  box-shadow: 0 6px 18px rgba(0,0,0,.22); }
section[data-testid="stSidebar"] [data-testid="stPageLink"] { margin: 0 0 6px 0; }
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"] {
  width:100%; padding: 14px 18px; border-radius: 14px; background: transparent; text-decoration:none; transition: background .15s;
}
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"]:hover { background: rgba(255,255,255,.10); }
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"] p,
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"] span { font-size:1rem; margin:0; color:#D3DEF6 !important; }
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"][aria-current="page"] { background:#FFFFFF; box-shadow: 0 6px 18px rgba(0,0,0,.22); }
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"][aria-current="page"] p,
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"][aria-current="page"] span { color: var(--nvd) !important; font-weight:700; }
.nv-foot { position: relative; margin-top: 28px; font-size:.78rem; color:#9FB3E6 !important; line-height:1.5; padding: 0 8px; }
.nv-foot b { color:#fff !important; }

/* ---------- header halaman ---------- */
.nv-hero { background: linear-gradient(120deg, #0B2A6F 0%, #222F87 60%, #3445B0 100%); border-radius: 20px; padding: 30px 34px;
  display:flex; align-items:center; gap:24px; box-shadow: 0 10px 30px rgba(11,42,111,.22); margin-bottom: 22px; }
.nv-hero img { width:84px; height:84px; border-radius:50%; box-shadow: 0 0 0 5px rgba(255,255,255,.18); }
.nv-hero h1 { color:#fff !important; margin:0; font-size:2.1rem; padding:0; }
.nv-hero p { color:#C9D6F5 !important; margin:6px 0 0 0; font-size:1.02rem; }
.nv-head { border-left: 6px solid var(--nv); padding: 4px 0 4px 16px; margin-bottom: 14px; }
.nv-head .e { font-size:.72rem; letter-spacing:.14em; color:#6C7CC4; font-weight:700; }
.nv-head h1 { color: var(--nvd); font-size: 1.9rem; margin: 2px 0 0 0; padding:0; }
.nv-cards { display:grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap:16px; margin: 6px 0 10px 0; }
.nv-card { background:#fff; border:1px solid var(--line); border-radius:16px; padding:18px 20px; box-shadow: 0 2px 10px rgba(34,47,135,.06); }
.nv-card .n { color: var(--nv); font-weight:800; font-size:.8rem; letter-spacing:.1em; }
.nv-card .h { color: var(--nvd); font-weight:700; font-size:1.05rem; margin:4px 0 6px 0; }
.nv-card .d { color:#55607F; font-size:.9rem; line-height:1.5; }

/* ---------- komponen ---------- */
h2, h3, h4 { color: var(--nvd); }
.stButton > button, .stDownloadButton > button { border-radius: 10px; border: 1px solid var(--line); background:#fff; color: var(--nvd); font-weight:600; padding: .5rem 1.1rem; }
.stButton > button:hover, .stDownloadButton > button:hover { border-color: var(--nv); color: var(--nv); }
.stButton > button[kind="primary"] { background: var(--nv); color:#fff; border: 0; box-shadow: 0 4px 12px rgba(34,47,135,.28); }
.stButton > button[kind="primary"]:hover { background: var(--nvd); color:#fff; }
div[data-testid="stExpander"] { background:#fff; border:1px solid var(--line); border-radius:14px; }
div[data-testid="stAlert"] { border-radius: 12px; }
div[data-baseweb="select"] > div, .stTextInput input, .stTextArea textarea, .stDateInput input { border-radius: 10px; }
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { border:1px solid var(--line); border-radius: 12px; overflow:hidden; }
hr { border-color: var(--line); }
</style>
"""


def inject():
    st.markdown(CSS, unsafe_allow_html=True)


def sidebar_brand():
    b = logo_b64()
    img = f'<img src="data:image/png;base64,{b}">' if b else ""
    st.markdown(f'<div class="nv-brand">{img}<div><div class="t">Nivea Data Hub</div>'
                f'<div class="s">TIM DATA ANALYST</div></div></div><div class="nv-sep"></div>', unsafe_allow_html=True)


def sidebar_footer():
    st.markdown('<div class="nv-foot">Alur kerja:<br><b>Cek dulu</b> → baca laporan → <b>Tulis</b><br><br>'
                'Warna Emas Indonesia · Tim Data Analyst</div>', unsafe_allow_html=True)


def header(title, eyebrow="NIVEA DATA HUB"):
    st.markdown(f'<div class="nv-head"><div class="e">{eyebrow}</div><h1>{title}</h1></div>', unsafe_allow_html=True)


def hero(title, text):
    b = logo_b64()
    img = f'<img src="data:image/png;base64,{b}">' if b else ""
    st.markdown(f'<div class="nv-hero">{img}<div><h1>{title}</h1><p>{text}</p></div></div>', unsafe_allow_html=True)


def cards(items):
    html = "".join(f'<div class="nv-card"><div class="n">{n}</div><div class="h">{h}</div><div class="d">{d}</div></div>'
                   for n, h, d in items)
    st.markdown(f'<div class="nv-cards">{html}</div>', unsafe_allow_html=True)


LOGIN_CSS = """
<style>
.stApp { background: linear-gradient(135deg, #0B2A6F 0%, #222F87 55%, #3445B0 100%); }
section[data-testid="stSidebar"], button[data-testid="stSidebarCollapsedControl"] { display:none; }
.nv-login { max-width: 470px; margin: 7vh auto 18px auto; background:#fff; border-radius: 24px; padding: 40px 38px 30px 38px;
  text-align:center; box-shadow: 0 24px 60px rgba(0,0,0,.35); }
.nv-login img { width:104px; height:104px; border-radius:50%; box-shadow: 0 0 0 6px #E6EBFA; }
.nv-login h1 { color:#0B2A6F; font-size:1.9rem; margin: 18px 0 4px 0; padding:0; }
.nv-login p { color:#55607F; margin: 0 0 6px 0; font-size:.98rem; line-height:1.5; }
.nv-login .f { color:#8A95B5; font-size:.78rem; margin-top: 18px; }
</style>
"""


def login_card(subtitle, footer="Warna Emas Indonesia · Tim Data Analyst"):
    b = logo_b64()
    img = f'<img src="data:image/png;base64,{b}">' if b else ""
    st.markdown(LOGIN_CSS + f'<div class="nv-login">{img}<h1>Nivea Data Hub</h1><p>{subtitle}</p>'
                f'<div class="f">{footer}</div></div>', unsafe_allow_html=True)
