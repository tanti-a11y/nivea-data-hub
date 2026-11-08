"""Uji asap tampilan tanpa Streamlit asli: semua halaman harus bisa dirender tanpa NameError/TypeError."""
import os, sys, tempfile
from unittest.mock import MagicMock
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class SS(dict):
    def setdefault(self, k, v=None): return dict.setdefault(self, k, v)
st = MagicMock()
st.session_state = SS()
def _col():
    c = MagicMock()
    c.selectbox.side_effect = lambda label, opts, index=0, **k: list(opts)[index]
    import datetime as _d
    c.date_input.side_effect = lambda label, value=None, **k: value or _d.date.today()
    return c
st.columns.side_effect = lambda n, *a, **k: [_col() for _ in range(n if isinstance(n, int) else len(n))]
st.button.return_value = False
st.checkbox.side_effect = lambda label, value=False, **k: value
st.radio.side_effect = lambda label, opts, **k: opts[0]
st.selectbox.side_effect = lambda label, opts, **k: opts[0]
st.text_input.side_effect = lambda label, value="", **k: value
st.multiselect.return_value = []
st.file_uploader.return_value = None
sys.modules["streamlit"] = st

from hub import views, auth
from hub.config import build_config
from hub.ctx import Ctx
d = tempfile.mkdtemp()
sec = {"gcp_service_account": {"type": "service_account", "client_email": "x@y"}}
p = auth.write_runtime_files(d, sec, build_config())
ctx = Ctx(sec, build_config(), d, p)
for name in ["home", "tracker", "orders", "product", "tapsc", "gmv", "settings"]:
    getattr(views, name)(ctx)
    print("render ok:", name)

# alur guided_run: tombol Cek dulu ditekan -> hasil gagal ditampilkan dengan pesan ramah
from hub import widgets
from hub.runner import RunResult
st.button.side_effect = lambda label, **k: "Cek dulu" in label
widgets.guided_run("t", ("a",), lambda w: RunResult(False, "boom", "Quota exceeded 429"))
assert any("Google membatasi" in str(c) for c in st.warning.call_args_list), "pesan ramah tidak muncul"
widgets.guided_run("t2", ("a",), lambda w: RunResult(True, "semua OK", ""))
assert st.session_state["guided_t2"]["dry"].ok
print("guided_run ok")
