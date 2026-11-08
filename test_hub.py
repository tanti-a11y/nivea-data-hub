import datetime as dt, json, os, sys, tempfile, types
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from hub import auth, runner, services as svc
from hub.config import build_config
from hub.ctx import Ctx

def t_runner():
    def ok():
        print("halo", sys.argv[1:])
    r = runner.run_main(ok, ["--a", "1"]); assert r.ok and "halo ['--a', '1']" in r.text
    def ex(): sys.exit("kolom hilang")
    r = runner.run_main(ex); assert not r.ok and r.error == "kolom hilang"
    def ask(): input("Lanjut? ")
    r = runner.run_main(ask); assert not r.ok and "interaktif" in r.error
    r = runner.run_main(ask, answers=["y"]); assert r.ok
    mod = types.ModuleType('m'); mod.f = lambda: 1
    def use(): print(mod.f())
    r = runner.run_main(use, patches={(mod, "f"): (lambda: 42)}); assert "42" in r.text and mod.f() == 1
    def boom(): raise ValueError("x")
    r = runner.run_main(boom); assert not r.ok and "ValueError" in r.error
    print("runner ok")

def t_auth():
    sa = {"gcp_service_account": {"type": "service_account", "client_email": "a@b.iam", "private_key": "K\\nL"}, "app_password": "x"}
    d = tempfile.mkdtemp(); p = auth.write_runtime_files(d, sa, build_config())
    cfg = json.load(open(p)); assert cfg["credentials_file"] == "service_account.json"
    assert json.load(open(os.path.join(d, "service_account.json")))["private_key"] == "K\nL"
    assert auth.service_account_email(sa) == "a@b.iam"
    oa = {"google_oauth": {"client_id": "i", "client_secret": "s", "refresh_token": "r"}}
    d = tempfile.mkdtemp(); auth.write_runtime_files(d, oa, build_config())
    tok = json.load(open(os.path.join(d, "token.json"))); assert tok["refresh_token"] == "r" and "spreadsheets" in tok["scopes"][1]
    assert "installed" in json.load(open(os.path.join(d, "credentials.json")))
    try: auth.write_runtime_files(tempfile.mkdtemp(), {}, {}); assert False
    except auth.AuthError: pass
    print("auth ok")

def t_argv():
    """Pastikan argumen yang dirakit services diterima argparse engine asli (tanpa tersangkut 'unrecognized arguments')."""
    d = tempfile.mkdtemp(); auth.write_runtime_files(d, {"gcp_service_account": {"type": "service_account"}}, build_config())
    ctx = Ctx({}, build_config(), d, os.path.join(d, "config.json"))
    calls = [
        lambda: svc.run_tracker(ctx, "all", dt.date(2026,11,1), dt.date(2026,11,30), False, False, False, False, True),
        lambda: svc.run_orders(ctx, ["o1","o2"], ["p1"], True, False),
        lambda: svc.run_product(ctx, ["f1","f2"], False, True),
    ]
    for c in calls:
        r = c()
        assert "unrecognized" not in r.text and "usage:" not in r.text, r.text[-400:]
        assert not r.ok  # tanpa library/kredensial asli memang harus gagal, tapi bukan karena argumen
        print("  ->", r.error[:90])
    print("argv ok")

if __name__ == "__main__":
    t_runner(); t_auth(); t_argv()
