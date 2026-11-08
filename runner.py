"""Menjalankan engine lama (yang aslinya CLI) dari dalam Streamlit tanpa mengubah logikanya.

- stdout/stderr ditangkap jadi teks laporan
- input() dimatikan (engine tidak boleh menunggu ketikan)
- sys.exit(pesan) diubah jadi pesan error biasa
"""
import builtins
import contextlib
import io
import sys
import threading
import traceback
from dataclasses import dataclass, field

_LOCK = threading.Lock()


@dataclass
class RunResult:
    ok: bool
    text: str
    error: str = ""
    extra: dict = field(default_factory=dict)


class NeedsInput(Exception):
    pass


def _no_input(prompt=""):
    raise NeedsInput(f"Engine meminta input interaktif: {prompt!r}")


def run_main(main_fn, argv=None, answers=None, patches=None):
    """Jalankan main_fn() dengan sys.argv=argv. answers = daftar jawaban untuk input() (opsional).
    patches = {(objek, 'nama'): nilai_pengganti} untuk mengganti fungsi di engine selama dijalankan."""
    answers = list(answers or [])
    patches = patches or {}

    def fake_input(prompt=""):
        if answers:
            return answers.pop(0)
        return _no_input(prompt)

    out = io.StringIO()
    old_argv = sys.argv[:]
    olds = [(o, n, getattr(o, n)) for (o, n) in patches]
    ok, err = True, ""
    with _LOCK:
        try:
            sys.argv = ["engine"] + list(argv or [])
            for (o, n), v in patches.items():
                setattr(o, n, v)
            old_input = builtins.input
            builtins.input = fake_input
            try:
                with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
                    main_fn()
            finally:
                builtins.input = old_input
        except SystemExit as e:
            if e.code not in (None, 0):
                ok = False
                err = str(e.code) if not isinstance(e.code, int) else f"Proses berhenti (kode {e.code})"
        except Exception as e:  # noqa
            ok = False
            err = f"{type(e).__name__}: {e}"
            out.write("\n" + traceback.format_exc())
        finally:
            sys.argv = old_argv
            for o, n, v in olds:
                setattr(o, n, v)
    return RunResult(ok=ok, text=out.getvalue(), error=err)
