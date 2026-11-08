import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from hub.access import is_allowed as a
assert a("tanti@warnaemasindonesia.com", True)
assert a("Tanti@WarnaEmasIndonesia.com", None)
assert not a("x@gmail.com", True)
assert not a("x@warnaemasindonesia.com.evil.com", True)      # bukan domain yang sama
assert not a("x@evilwarnaemasindonesia.com", True)            # subdomain/akhiran palsu
assert not a("x@warnaemasindonesia.com", False)               # email belum terverifikasi
assert not a("", True) and not a(None, True)
assert a("tamu@gmail.com", True, extra=["tamu@gmail.com"])
print("access ok")
