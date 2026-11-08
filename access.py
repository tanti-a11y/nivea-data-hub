"""Aturan siapa yang boleh masuk (murni logika, mudah dites)."""


def is_allowed(email, verified, domain="warnaemasindonesia.com", extra=()):
    """Email harus terverifikasi Google dan berakhiran @domain (atau ada di daftar tambahan)."""
    email = (email or "").strip().lower()
    if not email or verified is False:
        return False
    if email in {e.strip().lower() for e in extra or ()}:
        return True
    return email.endswith("@" + domain.strip().lower().lstrip("@"))
