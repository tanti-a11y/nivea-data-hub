"""Daftar file di folder Drive (untuk dipilih pengguna)."""

GSHEET = "application/vnd.google-apps.spreadsheet"


def list_folder(drive, folder_id, name_contains=None, exts=(".csv", ".xlsx", ".xls")):
    q = f"'{folder_id}' in parents and trashed = false"
    if name_contains:
        q += f" and name contains '{name_contains}'"
    files, token = [], None
    while True:
        resp = drive.files().list(
            q=q, pageToken=token, pageSize=1000,
            fields="nextPageToken, files(id, name, mimeType, modifiedTime)",
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        files += resp.get("files", [])
        token = resp.get("nextPageToken")
        if not token:
            break
    ok = [f for f in files if f["mimeType"] == GSHEET or f["name"].lower().endswith(exts)]
    return sorted(ok, key=lambda f: f["modifiedTime"], reverse=True)


def label(f):
    return f"{f['name']}  ·  diubah {f['modifiedTime'][:16].replace('T', ' ')} UTC"
