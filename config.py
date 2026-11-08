"""Pengaturan bawaan (ID folder / spreadsheet). Bisa ditimpa lewat bagian [config] di Streamlit secrets."""
import copy

DEFAULT_CONFIG = {
    # --- Tracker SV / LS (nivea_tracker.py) ---
    "tracker_spreadsheet_id": "1avlQXBatlbgT8xjYi8s_h5S4oiIwBN1A-6TnlySA1QQ",
    "tracker_sheet": "SV",
    "header_row": 2,
    "daily_folder_id": "1W1mWOI0zPHmEsotWYYGwyX3eT1fwbaXt",
    "ls_sheet": "LS",
    "ls_header_row": 1,
    "ls_folder_id": "1guL2f4b2IU6LeLdrDSeeFga5MfGgCuoM",
    "product_bank_file_id": "1N7CKyWKEg1DKKgJHNWMvK_PXfHpAGw-9",
    "product_bank_sheet": "Product Bank",
    "creator_tiering_file_id": "1vWTJul4zyIhaFxuGRUvmDeCr5jGR9lGW",
    "creator_tiering_sheet": "Creator Tiering",
    "gmv_column": "Video-attributed GMV (Rp)",
    "default_tier": "OPEN PLAN",
    "sv_column_map": {},
    "ls_column_map": {},
    # --- Order Affiliate + Patner & Product Performance (affiliate_merge.py) ---
    "affiliate_order_folder_id": "1iUzLAYExrm4dNJYH5Y4iU18O2ZAQHKFc",
    "affiliate_partner_folder_id": "1vhvcJX8yaSXUoADQeOhb6QI2_xYWVoau",
    "merge_output_spreadsheet_id": "1hrwzM1pnVZ5l-yYQ3cKnS480XJJUKW5X1bBB9I-T3Hg",
    "product_folder_id": "1uJc9bov06GNPXk9-jXMF-Grl93RD14C1",
}

# Penjelasan singkat (bahasa awam) untuk halaman Pengaturan
LABELS = {
    "tracker_spreadsheet_id": "Spreadsheet tracker (tab SV & LS)",
    "daily_folder_id": "Folder Drive export harian SV (Video)",
    "ls_folder_id": "Folder Drive export harian LS (Live)",
    "product_bank_file_id": "File Product Bank (CAT/PRO)",
    "creator_tiering_file_id": "File Creator Tiering",
    "affiliate_order_folder_id": "Folder Drive Affiliate_Order",
    "affiliate_partner_folder_id": "Folder Drive Affiliate_Patner",
    "merge_output_spreadsheet_id": "Spreadsheet output Order & Product Performance",
    "product_folder_id": "Folder Drive Product Performance harian",
}


def build_config(overrides=None):
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    for k, v in (overrides or {}).items():
        cfg[k] = v
    return cfg
