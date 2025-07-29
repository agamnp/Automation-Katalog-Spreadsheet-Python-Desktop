import pandas as pd
from tkinter import filedialog
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import os
import time


def setup_google_sheets():
    load_dotenv()
    creds_path = os.getenv("GOOGLE_CREDS_PATH")
    creds = Credentials.from_service_account_file(
    creds_path,
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
)
    return gspread.authorize(creds)


def main_hapus_pengadaan(logger=print):
    logger("📤 Silakan pilih file Excel (.xlsx) yang berisi data penghapusan...")
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        logger("❌ Tidak ada file dipilih. Proses dibatalkan.")
        return

    df = pd.read_excel(file_path)
    logger(f"✅ File dibaca: {file_path}")
    logger(f"🔍 Jumlah data: {len(df)}")

    gc = setup_google_sheets()
    spreadsheet_name = "automasi katalog"
    sh = gc.open(spreadsheet_name)
    sheet = sh.worksheet("Form Pengadaan")

    all_data = sheet.get_all_values()
    headers = all_data[0]
    rows = all_data[1:]

    deleted_count = 0

    for index, row_excel in df.iterrows():
        uuid = str(row_excel.get("UUID", "")).strip()
        judul = str(row_excel.get("Judul", "")).strip().lower()
        isbn = str(row_excel.get("ISBN Cetak", "")).strip()
        isbn_e = str(row_excel.get("ISBN Elektronik", "")).strip()

        for i, row_sheet in enumerate(rows):
            row_judul = row_sheet[1].strip().lower()
            row_isbn = row_sheet[2].strip()
            row_isbn_e = row_sheet[3].strip()
            row_uuid = row_sheet[4].strip()

            if (
                uuid and uuid == row_uuid
                or judul and judul == row_judul
                or isbn and isbn == row_isbn
                or isbn_e and isbn_e == row_isbn_e
            ):
                real_row_number = i + 2  # offset karena header di baris 1
                sheet.delete_row(real_row_number)
                logger(f"🗑️ Baris {real_row_number} dihapus: {row_sheet[1]}")
                deleted_count += 1
                time.sleep(1)
                break  # hanya hapus satu baris yang cocok

    logger(f"🎯 Total baris dihapus: {deleted_count}")
