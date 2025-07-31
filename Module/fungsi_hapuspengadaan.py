import pandas as pd
import time
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
from tkinter import filedialog

# Fungsi bantu untuk aman konversi string
def safe_str(val):
    if pd.notna(val):
        return str(val).strip().replace("-", "").replace(" ", "")
    return ""

# Setup credentials Google Sheets API
def setup_sheets_api():
    load_dotenv()
    creds_path = os.getenv("GOOGLE_CREDS_PATH")
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    return build('sheets', 'v4', credentials=creds)

# Fungsi batch hapus
def batch_delete_rows(service, spreadsheet_id, sheet_id, rows_to_delete, logger=print):
    requests = []
    for row_idx in sorted(rows_to_delete, reverse=True):
        requests.append({
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": row_idx - 1,
                    "endIndex": row_idx
                }
            }
        })

    if not requests:
        logger("⚠️ Tidak ada baris untuk dihapus.")
        return 0

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()
        logger(f"✅ Berhasil hapus {len(rows_to_delete)} baris.")
        return len(rows_to_delete)
    except Exception as e:
        logger(f"❌ Gagal hapus baris: {e}")
        return 0

# Fungsi utama
def main_hapus_pengadaan(logger=print):
    logger("📤 Silakan pilih file Excel (.xlsx) yang berisi data penghapusan...")
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        logger("❌ Tidak ada file dipilih. Proses dibatalkan.")
        return

    df = pd.read_excel(file_path)
    logger(f"✅ File dibaca: {file_path}")
    logger(f"🔍 Jumlah data: {len(df)}")

    service = setup_sheets_api()
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if not spreadsheet_id:
        logger("❌ SPREADSHEET_ID tidak ditemukan di environment.")
        return

    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])
    excluded_sheets = os.getenv("EXCLUDED_SHEETS", "")
    excluded_sheets = [s.strip() for s in excluded_sheets.split(",") if s.strip()]
    total_deleted = 0

    for sheet in sheets:
        title = sheet['properties']['title']
        sheet_id = sheet['properties']['sheetId']

        if title in excluded_sheets:
            logger(f"➡️ Sheet '{title}' dilewati karena termasuk excluded_sheets.")
            continue

        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{title}'!A1:AB9000"
        ).execute()

        all_data = result.get('values', [])
        if len(all_data) < 10:
            logger(f"⚠️ Sheet '{title}' tidak memiliki cukup baris. Dilewati.")
            continue

        headers = all_data[8]
        rows = all_data[9:]
        header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
        required_cols = [
           # "judul*",
            "isbn cetak", "isbn elektronik*", "uuid"]
        if not all(col in header_map for col in required_cols):
            logger(f"⚠️ Sheet '{title}' tidak memiliki semua kolom penting. Dilewati.")
            continue

     #   idx_judul = header_map["judul*"]
        idx_isbn = header_map["isbn cetak"]
        idx_isbn_e = header_map["isbn elektronik*"]
        idx_uuid = header_map["uuid"]

        rows_to_delete = []

        for index, row_excel in df.iterrows():
            uuid = safe_str(row_excel.get("UUID"))
        #    judul = safe_str(row_excel.get("Judul*")).lower()
            isbn = safe_str(row_excel.get("ISBN Cetak")).lower()
            isbn_e = safe_str(row_excel.get("ISBN Elektronik*")).lower()

            for i, row_sheet in enumerate(rows):
                try:
             #       row_judul = safe_str(row_sheet[idx_judul]).lower()
                    row_isbn = safe_str(row_sheet[idx_isbn]).lower()
                    row_isbn_e = safe_str(row_sheet[idx_isbn_e]).lower()
                    row_uuid = safe_str(row_sheet[idx_uuid])
                except IndexError:
                    continue

                if (
                    (uuid and uuid == row_uuid)
                #    or (judul and judul == row_judul)
                    or (isbn and isbn == row_isbn)
                    or (isbn_e and isbn_e == row_isbn_e)
                ):
                   # logger(f"🔍 Match: UUID='{uuid}', Judul='{judul}' di sheet '{title}'")
                    logger(f"🔍 Match: UUID='{uuid}' di sheet '{title}'")
                    real_row_number = i + 10
                    rows_to_delete.append(real_row_number)
                    break  # keluar dari loop sheet, lanjut ke baris Excel berikutnya

        if not rows_to_delete:
            logger(f"⚠️ Tidak ditemukan baris cocok di sheet '{title}'.")

        deleted_count = batch_delete_rows(service, spreadsheet_id, sheet_id, rows_to_delete, logger)
        total_deleted += deleted_count
        time.sleep(1)

    logger(f"🎯 Total baris dihapus di semua sheet: {total_deleted}")

if __name__ == "__main__":
    main_hapus_pengadaan()
