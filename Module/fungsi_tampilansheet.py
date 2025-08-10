import gspread
from google.oauth2.service_account import Credentials
import time
from dotenv import load_dotenv
import os
from googleapiclient.discovery import build
from gspread.utils import a1_range_to_grid_range, rowcol_to_a1
import re
import logging
from globals import get_stop_requested, set_stop_requested


# Load environment variables
load_dotenv(dotenv_path="C:/Users/praaayogi/Documents/GitHub/Automation-Katalog-Spreadsheet-Python-Desktop/.env")

def setup_google_sheets():
    """Setup Google Sheets dengan credentials yang diperbaharui"""
    load_dotenv()
    creds_path = os.getenv("GOOGLE_CREDS_PATH")
    if not creds_path or not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"File credentials.json tidak ditemukan di path: {creds_path}"
        )

    # Gunakan google.oauth2.service_account.Credentials (sudah modern)
    creds = Credentials.from_service_account_file(
        creds_path,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    
    # Pastikan menggunakan gspread terbaru
    gc = gspread.authorize(creds)
    return gc

def get_sheets_service():
    """Get Google Sheets API service dengan credentials modern"""
    creds_path = os.getenv("GOOGLE_CREDS_PATH")
    creds = Credentials.from_service_account_file(
        creds_path,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)

def autofill_column_general(
    sheet, col_letter, start_row, value_or_formula, mode="static", start_number=1
):
    """Autofill kolom dengan berbagai mode"""
    total_rows = len(sheet.col_values(3))  # Kolom C sebagai acuan
    last_row = max(total_rows, start_row)
    num_rows = last_row - start_row + 1
    autofill_range = f"{col_letter}{start_row}:{col_letter}{last_row}"
    
    if mode == "number":
        values = [[start_number + i] for i in range(num_rows)]
    elif mode == "dynamic":
        values = [
            [value_or_formula.format(row=row)] for row in range(start_row, last_row + 1)
        ]
    else:
        values = [[value_or_formula]] * num_rows
    
    try:
        sheet.update(
            range_name=autofill_range, 
            values=values, 
            value_input_option="USER_ENTERED"
        )
        logging.info(f"✅ Kolom {col_letter} berhasil diisi ({num_rows} baris)")
    except Exception as e:
        error_msg = f"⚠️ Gagal mengisi kolom {col_letter}: {e}"
        print(error_msg)
        logging.error(error_msg)
    
    time.sleep(min(2, num_rows * 0.02))  # Rate limiting

def add_formulas(sheet, retries=3):
    """Tambahkan rumus rekap dengan retry mechanism"""
    max_rows = len(sheet.col_values(3))
    max_rows = max(max_rows, 10)
    
    formula_data = [
        {"range": f"{sheet.title}!G2", "values": [[f"=COUNTA(C10:C{max_rows})"]]},
        {"range": f"{sheet.title}!G3", "values": [[f"=SUM(Y10:Y{max_rows})"]]},
        {"range": f"{sheet.title}!G4", "values": [[f"=AVERAGE(Y10:Y{max_rows})"]]},
        {"range": f"{sheet.title}!J2", "values": [[f"=COUNTA(Z10:Z{max_rows})"]]},
        {"range": f"{sheet.title}!J3", "values": [[f"=SUM(Z10:Z{max_rows})"]]},
        {"range": f"{sheet.title}!J4", "values": [[f"=SUM(AA10:AA{max_rows})"]]},
        {"range": f"{sheet.title}!J5", "values": [[f'=AVERAGEIF(Z10:Z{max_rows}, ">0", AA10:AA{max_rows})']]},
    ]
    
    for attempt in range(retries):
        try:
            sheet.spreadsheet.values_batch_update({
                "valueInputOption": "USER_ENTERED",
                "data": formula_data
            })
            logging.info("✅ Rumus rekap berhasil ditambahkan")
            break
        except Exception as e:
            error_msg = f"⚠️ Gagal menambahkan rumus rekap (Attempt {attempt + 1}): {e}"
            print(error_msg)
            logging.warning(error_msg)
            if attempt < retries - 1:
                time.sleep(5)
    else:
        logging.error("❌ Gagal menambahkan rumus rekap setelah beberapa percobaan")

def clear_rows_after_table(sheet, data_col="C", start_row=10, logger=print):
    """Hapus baris kosong setelah data terakhir di tabel"""
    try:
        # Cari baris terakhir yang ada data
        col_values = sheet.col_values(ord(data_col.upper()) - 64)  # Convert C->3
        last_data_row = len(col_values)
        
        # Cek total baris di sheet
        sheet_rows = sheet.row_count
        
        if last_data_row < sheet_rows:
            # Ada baris kosong yang perlu dihapus
            rows_to_delete = sheet_rows - last_data_row
            
            # Hapus baris kosong (dari baris terakhir data + 1)
            sheet.delete_rows(last_data_row + 1, sheet_rows)
            
            msg = f"🗑️ Dihapus {rows_to_delete} baris kosong setelah baris {last_data_row}"
            logger(msg)
            logging.info(msg)
        else:
            logger("✅ Tidak ada baris kosong untuk dihapus")
            
    except Exception as e:
        error_msg = f"⚠️ Gagal menghapus baris kosong: {e}"
        logger(error_msg)
        logging.error(error_msg)

def ensure_filter_and_freeze(sheet, logger=print):
    """Setup filter dan freeze panes"""
    try:
        header_values = sheet.row_values(9)
        last_col_index = len(header_values)
        if last_col_index == 0:
            logger("⚠️ Tidak ada header di baris 9. Filter dilewati.")
            return
        
        last_col_letter = gspread.utils.rowcol_to_a1(1, last_col_index).rstrip("1")
        filter_range = f"A9:{last_col_letter}9"
        
        sheet.set_basic_filter(filter_range)
        logger(f"🔍 Filter: {filter_range}")
        
        sheet.freeze(rows=9, cols=10)
        logger(f"❄️ Freeze: Baris 9, Kolom J")
        
        logging.info(f"✅ Filter dan freeze berhasil diatur: {filter_range}")
        
    except Exception as e:
        error_msg = f"⚠️ Gagal mengatur filter/freeze: {e}"
        print(error_msg)
        logging.error(error_msg)

def rename_sheets_from_index(spreadsheet, sheet_order_start, zero_pad=3):
    """Rename semua sheet mulai dari urutan tertentu"""
    sheet_index_start = sheet_order_start - 1
    sheets = spreadsheet.worksheets()[sheet_index_start:]
    
    for i, sheet in enumerate(sheets, start=sheet_order_start):
        sheet_number = f"{i:0{zero_pad}}"
        rename_sheet_with_number(spreadsheet, sheet, sheet_number)

def rename_sheet_with_number(spreadsheet, sheet, sheet_number):
    """Rename sheet individual dengan numbering"""
    old_title = sheet.title
    parts = old_title.split(".", 1)
    base_title = (
        parts[1].strip() if len(parts) > 1 and parts[0].isdigit() else old_title
    )
    base_title = base_title.replace(".", "")
    new_title = f"{sheet_number}.{base_title}"
    
    if new_title != old_title:
        try:
            sheet.update_title(new_title)
            msg = f"🔤 Rename: '{old_title}' → '{new_title}'"
            
            logging.info(msg)
            sheet = spreadsheet.worksheet(new_title)
        except Exception as e:
            error_msg = f"⚠️ Gagal mengganti nama sheet '{old_title}': {e}"
            print(error_msg)
            logging.error(error_msg)
    
    return sheet, new_title

def create_named_range_from_sheet_name(
    spreadsheet_id, sheet, header_row=9, col_start="A", col_end="Z"
):
    """Membuat named range dari nama sheet"""
    sheet_name = sheet.title
    clean_name = re.sub(r"[^a-zA-Z]", "", sheet_name)
    if not clean_name:
        logging.warning(f"⚠️ Nama sheet '{sheet_name}' kosong setelah dibersihkan. Skip.")
        return

    try:
        service = get_sheets_service()
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        
        sheet_id = next(
            (s["properties"]["sheetId"] for s in spreadsheet["sheets"] 
             if s["properties"]["title"] == sheet_name), None
        )
        
        if sheet_id is None:
            logging.error(f"⚠️ Sheet ID untuk '{sheet_name}' tidak ditemukan.")
            return

        # Get actual data range
        gc = setup_google_sheets()
        sh = gc.open_by_key(spreadsheet_id)
        ws = sh.worksheet(sheet_name)
        
        col_index = ord(col_start.upper()) - 64
        col_values = ws.col_values(col_index)
        last_row = len(col_values)
        
        if last_row < header_row:
            logging.warning(f"⚠️ Sheet '{sheet_name}' tidak punya data setelah baris header.")
            return

        # Create range
        a1_notation_only = f"{col_start}{header_row}:{col_end}{last_row}"
        range_a1 = f"'{sheet_name}'!{a1_notation_only}"
        grid_range = a1_range_to_grid_range(a1_notation_only, sheet_id)

        # Delete existing range if exists
        existing_ranges = spreadsheet.get("namedRanges", [])
        existing_id = next(
            (r["namedRangeId"] for r in existing_ranges if r["name"] == clean_name), None
        )

        requests = []
        if existing_id:
            requests.append({"deleteNamedRange": {"namedRangeId": existing_id}})
        
        requests.append({
            "addNamedRange": {
                "namedRange": {
                    "name": clean_name,
                    "range": grid_range
                }
            }
        })

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, 
            body={"requests": requests}
        ).execute()

        msg = f"🏷️ Named range '{clean_name}' Range → {range_a1}"
        
        logging.info(msg)
        
    except Exception as e:
        error_msg = f"⚠️ Gagal membuat named range untuk '{sheet_name}': {e}"
        print(error_msg)
        logging.error(error_msg)

def atur_border_dan_format_sheet(sheet, spreadsheet_id):
    """Atur border dan formatting sheet"""
    try:
        service = get_sheets_service()
        sheet_id = sheet._properties["sheetId"]
        max_rows = max(len(sheet.col_values(3)), 10)

        def range_obj(start_row, end_row, start_col, end_col):
            return {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col,
            }

        def full_border(range_obj):
            border_style = {
                "style": "SOLID",
                "color": {"red": 0, "green": 0, "blue": 0},
            }
            return {
                "updateBorders": {
                    "range": range_obj,
                    "top": border_style,
                    "bottom": border_style,
                    "left": border_style,
                    "right": border_style,
                    "innerHorizontal": border_style,
                    "innerVertical": border_style,
                }
            }

        requests = []

        # Add borders
        requests.append(full_border(range_obj(0, 5, 5, 7)))  # F1:G4
        requests.append(full_border(range_obj(0, 5, 8, 10)))  # I1:J5
        requests.append(full_border(range_obj(8, max_rows, 0, 28)))  # A9:AB{maxRows}

        # Format alignment & font
        alignment_zones = [
            (range_obj(0, 5, 5, 10), "CENTER", "MIDDLE", False),  # F1:J5
            (range_obj(9, max_rows, 0, 2), "CENTER", "MIDDLE", False),  # A10:B
            (range_obj(9, max_rows, 2, 5), "LEFT", "MIDDLE", False),  # C10:E
            (range_obj(9, max_rows, 5, 7), "CENTER", "MIDDLE", False),  # F10:G
            (range_obj(9, max_rows, 7, 9), "LEFT", "MIDDLE", False),  # H10:I
            (range_obj(9, max_rows, 9, 10), "CENTER", "MIDDLE", False),  # J10
            (range_obj(9, max_rows, 10, 24), "LEFT", "MIDDLE", False),  # K10:X
            (range_obj(9, max_rows, 24, 28), "CENTER", "MIDDLE", False),  # Y10:AB
        ]

        for rng, h_align, v_align, bold in alignment_zones:
            requests.append({
                "repeatCell": {
                    "range": rng,
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": h_align,
                            "verticalAlignment": v_align,
                            "textFormat": {"bold": bold},
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat)",
                }
            })

        # Format currency
        currency_ranges = [
            range_obj(3, 4, 6, 7),  # G4
            range_obj(2, 4, 6, 7),  # G3:G4
            range_obj(3, 4, 9, 10),  # J4
        ]
        
        for cr in currency_ranges:
            requests.append({
                "repeatCell": {
                    "range": cr,
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "CURRENCY",
                                "pattern": "[$Rp-421] #,##0",
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            })

        # Execute all requests
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, 
            body={"requests": requests}
        ).execute()

        
        logging.info("✅ Format border dan alignment berhasil diterapkan.")
        
    except Exception as e:
        error_msg = f"⚠️ Gagal mengatur border/alignment: {e}"
        print(error_msg)
        logging.error(error_msg)

def main_tampilan_sheet(logger=print):
    """Fungsi utama untuk memproses tampilan sheet"""
    try:
        logger("📄 Menjalankan tampilan sheet...")
        
        # Load environment variables
        SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "")
        SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
        if not SPREADSHEET_ID:
            logger("❌ SPREADSHEET_ID tidak ditemukan.")
            return

        excluded_sheets = os.getenv("EXCLUDED_SHEETS", "")
        excluded_sheets = [s.strip() for s in excluded_sheets.split(",") if s.strip()]
        
        SHEET_MULAI = int(os.getenv("SHEET_MULAI", "1"))
        START_SHEET_INDEX = SHEET_MULAI + 2
        START_ROW = 10
        
        # Setup Google Sheets
        gc = setup_google_sheets()
        sh = gc.open_by_key(SPREADSHEET_ID)
        worksheets = sh.worksheets()

        logger(f"🔄 Mulai proses semua sheet...\n")
        logger(f"📂 Nama Spreadsheet: {sh.title}")
        logging.info(f"Memulai proses untuk spreadsheet: {sh.title}")

        sheet_number = SHEET_MULAI if SHEET_MULAI > 0 else 1

        for i, sheet in enumerate(worksheets[START_SHEET_INDEX:], start=START_SHEET_INDEX):
            
            if get_stop_requested():
                logger("⏹️ Proses dihentikan oleh pengguna.")
                logging.info("Proses dihentikan oleh pengguna")
                break

            if sheet.title in excluded_sheets:
                logger(f"➡️ Sheet '{sheet.title}' dilewati.")
                continue

            # Rename sheet
            sheet, new_title = rename_sheet_with_number(sh, sheet, sheet_number)
            sheet_number += 1

            logger(f"✅ Memproses Sheet: {new_title}")

            # Process columns
            autofill_column_general(sheet, "A", START_ROW, "", mode="number")
            logger("✅ Nomor urut di kolom A selesai")

            autofill_column_general(
                sheet,
                "B",
                START_ROW,
                '=HYPERLINK("https://mocostore.moco.co.id/catalog/"&AB{row};"Klik Disini")',
                mode="dynamic",
            )
            logger("✅ Kolom B diisi hyperlink")

            autofill_column_general(
                sheet, "AA", START_ROW, "=Y{row}*Z{row}", mode="dynamic"
            )
            logger("✅ Kolom AA dihitung dari Y*Z")

            # Apply formatting and features
            clear_rows_after_table(sheet, data_col="C", logger=logger)
            ensure_filter_and_freeze(sheet, logger)
            add_formulas(sheet)
            atur_border_dan_format_sheet(sheet, spreadsheet_id=sh.id)

            create_named_range_from_sheet_name(
                spreadsheet_id=sh.id, 
                sheet=sheet, 
                header_row=10, 
                col_start="J", 
                col_end="J"
            )

            logger(f"🎯 Proses sheet '{new_title}' selesai.")
            logger("")

        logger("🎉 Semua sheet selesai diproses!")
        logging.info("🎉 Semua sheet selesai diproses!")

    except Exception as e:
        error_msg = f"❌ Terjadi error saat proses utama: {e}"
        logger(error_msg)
        logging.error(error_msg)

# if __name__ == '__main__':
#     main_tampilan_sheet()