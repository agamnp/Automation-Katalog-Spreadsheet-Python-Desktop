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



# Konfigurasi logging
logging.basicConfig(
    filename="log_proses_katalog.txt",  # Nama file log
    level=logging.INFO,  # Level log: DEBUG, INFO, WARNING, ERROR, CRITICAL
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

from dotenv import load_dotenv
load_dotenv(dotenv_path="C:/Users/praaayogi/Documents/GitHub/Automation-Katalog-Spreadsheet-Python-Desktop/.env")

# Setup autentikasi
# Load environment variables dari .env file
def setup_google_sheets():
    load_dotenv()
    creds_path = os.getenv("GOOGLE_CREDS_PATH")
    if not creds_path or not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"File credentials.json tidak ditemukan di path: {creds_path}"
        )

    creds = Credentials.from_service_account_file(
        creds_path,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return gspread.authorize(creds)


# Autofill kolom
def autofill_column_general(
    sheet, col_letter, start_row, value_or_formula, mode="static", start_number=1
):
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
            range_name=autofill_range, values=values, value_input_option="USER_ENTERED"
        )
    except Exception as e:
        print(f"⚠️ Gagal mengisi kolom {col_letter}: {e}")
    time.sleep(min(2, num_rows * 0.02))  # 20ms per baris, max 2s


# Tambahkan rumus rekap
def add_formulas(sheet, retries=3):
    max_rows = len(sheet.col_values(3))
    max_rows = max(max_rows, 10)
    for attempt in range(retries):
        try:
            sheet.spreadsheet.values_batch_update(
                {
                    "valueInputOption": "USER_ENTERED",
                    "data": [
                        {
                            "range": f"{sheet.title}!G2",
                            "values": [[f"=COUNTA(C10:C{max_rows})"]],
                        },
                        {
                            "range": f"{sheet.title}!G3",
                            "values": [[f"=SUM(Y10:Y{max_rows})"]],
                        },
                        {
                            "range": f"{sheet.title}!G4",
                            "values": [[f"=AVERAGE(Y10:Y{max_rows})"]],
                        },
                        {
                            "range": f"{sheet.title}!J2",
                            "values": [[f"=COUNTA(Z10:Z{max_rows})"]],
                        },
                        {
                            "range": f"{sheet.title}!J3",
                            "values": [[f"=SUM(Z10:Z{max_rows})"]],
                        },
                        {
                            "range": f"{sheet.title}!J4",
                            "values": [[f"=SUM(AA10:AA{max_rows})"]],
                        },
                        {
                            "range": f"{sheet.title}!J5",
                            "values": [
                                [
                                    f'=AVERAGEIF(Z10:Z{max_rows}, ">0", AA10:AA{max_rows})'
                                ]
                            ],
                        },
                    ],
                }
            )
            break  # keluar loop jika berhasil
        except Exception as e:
            print(f"⚠️ Gagal menambahkan rumus rekap (Attempt {attempt + 1}): {e}")
            time.sleep(5)  # tunggu sebelum coba lagi
    else:
        print("❌ Gagal menambahkan rumus rekap setelah beberapa percobaan")


# tambahkan filter
def ensure_filter_and_freeze(sheet,logger=print):
    try:
        # Cari jumlah kolom dari baris header (baris 9)
        header_values = sheet.row_values(9)
        last_col_index = len(header_values)
        if last_col_index == 0:
            print("⚠️  Tidak ada header di baris 9. Filter dilewati.")
            return
        # Konversi index ke notasi kolom (misal 27 → 'AA')
        last_col_letter = gspread.utils.rowcol_to_a1(1, last_col_index).rstrip("1")
        filter_range = f"A9:{last_col_letter}9"
        # Set filter dinamis
        sheet.set_basic_filter(filter_range)
        logger(f"🔍 Filter     : {filter_range}")
        
        
        # Set freeze ke baris 9 dan kolom 10 (kolom J)
        sheet.freeze(rows=9, cols=10)
        logger(f"❄️  Freeze     : Baris 9, Kolom J")
        
    except Exception as e:
        print(f"⚠️  Gagal mengatur filter/freeze: {e}")


# nomor urut di spreadsheet
def rename_sheets_from_index(spreadsheet, sheet_order_start, zero_pad=3):
    """
    Rename semua sheet mulai dari urutan ke-N (bukan index ke-N).
    Penomoran juga dimulai dari N.
    """
    sheet_index_start = sheet_order_start - 1  # karena list index mulai dari 0
    sheets = spreadsheet.worksheets()[sheet_index_start:]
    for i, sheet in enumerate(sheets, start=sheet_order_start):
        sheet_number = f"{i:0{zero_pad}}"  # Contoh: 070, 071, ...
        rename_sheet_with_number(spreadsheet, sheet, sheet_number)


# Rename sheet individual
def rename_sheet_with_number(spreadsheet, sheet, sheet_number):
    old_title = sheet.title
    parts = old_title.split(".", 1)
    base_title = (
        parts[1].strip() if len(parts) > 1 and parts[0].isdigit() else old_title
    )
    base_title = base_title.replace(".", "")  # ganti titik jadi strip
    new_title = f"{sheet_number}.{base_title}"
    if new_title != old_title:
        try:
            sheet.update_title(new_title)
            msg = f"🔤 Rename     : '{old_title}' → '{new_title}'"
            print(msg)
            logging.info(msg)
            sheet = spreadsheet.worksheet(new_title)
        except Exception as e:
            print(f"⚠️ Gagal mengganti nama sheet '{old_title}': {e}")
    return sheet, new_title


# === Rename Named Range ===
def create_named_range_from_sheet_name(
    spreadsheet_id, sheet, header_row=9, col_start="A", col_end="Z"
):
    """
    Membuat named range berdasarkan nama sheet (dibersihkan jadi hanya huruf),
    dan range dari baris header sampai baris terakhir data aktual.
    """
    sheet_name = sheet.title
    clean_name = re.sub(r"[^a-zA-Z]", "", sheet_name)
    if not clean_name:
        print(f"⚠️ Nama sheet '{sheet_name}' kosong setelah dibersihkan. Skip.")
        return

    creds = Credentials.from_service_account_file(
        os.getenv("GOOGLE_CREDS_PATH"),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    service = build("sheets", "v4", credentials=creds)

    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    
    sheet_id = next(
        (
            s["properties"]["sheetId"]
            for s in spreadsheet["sheets"]
            if s["properties"]["title"] == sheet_name
        ),
        None,
    )
    if sheet_id is None:
        print(f"⚠️ Sheet ID untuk '{sheet_name}' tidak ditemukan.")
        return

    gc = setup_google_sheets()
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(sheet_name)

    col_index = ord(col_start.upper()) - 64
    col_values = ws.col_values(col_index)
    last_row = len(col_values)
    if last_row < header_row:
        print(f"⚠️ Sheet '{sheet_name}' tidak punya data setelah baris header.")
        return

    # Clean named range name
    clean_name = re.sub(r"[^a-zA-Z]", "", sheet_name)
    if not clean_name:
        print(f"⚠️ Nama sheet '{sheet_name}' kosong setelah dibersihkan.")
        return

    # Range tanpa sheet name untuk konversi
    a1_notation_only = f"{col_start}{header_row}:{col_end}{last_row}"
    range_a1 = f"'{sheet_name}'!{a1_notation_only}"
    grid_range = a1_range_to_grid_range(a1_notation_only, sheet_id)

    # Hapus existing range dengan nama yang sama
    existing_ranges = spreadsheet.get("namedRanges", [])
    existing_id = next(
        (r["namedRangeId"] for r in existing_ranges if r["name"] == clean_name), None
    )

    requests = []
    if existing_id:
        requests.append({"deleteNamedRange": {"namedRangeId": existing_id}})
    requests.append(
        {"addNamedRange": {"namedRange": {"name": clean_name, "range": grid_range}}}
    )

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests}
    ).execute()

    msg = f"🏷️ Named range '{clean_name}' Range → {range_a1}"
    print(msg)
    logging.info(msg)


def atur_border_dan_format_sheet(sheet, spreadsheet_id):
    try:
        from googleapiclient.discovery import build
        from google.oauth2.service_account import Credentials
        import os

        creds = Credentials.from_service_account_file(
            os.getenv("GOOGLE_CREDS_PATH"),
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        service = build("sheets", "v4", credentials=creds)

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
            return {
                "updateBorders": {
                    "range": range_obj,
                    "top": {
                        "style": "SOLID",
                        "color": {"red": 0, "green": 0, "blue": 0},
                    },
                    "bottom": {
                        "style": "SOLID",
                        "color": {"red": 0, "green": 0, "blue": 0},
                    },
                    "left": {
                        "style": "SOLID",
                        "color": {"red": 0, "green": 0, "blue": 0},
                    },
                    "right": {
                        "style": "SOLID",
                        "color": {"red": 0, "green": 0, "blue": 0},
                    },
                    "innerHorizontal": {
                        "style": "SOLID",
                        "color": {"red": 0, "green": 0, "blue": 0},
                    },
                    "innerVertical": {
                        "style": "SOLID",
                        "color": {"red": 0, "green": 0, "blue": 0},
                    },
                }
            }

        # Format dan border
        requests = []

        # Borders
        requests.append(full_border(range_obj(0, 5, 5, 7)))  # F1:G4
        requests.append(full_border(range_obj(0, 5, 8, 10)))  # I1:J5
        requests.append(full_border(range_obj(8, max_rows, 0, 28)))  # A9:AB{maxRows}

        # Format alignment & font
        alignment_zones = [
            # range, horizontal, vertical, bold
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
            requests.append(
                {
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
                }
            )

        # Format angka (Rp)
        currency_ranges = [
            range_obj(3, 4, 6, 7),  # G4
            range_obj(2, 4, 6, 7),  # G3:G4
            range_obj(3, 4, 9, 10),  # J4
        ]
        for cr in currency_ranges:
            requests.append(
                {
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
                }
            )

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()

        print("✅ Format border dan alignment berhasil diterapkan.")
        logging.info("✅ Format border dan alignment berhasil diterapkan.")
    except Exception as e:
        print(f"⚠️ Gagal mengatur border/alignment: {e}")
        logging.error(f"Gagal mengatur border/alignment: {e}")


# Fungsi utama
def main_tampilan_sheet(logger=print):
    try:
        logger("📄 Menjalankan tampilan sheet...")
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
        gc = setup_google_sheets()
        sh = gc.open_by_key(SPREADSHEET_ID)
        worksheets = sh.worksheets()

        
        logger(f"🔄 Mulai proses semua sheet...\n")
        logger(f"📂 Nama Spreadsheet: {sh.title}")

        sheet_number = SHEET_MULAI if SHEET_MULAI > 0 else 1

        for i, sheet in enumerate(worksheets[START_SHEET_INDEX:], start=START_SHEET_INDEX):

            if get_stop_requested():
                logger("⏹️ Proses dihentikan oleh pengguna.")
                break

            if sheet.title in excluded_sheets:
                logger(msg = f"➡️ Sheet '{sheet.title}' dilewati.")
                continue

            sheet, new_title = rename_sheet_with_number(sh, sheet, sheet_number)
            sheet_number += 1

            
            logger(f"✅ Memproses Sheet: {new_title}")

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
            
            logger(f"✅ Kolom AA dihitung dari Y*Z")
            

            ensure_filter_and_freeze(sheet)
            add_formulas(sheet)
            atur_border_dan_format_sheet(sheet, spreadsheet_id=sh.id)

            create_named_range_from_sheet_name(
                spreadsheet_id=sh.id, sheet=sheet, header_row=10, col_start="J", col_end="J"
            )

            
            logger(f"🎯 Proses sheet '{new_title}' selesai.")
            logger("")

        print("🎉 Semua sheet selesai diproses!")
        logger("🎉 Semua sheet selesai diproses!")

    except Exception as e:
        logger(f"❌ Terjadi error saat proses utama: {e}")



# if __name__ == '__main__':
#    main()
