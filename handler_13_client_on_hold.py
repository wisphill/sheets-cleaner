import os
from typing import List, Set, Optional
import openpyxl
import sys

# Khai báo tiêu đề cột chuẩn (so sánh nguyên văn)
CLIENT_STATUS_SHEET_NAME = "ClientStatus"
EXACT_CLIENT_ID_HEADER = "ClientId*"
EXACT_STATUS_HEADER = "Status*"
EXACT_EFFECTIVE_DATE_HEADER = "EffectiveDate*"
EXACT_AC_NUMBER_HEADER = "AC Number"  # Hoặc "ACNumber" tùy tiêu đề trong file Client_on_hold.xlsx

ON_HOLD_FILE_PATH = "./data/Clients on hold as of 28.9.2026.xlsx"

# ==============================================================================
# BƯỚC 1: Đọc danh sách AC Number từ file Client_on_hold.xlsx
# ==============================================================================
def get_on_hold_client_ids(
    file_path: str = "Client_on_hold.xlsx",
) -> Set[str]:
    """1.

    Refer tới file 'Client_on_hold.xlsx' và trích xuất tất cả các AC Number
    (Clients đang On Hold).
    """
    if not os.path.exists(file_path):
        print(
            f"❌ File '{file_path}' không tồn tại. Không thể lấy danh sách Client On Hold."
        )
        return set()

    wb_hold = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb_hold.active  # Lấy sheet đầu tiên

    ac_number_col_idx = None
    on_hold_ids = set()

    # Tìm vị trí cột chứa AC Number
    for col in range(1, sheet.max_column + 1):
        cell_val = sheet.cell(row=1, column=col).value
        if cell_val is not None and EXACT_AC_NUMBER_HEADER in str(cell_val):
            ac_number_col_idx = col
            break

    # Nếu không tìm thấy bằng tiêu đề chính xác, exit
    if not ac_number_col_idx:
        print(
            f"❌ Không tìm thấy cột '{EXACT_AC_NUMBER_HEADER}' trong file '{file_path}'. Dừng chương trình."
        )
        sys.exit(1)  # Thoát chương trình kèm mã lỗi status code = 1

    # Đọc dữ liệu từ hàng 2 trở đi
    for row in range(2, sheet.max_row + 1):
        val = sheet.cell(row=row, column=ac_number_col_idx).value
        if val is not None and str(val).strip() != "":
            on_hold_ids.add(str(val).strip())

    print(
        f"📋 Đã tải {len(on_hold_ids)} AC Number (On Hold) từ file '{file_path}'."
    )
    return on_hold_ids


from typing import Set
import openpyxl

# ==============================================================================
# BƯỚC 2: Cập nhật Status* thành "On Hold"
# ==============================================================================
def update_on_hold_clients_status(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    client_id_col: int,
    status_col: int,
    on_hold_ids: Set[str],
) -> Set[str]:
    """Search trong 'ClientId*' bằng AC Number (dùng Partial Match).

    Update 'Status*' thành 'On Hold'. Log ra danh sách AC Number không tìm thấy.
    """
    matched_on_hold_ids = set()

    for row in range(3, sheet.max_row + 1):  # Bỏ qua Hàng 2 Guideline
        client_id_val = sheet.cell(row=row, column=client_id_col).value

        if client_id_val is not None:
            client_id_str = str(client_id_val).strip()

            # Partial match: Kiểm tra xem có AC Number nào nằm trong ClientId* hay không
            for ac_num in on_hold_ids:
                if ac_num in client_id_str:
                    # Đánh dấu tìm thấy
                    matched_on_hold_ids.add(ac_num)

                    # Update Status* = "On Hold"
                    sheet.cell(row=row, column=status_col).value = "On Hold"
                    break

    # Ghi log các AC Number không tìm thấy trong sheet ClientStatus
    unmatched_ids = on_hold_ids - matched_on_hold_ids
    if unmatched_ids:
        print(
            f"⚠️ [LOG] {len(unmatched_ids)} AC Number không tìm thấy trong sheet 'ClientStatus': {sorted(list(unmatched_ids))}"
        )

    return matched_on_hold_ids

# ==============================================================================
# BƯỚC MỚI: Cập nhật EffectiveDate* cho các dòng "On Hold"
# ==============================================================================
def update_effective_date_for_on_hold_clients(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    status_col: int,
    effective_date_col: int,
    effective_date_value: str = "TODAY",
) -> int:
    """Duyệt lại tất cả các dòng có Status* = 'On Hold' và cập nhật
    EffectiveDate*."""
    updated_count = 0

    for row in range(3, sheet.max_row + 1):  # Bỏ qua Hàng 2 Guideline
        status_val = sheet.cell(row=row, column=status_col).value

        # Kiểm tra nếu Status* đúng bằng "On Hold"
        if status_val is not None and str(status_val).strip() == "On Hold":
            sheet.cell(
                row=row, column=effective_date_col
            ).value = effective_date_value
            updated_count += 1

    print(
        f"📅 Sheet '{sheet.title}': Đã cập nhật EffectiveDate* = '{effective_date_value}' cho {updated_count} dòng 'On Hold'."
    )
    return updated_count

# ==============================================================================
# BƯỚC 3: Xóa tất cả các dòng có Status* KHÔNG PHẢI "On Hold"
# ==============================================================================
import openpyxl


def delete_non_on_hold_rows_fast(
    sheet: openpyxl.worksheet.worksheet.Worksheet, status_col: int
) -> int:
    """Xóa các dòng Status* != 'On Hold' bằng cơ chế lọc & ghi đè (Rewrite)

    Tối ưu siêu nhanh cho 15k+ rows.
    """
    total_original_rows = sheet.max_row
    if total_original_rows < 3:
        return 0

    # 1. Lưu lại thông tin Hàng 1 (Header) và Hàng 2 (Guideline)
    headers = [sheet.cell(row=1, column=c).value for c in range(1, sheet.max_column + 1)]
    guidelines = [sheet.cell(row=2, column=c).value for c in range(1, sheet.max_column + 1)]

    # 2. Duyệt qua bộ nhớ RAM để lọc ra các dòng có Status* == "On Hold"
    kept_rows_data = []
    for r in range(3, total_original_rows + 1):
        status_val = sheet.cell(row=r, column=status_col).value
        if status_val is not None and str(status_val).strip() == "On Hold":
            # Lấy toàn bộ giá trị các ô trong dòng này
            row_values = [sheet.cell(row=r, column=c).value for c in range(1, sheet.max_column + 1)]
            kept_rows_data.append(row_values)

    # 3. Xóa toàn bộ dữ liệu ô cũ trong Sheet
    sheet.delete_rows(1, amount=total_original_rows)

    # 4. Ghi lại Header (Hàng 1) và Guideline (Hàng 2)
    sheet.append(headers)
    sheet.append(guidelines)

    # 5. Ghi lại tất cả các dòng giữ lại (từ Hàng 3 trở đi)
    for row_data in kept_rows_data:
        sheet.append(row_data)

    deleted_count = (total_original_rows - 2) - len(kept_rows_data)
    print(
        f"⚡ [Fast Delete] Đã lọc xong {total_original_rows - 2} dòng. Giữ lại {len(kept_rows_data)} dòng On Hold, xóa {deleted_count} dòng."
    )

    return deleted_count


# ==============================================================================
# HÀM CHA: Điều phối chính cho worksheet 'ClientStatus'
# ==============================================================================
def clean_client_status_sheet(
    workbook: openpyxl.Workbook,
    on_hold_file_path: str = ON_HOLD_FILE_PATH,
) -> openpyxl.Workbook:
    """Hàm cha: Điều phối quy trình dọn dẹp ClientStatus."""
    if CLIENT_STATUS_SHEET_NAME not in workbook.sheetnames:
        print(
            f"⏩ Worksheet '{CLIENT_STATUS_SHEET_NAME}' không tồn tại trong workbook -> Bỏ qua."
        )
        return workbook

    # 1. Gọi hàm BƯỚC 1: Lấy danh sách AC Number từ file external
    on_hold_ids = get_on_hold_client_ids(file_path=on_hold_file_path)
    if not on_hold_ids:
        print(
            f"⚠️ Danh sách On Hold rỗng. Bỏ qua xử lý sheet '{CLIENT_STATUS_SHEET_NAME}'."
        )
        return workbook

    sheet = workbook[CLIENT_STATUS_SHEET_NAME]

    # Tìm vị trí các cột mục tiêu ở Hàng 1
    client_id_col = None
    status_col = None
    effective_date_col = None

    for col in range(1, sheet.max_column + 1):
        cell_val = sheet.cell(row=1, column=col).value
        if cell_val is not None:
            val_str = str(cell_val)
            if val_str == EXACT_CLIENT_ID_HEADER:
                client_id_col = col
            elif val_str == EXACT_STATUS_HEADER:
                status_col = col
            elif val_str == EXACT_EFFECTIVE_DATE_HEADER:
                effective_date_col = col

    if not client_id_col or not status_col:
        print(
            f"⚠️ Sheet '{CLIENT_STATUS_SHEET_NAME}': Không tìm thấy cột '{EXACT_CLIENT_ID_HEADER}' hoặc '{EXACT_STATUS_HEADER}' -> Bỏ qua."
        )
        return workbook

    # 2. Gọi hàm BƯỚC 2: Partial match, update "On Hold", "TODAY" và ghi Log
    update_on_hold_clients_status(
        sheet=sheet,
        client_id_col=client_id_col,
        status_col=status_col,
        on_hold_ids=on_hold_ids,
    )

    update_effective_date_for_on_hold_clients(
        sheet=sheet,
        status_col=status_col,
        effective_date_col=effective_date_col
    )

    deleted_count = delete_non_on_hold_rows_fast(
        sheet=sheet, status_col=status_col
    )

    print(
        f"✨ Sheet '{CLIENT_STATUS_SHEET_NAME}': Đã giữ lại các Client On Hold và xóa {deleted_count} dòng không phải 'On Hold'."
    )

    return workbook
