import copy
import os
import sys
from typing import Dict, List, Optional, Set
import openpyxl

# Khai báo các tiêu đề cột chuẩn
CLIENT_STATUS_SHEET_NAME = "ClientStatus"
EXACT_CLIENT_ID_HEADER = "ClientId*"
EXACT_STATUS_HEADER = "Status*"
EXACT_EFFECTIVE_DATE_HEADER = "EffectiveDate*"
EXACT_AC_NUMBER_HEADER = "AC Number"

# Cột cần xóa nội dung khi On Hold
EXACT_STATUS_REASON_HEADER = "StatusReason"
EXACT_NOTE_HEADER = "Note"

ON_HOLD_FILE_PATH = "./data/Clients on hold as of 28.9.2026.xlsx"


def is_empty_or_none(val) -> bool:
    """Helper kiểm tra giá trị ô có rỗng hay không."""
    return val is None or str(val).strip() == ""


# ==============================================================================
# BƯỚC 1: Đọc danh sách AC Number từ file Client_on_hold.xlsx
# ==============================================================================
def get_on_hold_client_ids(
    file_path: str = ON_HOLD_FILE_PATH,
) -> Set[str]:
    """Trích xuất tất cả các AC Number (Clients đang On Hold) từ file excel."""
    if not os.path.exists(file_path):
        print(
            f"❌ File '{file_path}' không tồn tại. Không thể lấy danh sách"
            " Client On Hold."
        )
        return set()

    wb_hold = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb_hold.active

    ac_number_col_idx = None
    on_hold_ids = set()

    for col in range(1, sheet.max_column + 1):
        cell_val = sheet.cell(row=1, column=col).value
        if cell_val is not None and EXACT_AC_NUMBER_HEADER in str(cell_val):
            ac_number_col_idx = col
            break

    if not ac_number_col_idx:
        print(
            f"❌ Không tìm thấy cột '{EXACT_AC_NUMBER_HEADER}' trong file"
            f" '{file_path}'. Dừng chương trình."
        )
        sys.exit(1)

    for row in range(2, sheet.max_row + 1):
        val = sheet.cell(row=row, column=ac_number_col_idx).value
        if not is_empty_or_none(val):
            on_hold_ids.add(str(val).strip())

    print(
        f"📋 Đã tải {len(on_hold_ids)} AC Number (On Hold) từ file '{file_path}'."
    )
    return on_hold_ids


# ==============================================================================
# BƯỚC 2, 3 & 4: Xử lý gom nhóm, Cập nhật On Hold & Giữ lại Status Discharged
# ==============================================================================
def clean_client_status_sheet(
    workbook: openpyxl.Workbook,
    on_hold_file_path: str = ON_HOLD_FILE_PATH,
) -> openpyxl.Workbook:
    """Hàm điều phối dọn dẹp ClientStatus:

    1. Gom nhóm theo ClientId* (lấy dòng dưới cùng / mới nhất cho mỗi ClientId*).
    2. Chỉ giữ lại:
       - Các Client thuộc danh sách On Hold (Cập nhật Status="On Hold", EffectiveDate="TODAY", xóa StatusReason, Note).
       - Các Client có Status = "Discharged" sau khi đã gom nhóm.
    3. Loại bỏ tất cả các Client/Dòng khác không thỏa mãn 2 điều kiện trên.
    4. Tối ưu hiệu năng bằng In-Place Shift và bảo toàn 100% Format gốc.
    """
    if CLIENT_STATUS_SHEET_NAME not in workbook.sheetnames:
        print(
            f"⏩ Worksheet '{CLIENT_STATUS_SHEET_NAME}' không tồn tại trong"
            " workbook -> Bỏ qua."
        )
        return workbook

    # 1. Tải danh sách AC Number On Hold
    on_hold_ac_numbers = get_on_hold_client_ids(file_path=on_hold_file_path)

    sheet = workbook[CLIENT_STATUS_SHEET_NAME]

    # Bản đồ lưu chỉ số cột
    col_indices = {}
    for col in range(1, sheet.max_column + 1):
        cell_val = sheet.cell(row=1, column=col).value
        if cell_val is not None:
            val_str = str(cell_val).strip()
            col_indices[val_str] = col

    client_id_col = col_indices.get(EXACT_CLIENT_ID_HEADER)
    status_col = col_indices.get(EXACT_STATUS_HEADER)
    effective_date_col = col_indices.get(EXACT_EFFECTIVE_DATE_HEADER)
    status_reason_col = col_indices.get(EXACT_STATUS_REASON_HEADER)
    note_col = col_indices.get(EXACT_NOTE_HEADER)

    if not client_id_col or not status_col:
        print(
            f"⚠️ Sheet '{CLIENT_STATUS_SHEET_NAME}': Không tìm thấy cột"
            f" '{EXACT_CLIENT_ID_HEADER}' hoặc '{EXACT_STATUS_HEADER}' ở Hàng 1"
            " -> Bỏ qua."
        )
        return workbook

    total_cols = sheet.max_column
    total_rows = sheet.max_row

    if total_rows < 3:
        return workbook

    # --- BƯỚC 1: GOM NHÓM THEO CLIENT ID (LẤY DÒNG CUỐI CÙNG/MỚI NHẤT) ---
    # Phân loại riêng dòng thuộc On Hold và dòng không thuộc On Hold
    normal_clients_target_row: Dict[str, int] = {}
    on_hold_clients_target_row: Dict[str, int] = {}

    for r in range(3, total_rows + 1):
        client_id_val = sheet.cell(row=r, column=client_id_col).value

        if is_empty_or_none(client_id_val):
            continue

        client_id_str = str(client_id_val).strip()

        # Kiểm tra ClientId có thuộc danh sách On Hold hay không
        matched_ac = None
        for ac_num in on_hold_ac_numbers:
            if ac_num in client_id_str:
                matched_ac = ac_num
                break

        if matched_ac:
            # Lưu dòng cuối cùng tìm thấy cho client On Hold này
            on_hold_clients_target_row[matched_ac] = r
        else:
            # Lưu dòng cuối cùng tìm thấy cho client thông thường
            normal_clients_target_row[client_id_str] = r

    target_rows_set = set(on_hold_clients_target_row.values()) | set(normal_clients_target_row.values())
    on_hold_rows_set = set(on_hold_clients_target_row.values())

    # --- BƯỚC 2: TỐI ƯU VỚI IN-PLACE SHIFT (CHỈ GIỮ ON HOLD VÀ DISCHARGED) ---
    write_row = 3
    matched_on_hold_count = 0
    discharged_retained_count = 0
    deleted_other_count = 0

    for read_row in range(3, total_rows + 1):
        if read_row not in target_rows_set:
            continue

        is_on_hold = read_row in on_hold_rows_set

        # Đọc giá trị Status hiện tại của dòng được chọn
        current_status_val = sheet.cell(row=read_row, column=status_col).value
        status_str = str(current_status_val).strip().lower() if not is_empty_or_none(current_status_val) else ""

        # NẾU KHÔNG PHẢI ON HOLD: Chỉ giữ lại khi Status = "discharged"
        if not is_on_hold:
            if status_str == "discharged":
                discharged_retained_count += 1
            else:
                # Bỏ qua các dòng không nằm trong danh sách On Hold và không phải Discharged
                deleted_other_count += 1
                continue
        else:
            matched_on_hold_count += 1

        # Nếu vị trí ghi (write_row) khác vị trí đọc (read_row), di chuyển dữ liệu & format
        if write_row != read_row:
            for c in range(1, total_cols + 1):
                src_cell = sheet.cell(row=read_row, column=c)
                dst_cell = sheet.cell(row=write_row, column=c)

                dst_cell.value = src_cell.value
                if src_cell.has_style:
                    dst_cell.number_format = src_cell.number_format
                    dst_cell.font = copy.copy(src_cell.font)
                    dst_cell.border = copy.copy(src_cell.border)
                    dst_cell.fill = copy.copy(src_cell.fill)
                    dst_cell.alignment = copy.copy(src_cell.alignment)

        # Cập nhật dữ liệu cho các dòng thuộc On Hold
        if is_on_hold:
            sheet.cell(row=write_row, column=status_col).value = "On Hold"
            if effective_date_col:
                sheet.cell(row=write_row, column=effective_date_col).value = "TODAY"
            if status_reason_col:
                sheet.cell(row=write_row, column=status_reason_col).value = None
            if note_col:
                sheet.cell(row=write_row, column=note_col).value = None

        write_row += 1

    # --- BƯỚC 3: XÓA DÒNG DƯ Ở CUỐI SHEET ---
    new_max_row = write_row - 1
    if total_rows > new_max_row:
        rows_to_delete = total_rows - new_max_row
        sheet.delete_rows(new_max_row + 1, amount=rows_to_delete)

    # Cảnh báo các AC Number không tìm thấy trong sheet
    unmatched_on_hold = on_hold_ac_numbers - set(on_hold_clients_target_row.keys())
    if unmatched_on_hold:
        print(
            f"⚠️ [LOG] {len(unmatched_on_hold)} AC Number On Hold không tìm thấy"
            f" trong sheet: {sorted(list(unmatched_on_hold))}"
        )

    deleted_count = (total_rows - 2) - (new_max_row - 2)

    print(
        f"⚡ Sheet '{CLIENT_STATUS_SHEET_NAME}': Hoàn tất lọc trùng & dọn dẹp.\n"
        f"   - Tổng số Client On Hold giữ lại: {matched_on_hold_count}\n"
        f"   - Tổng số Client Discharged giữ lại: {discharged_retained_count}\n"
        f"   - Tổng số dòng đã bị loại bỏ: {deleted_count}"
    )

    return workbook