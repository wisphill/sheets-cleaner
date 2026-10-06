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
EXACT_REASON_HEADER = "Reason"

# Cột StatusReason được điền từ file On Hold, cột Note bị xóa nội dung
EXACT_STATUS_REASON_HEADER = "StatusReason"
EXACT_NOTE_HEADER = "Note"

ON_HOLD_FILE_PATH = "./data/Clients on hold as of 06.10.2026.xlsx"

# Dấu phân cách trong cột 'Reason': "xxx1 - xxx2 - xxx3" (xxx3 là tùy chọn)
REASON_SEPARATOR = " - "


def is_empty_or_none(val) -> bool:
    """Helper kiểm tra giá trị ô có rỗng hay không."""
    return val is None or str(val).strip() == ""


def extract_status_reason(reason_val) -> Optional[str]:
    """Lấy StatusReason (xxx2) từ giá trị cột 'Reason' dạng "xxx1 - xxx2 - xxx3".

    xxx3 là tùy chọn. Nếu giá trị không có dấu phân cách thì trả về nguyên giá
    trị đã được trim.
    """
    if is_empty_or_none(reason_val):
        return None

    parts = [part.strip() for part in str(reason_val).split(REASON_SEPARATOR)]
    parts = [part for part in parts if part != ""]

    if not parts:
        return None
    if len(parts) == 1:
        return parts[0]
    return parts[1]


# ==============================================================================
# BƯỚC 1: Đọc danh sách AC Number + StatusReason từ file Client_on_hold.xlsx
# ==============================================================================
def get_on_hold_client_ids(
    file_path: str = ON_HOLD_FILE_PATH,
) -> Dict[str, Optional[str]]:
    """Trích xuất AC Number (Clients đang On Hold) và StatusReason từ file excel.

    Trả về bản đồ {AC Number: StatusReason}, trong đó StatusReason là phần xxx2
    của cột 'Reason' (định dạng "xxx1 - xxx2 - xxx3").
    """
    if not os.path.exists(file_path):
        print(
            f"❌ File '{file_path}' không tồn tại. Không thể lấy danh sách"
            " Client On Hold."
        )
        return {}

    wb_hold = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb_hold.active

    ac_number_col_idx = None
    reason_col_idx = None
    on_hold_clients: Dict[str, Optional[str]] = {}

    for col in range(1, sheet.max_column + 1):
        cell_val = sheet.cell(row=1, column=col).value
        if cell_val is None:
            continue
        header = str(cell_val).strip()
        if ac_number_col_idx is None and EXACT_AC_NUMBER_HEADER in header:
            ac_number_col_idx = col
        elif reason_col_idx is None and header.lower() == EXACT_REASON_HEADER.lower():
            reason_col_idx = col

    if not ac_number_col_idx:
        print(
            f"❌ Không tìm thấy cột '{EXACT_AC_NUMBER_HEADER}' trong file"
            f" '{file_path}'. Dừng chương trình."
        )
        sys.exit(1)

    if not reason_col_idx:
        print(
            f"⚠️ Không tìm thấy cột '{EXACT_REASON_HEADER}' trong file"
            f" '{file_path}' -> StatusReason sẽ để trống."
        )

    missing_reason: List[str] = []

    for row in range(2, sheet.max_row + 1):
        val = sheet.cell(row=row, column=ac_number_col_idx).value
        if is_empty_or_none(val):
            continue

        ac_number = str(val).strip()
        status_reason = None
        if reason_col_idx:
            status_reason = extract_status_reason(
                sheet.cell(row=row, column=reason_col_idx).value
            )

        if status_reason is None:
            missing_reason.append(ac_number)

        on_hold_clients[ac_number] = status_reason

    print(
        f"📋 Đã tải {len(on_hold_clients)} AC Number (On Hold) từ file"
        f" '{file_path}'."
    )

    distinct_reasons: Dict[str, int] = {}
    for status_reason in on_hold_clients.values():
        if status_reason:
            distinct_reasons[status_reason] = distinct_reasons.get(status_reason, 0) + 1

    if distinct_reasons:
        print(f"   - StatusReason trích xuất được từ cột '{EXACT_REASON_HEADER}':")
        for status_reason, count in sorted(
            distinct_reasons.items(), key=lambda kv: (-kv[1], kv[0])
        ):
            print(f"     • {status_reason!r}: {count}")

    if missing_reason:
        print(
            f"⚠️ {len(missing_reason)} AC Number không có StatusReason:"
            f" {sorted(missing_reason)}"
        )

    return on_hold_clients


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
       - Các Client thuộc danh sách On Hold (Cập nhật Status="On Hold", EffectiveDate="TODAY", StatusReason lấy từ cột 'Reason', xóa Note).
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

    # 1. Tải danh sách AC Number On Hold kèm StatusReason
    on_hold_status_reasons = get_on_hold_client_ids(file_path=on_hold_file_path)
    on_hold_ac_numbers: Set[str] = set(on_hold_status_reasons.keys())

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

    if not status_reason_col:
        print(
            f"⚠️ Sheet '{CLIENT_STATUS_SHEET_NAME}': Không tìm thấy cột"
            f" '{EXACT_STATUS_REASON_HEADER}' ở Hàng 1 -> Không thể ghi"
            " StatusReason."
        )

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
    # Bản đồ ngược: dòng được giữ lại -> AC Number On Hold tương ứng
    on_hold_row_to_ac: Dict[int, str] = {
        row: ac_num for ac_num, row in on_hold_clients_target_row.items()
    }

    # --- BƯỚC 2: TỐI ƯU VỚI IN-PLACE SHIFT (CHỈ GIỮ ON HOLD VÀ DISCHARGED) ---
    write_row = 3
    matched_on_hold_count = 0
    discharged_retained_count = 0
    deleted_other_count = 0
    applied_status_reasons: Dict[str, int] = {}
    rows_without_status_reason = 0

    for read_row in range(3, total_rows + 1):
        if read_row not in target_rows_set:
            continue

        is_on_hold = read_row in on_hold_row_to_ac

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
                status_reason = on_hold_status_reasons.get(
                    on_hold_row_to_ac[read_row]
                )
                sheet.cell(row=write_row, column=status_reason_col).value = status_reason
                if status_reason:
                    applied_status_reasons[status_reason] = (
                        applied_status_reasons.get(status_reason, 0) + 1
                    )
                else:
                    rows_without_status_reason += 1
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

    if applied_status_reasons:
        total_applied = sum(applied_status_reasons.values())
        print(
            f"   - Đã ghi '{EXACT_STATUS_REASON_HEADER}' cho {total_applied} dòng:"
        )
        for status_reason, count in sorted(
            applied_status_reasons.items(), key=lambda kv: (-kv[1], kv[0])
        ):
            print(f"     • {status_reason}: {count}")

    if rows_without_status_reason:
        print(
            f"⚠️ {rows_without_status_reason} dòng On Hold không có"
            f" '{EXACT_STATUS_REASON_HEADER}' -> để trống."
        )

    return workbook