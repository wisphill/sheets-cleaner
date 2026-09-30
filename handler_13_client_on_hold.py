import os
import sys
from typing import Dict, List, Optional, Set
import openpyxl
from openpyxl.styles import PatternFill

# Khai báo các tiêu đề cột chuẩn (So sánh nguyên văn/chuẩn hóa)
CLIENT_STATUS_SHEET_NAME = "ClientStatus"
EXACT_CLIENT_ID_HEADER = "ClientId*"
EXACT_STATUS_HEADER = "Status*"
EXACT_EFFECTIVE_DATE_HEADER = "EffectiveDate*"
EXACT_AC_NUMBER_HEADER = "AC Number"

# Cột cần xóa nội dung khi On Hold
EXACT_STATUS_REASON_HEADER = "StatusReason"
EXACT_NOTE_HEADER = "Note"

ON_HOLD_FILE_PATH = "./data/Clients on hold as of 28.9.2026.xlsx"


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
        if val is not None and str(val).strip() != "":
            on_hold_ids.add(str(val).strip())

    print(
        f"📋 Đã tải {len(on_hold_ids)} AC Number (On Hold) từ file '{file_path}'."
    )
    return on_hold_ids


# ==============================================================================
# BƯỚC 2 & 3: Xử lý gom nhóm, Cập nhật On Hold (tô đỏ) & Lấy dòng Status cuối cùng
# ==============================================================================
def clean_client_status_sheet(
    workbook: openpyxl.Workbook,
    on_hold_file_path: str = ON_HOLD_FILE_PATH,
) -> openpyxl.Workbook:
    """Hàm điều phối dọn dẹp ClientStatus:

    1. Lọc nhóm On Hold từ file bên ngoài: Update Status="On Hold", EffectiveDate="TODAY",
       xóa StatusReason, Note. Giữ 1 dòng duy nhất và tô màu đỏ.
    2. Các Client còn lại: Lấy dòng cuối cùng (Status mới nhất) cho mỗi ClientId*.
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

    # --- ĐỌC VÀ PHÂN LOẠI DỮ LIỆU ---
    # Cấu trúc lưu trữ: { client_id_str: [list_of_row_values] }
    normal_clients_map: Dict[str, List[list]] = {}

    # Cấu trúc cho nhóm On Hold: { matched_ac_number: [list_of_row_values] }
    on_hold_clients_map: Dict[str, List[list]] = {}

    for r in range(3, total_rows + 1):
        row_vals = [
            sheet.cell(row=r, column=c).value for c in range(1, total_cols + 1)
        ]
        client_id_val = sheet.cell(row=r, column=client_id_col).value

        if client_id_val is me_or_none(client_id_val):
            continue

        client_id_str = str(client_id_val).strip()

        # Kiểm tra xem row này có thuộc Client nằm trong danh sách On Hold hay không
        matched_ac = None
        for ac_num in on_hold_ac_numbers:
            if ac_num in client_id_str:
                matched_ac = ac_num
                break

        if matched_ac:
            if matched_ac not in on_hold_clients_map:
                on_hold_clients_map[matched_ac] = []
            on_hold_clients_map[matched_ac].append(row_vals)
        else:
            if client_id_str not in normal_clients_map:
                normal_clients_map[client_id_str] = []
            normal_clients_map[client_id_str].append(row_vals)

    # --- BƯỚC 2: TẠO DỮ LIỆU DÒNG GIỮ LẠI ---
    final_rows_data = []  # Danh sách các dòng dữ liệu sẽ ghi lại
    red_row_indexes = (
        set()
    )  # Lưu chỉ số dòng (1-indexed) để tô màu đỏ sau khi ghi

    # Dùng Fill đỏ đậm/nhẹ tùy nhu cầu (ở đây dùng Solid Red)
    red_fill = PatternFill(
        start_color="FFFF0000", end_color="FFFF0000", fill_type="solid"
    )

    current_output_row = 3  # Hàng dữ liệu bắt đầu từ Row 3

    # 1. Xử lý nhóm ON HOLD
    matched_on_hold_count = 0
    for ac_num, rows_list in on_hold_clients_map.items():
        matched_on_hold_count += 1
        # Lấy dòng đầu tiên đại diện cho client này
        rep_row = list(rows_list[0])

        # Cập nhật thông tin On Hold
        rep_row[status_col - 1] = "On Hold"
        if effective_date_col:
            rep_row[effective_date_col - 1] = "TODAY"

        # Xóa dữ liệu StatusReason và Note
        if status_reason_col:
            rep_row[status_reason_col - 1] = None
        if note_col:
            rep_row[note_col - 1] = None

        final_rows_data.append(rep_row)
        red_row_indexes.add(current_output_row)
        current_output_row += 1

    unmatched_on_hold = on_hold_ac_numbers - set(on_hold_clients_map.keys())
    if unmatched_on_hold:
        print(
            f"⚠️ [LOG] {len(unmatched_on_hold)} AC Number On Hold không tìm thấy"
            f" trong sheet: {sorted(list(unmatched_on_hold))}"
        )

    # 2. Xử lý nhóm CLIENT THƯỜNG (Lấy dòng cuối cùng / status mới nhất)
    for client_id, rows_list in normal_clients_map.items():
        latest_row = rows_list[-1]  # Lấy row có chỉ số lớn nhất (cuối cùng)
        final_rows_data.append(latest_row)
        current_output_row += 1

    # --- BƯỚC 3: XÓA VÀ GHI ĐÈ BẢNG (REWRITE) ---
    headers = [
        sheet.cell(row=1, column=c).value for c in range(1, total_cols + 1)
    ]
    guidelines = [
        sheet.cell(row=2, column=c).value for c in range(1, total_cols + 1)
    ]

    sheet.delete_rows(1, amount=total_rows)

    sheet.append(headers)
    sheet.append(guidelines)

    for row_vals in final_rows_data:
        sheet.append(row_vals)

    # --- BƯỚC 4: TÔ MÀU ĐỎ CHO CÁC DÒNG ON HOLD ĐƯỢC TẠO TỪ FILE NGOẠI BÙ ---
    for r_idx in red_row_indexes:
        for c_idx in range(1, total_cols + 1):
            sheet.cell(row=r_idx, column=c_idx).fill = red_fill

    deleted_count = (total_rows - 2) - len(final_rows_data)
    print(
        f"⚡ Sheet '{CLIENT_STATUS_SHEET_NAME}': Hoàn tất lọc trùng & dọn dẹp.\n"
        f"   - Tổng số Client On Hold (đã tô đỏ): {matched_on_hold_count}\n"
        f"   - Tổng số Client thường (lấy row cuối): {len(normal_clients_map)}\n"
        f"   - Tổng số dòng đã bị loại bỏ: {deleted_count}"
    )

    return workbook


def me_or_none(val) -> bool:
    """Helper kiểm tra giá trị ô có rỗng hay không."""
    return val is None or str(val).strip() == ""