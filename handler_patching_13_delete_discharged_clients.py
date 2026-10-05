import copy
import openpyxl

# Exact header names matching requirements
EXACT_MAIN_CLIENT_ID_HEADER = "Id*"
EXACT_CLIENT_ID_HEADER = "ClientId*"
EXACT_CLIENT_ID_HEADER_2 = "ClientId"
EXACT_CLIENT_ID_HEADER_3 = "Client ID*"
EXACT_STATUS_HEADER = "Status*"

# Target value to trigger deletion
TARGET_STATUS_VALUE = "Discharged"


def normalize_id(val) -> str:
    """Chuẩn hóa Client ID về dạng chuỗi sạch để so sánh chính xác."""
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def _copy_cell_formatting(src_cell, dst_cell):
    """Copy toàn bộ giá trị và định dạng từ src_cell sang dst_cell."""
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy.copy(src_cell.protection)
        dst_cell.alignment = copy.copy(src_cell.alignment)


def _clear_cell(cell):
    """Xóa dữ liệu và định dạng của ô."""
    cell.value = None


def _fast_filter_sheet(sheet, target_col_idx, ids_to_delete, start_row=3):
    """Lọc và loại bỏ dòng bằng thuật toán In-place Overwrite (ghi đè nội bộ).

    Nhanh gấp hàng chục lần so với việc gọi delete_rows() liên tục.
    """
    max_row = sheet.max_row
    max_col = sheet.max_column

    if max_row < start_row:
        return 0

    write_row = start_row
    deleted_count = 0

    for read_row in range(start_row, max_row + 1):
        cell_val = sheet.cell(row=read_row, column=target_col_idx).value
        norm_id = normalize_id(cell_val)

        # Nếu dòng thuộc danh sách cần xóa
        if norm_id in ids_to_delete:
            deleted_count += 1
            continue

        # Nếu cần di chuyển dòng lên trên
        if write_row != read_row:
            for col in range(1, max_col + 1):
                src_cell = sheet.cell(row=read_row, column=col)
                dst_cell = sheet.cell(row=write_row, column=col)
                _copy_cell_formatting(src_cell, dst_cell)

        write_row += 1

    # Nếu có dòng bị xóa, gọi delete_rows 1 LẦN DUY NHẤT ở cuối sheet
    if deleted_count > 0:
        sheet.delete_rows(write_row, deleted_count)

    return deleted_count


def remove_discharged_clients(
    workbook: openpyxl.Workbook,
    client_status_sheet_name: str = "ClientStatus",
    clients_sheet_name: str = "Clients",
) -> openpyxl.Workbook:
    """1. Lọc ra các ClientId* có Status* là 'Discharged' từ worksheet 'ClientStatus'.

    2. Xóa các dòng tương ứng trong worksheet 'Clients' (dựa theo Id*).
    3. Quét TẤT CẢ các worksheet và xóa các dòng chứa ClientId* bị loại bỏ.
    4. Tối ưu performance bằng In-place Overwrite + 1 lần delete_rows ở cuối.
    """
    # --- BƯỚC 1: Tìm danh sách ClientId* có Status* = 'Discharged' từ sheet ClientStatus ---
    if client_status_sheet_name not in workbook.sheetnames:
        print(
            f"⏩ Worksheet '{client_status_sheet_name}' không tồn tại trong"
            " workbook -> Bỏ qua."
        )
        return workbook

    status_sheet = workbook[client_status_sheet_name]
    status_client_id_col_idx = None
    status_col_idx = None

    for col in range(1, status_sheet.max_column + 1):
        cell_val = status_sheet.cell(row=1, column=col).value
        if cell_val is not None:
            val_str = str(cell_val).strip()
            if val_str in (EXACT_CLIENT_ID_HEADER, EXACT_CLIENT_ID_HEADER_2, EXACT_CLIENT_ID_HEADER_3):
                status_client_id_col_idx = col
            elif val_str == EXACT_STATUS_HEADER:
                status_col_idx = col

    if not status_client_id_col_idx or not status_col_idx:
        print(
            f"⚠️ Sheet '{client_status_sheet_name}': Không tìm thấy cột"
            f" '{EXACT_CLIENT_ID_HEADER}' hoặc '{EXACT_STATUS_HEADER}' ở Hàng 1"
            " -> Bỏ qua."
        )
        return workbook

    discharged_client_ids = set()

    # Quét dữ liệu từ Hàng 3 trở đi để thu thập ID bị Discharged
    for row in range(3, status_sheet.max_row + 1):
        status_val = status_sheet.cell(row=row, column=status_col_idx).value
        client_id_val = status_sheet.cell(
            row=row, column=status_client_id_col_idx
        ).value

        if (
            status_val is not None
            and str(status_val).strip() == TARGET_STATUS_VALUE
        ):
            norm_id = normalize_id(client_id_val)
            if norm_id:
                discharged_client_ids.add(norm_id)

    if not discharged_client_ids:
        print(
            f"ℹ️ Sheet '{client_status_sheet_name}': Không có Client nào có"
            f" Status là '{TARGET_STATUS_VALUE}'."
        )
        return workbook

    # Xóa các dòng Discharged trong sheet ClientStatus (dùng hàm tối ưu)
    deleted_status_count = _fast_filter_sheet(
        status_sheet,
        status_client_id_col_idx,
        discharged_client_ids,
        start_row=3,
    )

    print(
        f"🗑️ Sheet '{client_status_sheet_name}': Đã xóa {deleted_status_count}"
        f" dòng Discharged. Danh sách ClientId* bị loại bỏ:"
        f" {discharged_client_ids}"
    )

    # --- BƯỚC 2: Xóa trong sheet 'Clients' dựa trên Id* ---
    if clients_sheet_name in workbook.sheetnames:
        clients_sheet = workbook[clients_sheet_name]
        main_id_col_idx = None

        for col in range(1, clients_sheet.max_column + 1):
            cell_val = clients_sheet.cell(row=1, column=col).value
            if cell_val is not None:
                if str(cell_val).strip() == EXACT_MAIN_CLIENT_ID_HEADER:
                    main_id_col_idx = col
                    break

        if main_id_col_idx:
            deleted_clients_count = _fast_filter_sheet(
                clients_sheet, main_id_col_idx, discharged_client_ids, start_row=3
            )
            print(
                f"🧹 Sheet '{clients_sheet_name}': Đã xóa"
                f" {deleted_clients_count} dòng chứa ClientId* bị Discharged."
            )

    # --- BƯỚC 3: Quét TẤT CẢ các worksheet còn lại để làm sạch theo ClientId* ---
    for sheet_name in workbook.sheetnames:
        if sheet_name in (client_status_sheet_name, clients_sheet_name):
            continue

        sheet = workbook[sheet_name]
        client_id_col_idx = None

        # Tìm vị trí cột ClientId* hoặc ClientId
        for col in range(1, sheet.max_column + 1):
            cell_val = sheet.cell(row=1, column=col).value
            if cell_val is not None:
                val_str = str(cell_val).strip()
                if val_str in (EXACT_CLIENT_ID_HEADER, EXACT_CLIENT_ID_HEADER_2, EXACT_CLIENT_ID_HEADER_3):
                    client_id_col_idx = col
                    break

        if client_id_col_idx:
            deleted_child_count = _fast_filter_sheet(
                sheet, client_id_col_idx, discharged_client_ids, start_row=3
            )
            if deleted_child_count > 0:
                print(
                    f"🧹 Related Sheet '{sheet_name}': Đã xóa"
                    f" {deleted_child_count} dòng chứa ClientId* bị Discharged."
                )

    return workbook