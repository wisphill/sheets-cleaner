import openpyxl

# Exact header names matching requirements
EXACT_MAIN_CLIENT_ID_HEADER = "Id*"
EXACT_CLIENT_ID_HEADER = "ClientId*"
EXACT_CLIENT_ID_HEADER_2 = "ClientId"
EXACT_STATUS_HEADER = "Status*"

# Target value to trigger deletion
TARGET_STATUS_VALUE = "Discharged"


def remove_discharged_clients(
    workbook: openpyxl.Workbook,
    client_status_sheet_name: str = "ClientStatus",
    clients_sheet_name: str = "Clients",
) -> openpyxl.Workbook:
    """1. Lọc ra các ClientId* có Status* là 'Discharged' từ worksheet 'ClientStatus'.

    2. Xóa các dòng tương ứng trong worksheet 'Clients' (dựa theo Id*).
    3. Quét TẤT CẢ các worksheet và xóa các dòng chứa ClientId* bị loại bỏ.
    4. Sử dụng phương pháp Filter & Rebuild Sheet để đạt hiệu năng tối ưu.
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
            if val_str in (EXACT_CLIENT_ID_HEADER, EXACT_CLIENT_ID_HEADER_2):
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
    status_rows_to_keep = []

    # Giữ Hàng 1 (Header) và Hàng 2 (Guideline)
    for r in range(1, min(3, status_sheet.max_row + 1)):
        status_rows_to_keep.append(
            [
                status_sheet.cell(row=r, column=c).value
                for c in range(1, status_sheet.max_column + 1)
            ]
        )

    # Đọc dữ liệu từ Hàng 3 trở đi
    deleted_status_count = 0
    for row in range(3, status_sheet.max_row + 1):
        status_val = status_sheet.cell(row=row, column=status_col_idx).value
        client_id_val = status_sheet.cell(
            row=row, column=status_client_id_col_idx
        ).value

        is_discharged = (
            status_val is not None
            and str(status_val).strip() == TARGET_STATUS_VALUE
        )

        if is_discharged:
            deleted_status_count += 1
            if client_id_val is not None:
                discharged_client_ids.add(str(client_id_val).strip())
        else:
            status_rows_to_keep.append(
                [
                    status_sheet.cell(row=row, column=c).value
                    for c in range(1, status_sheet.max_column + 1)
                ]
            )

    if not discharged_client_ids:
        print(
            f"ℹ️ Sheet '{client_status_sheet_name}': Không có Client nào có"
            f" Status là '{TARGET_STATUS_VALUE.capitalize()}'."
        )
        return workbook

    # Cập nhật lại sheet ClientStatus
    _rebuild_sheet_content(status_sheet, status_rows_to_keep)
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
            clients_rows_to_keep = []
            deleted_clients_count = 0

            # Giữ Hàng 1 & 2
            for r in range(1, min(3, clients_sheet.max_row + 1)):
                clients_rows_to_keep.append(
                    [
                        clients_sheet.cell(row=r, column=c).value
                        for c in range(1, clients_sheet.max_column + 1)
                    ]
                )

            # Lọc các dòng Client
            for row in range(3, clients_sheet.max_row + 1):
                client_id_val = clients_sheet.cell(
                    row=row, column=main_id_col_idx
                ).value
                id_str = (
                    str(client_id_val).strip()
                    if client_id_val is not None
                    else ""
                )

                if id_str in discharged_client_ids:
                    deleted_clients_count += 1
                else:
                    clients_rows_to_keep.append(
                        [
                            clients_sheet.cell(row=row, column=c).value
                            for c in range(1, clients_sheet.max_column + 1)
                        ]
                    )

            _rebuild_sheet_content(clients_sheet, clients_rows_to_keep)
            print(
                f"🧹 Sheet '{clients_sheet_name}': Đã xóa"
                f" {deleted_clients_count} dòng chứa ClientId* bị Discharged."
            )

    # --- BƯỚC 3: Quét TẤT CẢ các worksheet khác để làm sạch theo ClientId* ---
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
                if val_str in (EXACT_CLIENT_ID_HEADER, EXACT_CLIENT_ID_HEADER_2):
                    client_id_col_idx = col
                    break

        # Tiến hành lọc và xóa dòng nếu sheet có cột ClientId*
        if client_id_col_idx:
            rows_to_keep = []
            deleted_child_count = 0

            # Giữ Hàng 1 & 2
            for r in range(1, min(3, sheet.max_row + 1)):
                rows_to_keep.append(
                    [
                        sheet.cell(row=r, column=c).value
                        for c in range(1, sheet.max_column + 1)
                    ]
                )

            # Lọc dữ liệu từ Hàng 3
            for row in range(3, sheet.max_row + 1):
                cell_val = sheet.cell(row=row, column=client_id_col_idx).value
                cell_str = (
                    str(cell_val).strip() if cell_val is not None else ""
                )

                if cell_str in discharged_client_ids:
                    deleted_child_count += 1
                else:
                    rows_to_keep.append(
                        [
                            sheet.cell(row=row, column=c).value
                            for c in range(1, sheet.max_column + 1)
                        ]
                    )

            if deleted_child_count > 0:
                _rebuild_sheet_content(sheet, rows_to_keep)
                print(
                    f"🧹 Related Sheet '{sheet_name}': Đã xóa"
                    f" {deleted_child_count} dòng chứa ClientId* bị Discharged."
                )

    return workbook


def _rebuild_sheet_content(
    worksheet: openpyxl.worksheet.worksheet.Worksheet, rows_data: list
):
    """Xóa toàn bộ nội dung worksheet và ghi lại dữ liệu đã lọc."""
    worksheet.delete_rows(1, worksheet.max_row)
    for row in rows_data:
        worksheet.append(row)