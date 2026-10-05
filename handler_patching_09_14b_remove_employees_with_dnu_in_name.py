import openpyxl

# Exact header names matching requirements
EXACT_MAIN_ID_HEADER = "Id*"
EXACT_FIRST_NAME_HEADER = "FirstName*"
EXACT_LAST_NAME_HEADER = "LastName*"
EXACT_EMPLOYEE_ID_HEADER = "EmployeeId*"
EXACT_EMPLOYEE_ID_HEADER_2 = "EmployeeId"

# Substring condition to trigger deletion
TARGET_DNU_SUBSTRING = "dnu"


def _rebuild_sheet_content(
    worksheet: openpyxl.worksheet.worksheet.Worksheet, rows_data: list
):
    """Ghi đè lại dữ liệu vào các ô hiện có để BẢO TOÀN FORMAT (đặc biệt là Row 1 & Row 2),

    sau đó xóa bỏ các dòng thừa ở phía dưới.
    """
    total_new_rows = len(rows_data)
    current_max_row = worksheet.max_row

    # 1. Ghi đè dữ liệu vào các row hiện có (giữ nguyên cell styles/formatting)
    # Gán thẳng .value vì Worksheet.cell(value=None) sẽ KHÔNG ghi đè,
    # khiến ô rỗng giữ lại giá trị cũ của dòng trước khi dồn.
    for row_idx, row_values in enumerate(rows_data, start=1):
        for col_idx, val in enumerate(row_values, start=1):
            worksheet.cell(row=row_idx, column=col_idx).value = val

    # 2. Nếu số dòng sau khi lọc ít hơn số dòng ban đầu, xóa các dòng dư ở cuối
    if current_max_row > total_new_rows:
        rows_to_delete = current_max_row - total_new_rows
        worksheet.delete_rows(total_new_rows + 1, amount=rows_to_delete)


def remove_dnu_employees(
    workbook: openpyxl.Workbook,
    employees_sheet_name: str = "Employees",
) -> openpyxl.Workbook:
    """1. Finds Employees in the 'Employees' sheet where 'FirstName*' or 'LastName*'
       contains 'DNU' and removes those records.
    2. Scans all other sheets for 'EmployeeId*' / 'EmployeeId' columns and
       removes related records for those DNU employees using a fast
       sheet-recreation approach.
    """
    if employees_sheet_name not in workbook.sheetnames:
        print(
            f"⏩ Worksheet '{employees_sheet_name}' does not exist in workbook"
            " -> Skipped."
        )
        return workbook

    emp_sheet = workbook[employees_sheet_name]

    id_col_idx = None
    fname_col_idx = None
    lname_col_idx = None

    # --- STEP 1: Find 'Id*', 'FirstName*', and 'LastName*' column indexes ---
    for col in range(1, emp_sheet.max_column + 1):
        cell_val = emp_sheet.cell(row=1, column=col).value
        if cell_val is not None:
            val_str = str(cell_val).strip()
            if val_str == EXACT_MAIN_ID_HEADER:
                id_col_idx = col
            elif val_str == EXACT_FIRST_NAME_HEADER:
                fname_col_idx = col
            elif val_str == EXACT_LAST_NAME_HEADER:
                lname_col_idx = col

    if not id_col_idx or not fname_col_idx or not lname_col_idx:
        print(
            f"⚠️ Sheet '{employees_sheet_name}': Could not find headers"
            f" '{EXACT_MAIN_ID_HEADER}', '{EXACT_FIRST_NAME_HEADER}', or"
            f" '{EXACT_LAST_NAME_HEADER}' on Row 1 -> Skipped."
        )
        return workbook

    # --- STEP 2: Filter 'Employees' sheet ---
    removed_emp_ids = set()
    emp_rows_to_keep = []

    # Preserve Headers (Row 1) and Guidelines (Row 2)
    for r in range(1, min(3, emp_sheet.max_row + 1)):
        emp_rows_to_keep.append(
            [
                emp_sheet.cell(row=r, column=c).value
                for c in range(1, emp_sheet.max_column + 1)
            ]
        )

    # Process Data Rows (Row 3 onwards)
    deleted_emp_count = 0
    for row in range(3, emp_sheet.max_row + 1):
        fname_val = emp_sheet.cell(row=row, column=fname_col_idx).value
        lname_val = emp_sheet.cell(row=row, column=lname_col_idx).value
        emp_id_val = emp_sheet.cell(row=row, column=id_col_idx).value

        # Standardize strings for search
        fname_str = (
            str(fname_val).strip().lower() if fname_val is not None else ""
        )
        lname_str = (
            str(lname_val).strip().lower() if lname_val is not None else ""
        )

        # Check if 'dnu' is IN FirstName* OR LastName*
        contains_dnu = (
            TARGET_DNU_SUBSTRING in fname_str
            or TARGET_DNU_SUBSTRING in lname_str
        )

        if contains_dnu:
            deleted_emp_count += 1
            if emp_id_val is not None:
                removed_emp_ids.add(str(emp_id_val).strip())
        else:
            # Keep valid employee rows
            row_data = [
                emp_sheet.cell(row=row, column=c).value
                for c in range(1, emp_sheet.max_column + 1)
            ]
            emp_rows_to_keep.append(row_data)

    if not removed_emp_ids:
        print(
            f"ℹ️ Sheet '{employees_sheet_name}': No employees found with 'DNU'"
            " in 'FirstName*' or 'LastName*'."
        )
        return workbook

    # Fast overwrite for 'Employees' sheet
    _rebuild_sheet_content(emp_sheet, emp_rows_to_keep)

    print(
        f"🗑️ Sheet '{employees_sheet_name}': Removed {deleted_emp_count} DNU"
        f" record(s). List of removed IDs: {removed_emp_ids}"
    )

    # --- STEP 3: Clean up all child sheets by EmployeeId ---
    for sheet_name in workbook.sheetnames:
        if sheet_name == employees_sheet_name:
            continue

        sheet = workbook[sheet_name]
        emp_id_col_idx = None

        # Find EmployeeId column position
        for col in range(1, sheet.max_column + 1):
            cell_val = sheet.cell(row=1, column=col).value
            if cell_val is not None:
                val_str = str(cell_val).strip()
                if val_str in (
                    EXACT_EMPLOYEE_ID_HEADER,
                    EXACT_EMPLOYEE_ID_HEADER_2,
                ):
                    emp_id_col_idx = col
                    break

        # Filter child sheet if EmployeeId column exists
        if emp_id_col_idx:
            rows_to_keep = []
            deleted_child_count = 0

            # Preserve Headers (Row 1) and Guidelines (Row 2)
            for r in range(1, min(3, sheet.max_row + 1)):
                rows_to_keep.append(
                    [
                        sheet.cell(row=r, column=c).value
                        for c in range(1, sheet.max_column + 1)
                    ]
                )

            # Evaluate Data Rows (Row 3 onwards)
            for row in range(3, sheet.max_row + 1):
                cell_val = sheet.cell(row=row, column=emp_id_col_idx).value
                cell_str = (
                    str(cell_val).strip() if cell_val is not None else ""
                )

                if cell_str in removed_emp_ids:
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
                    f"🧹 Related Sheet '{sheet_name}': Removed"
                    f" {deleted_child_count} row(s) matching DNU Employee IDs."
                )

    return workbook