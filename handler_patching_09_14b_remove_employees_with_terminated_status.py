import openpyxl

# Exact header names matching requirements
EXACT_MAIN_ID_HEADER = "Id*"
EXACT_STATUS_HEADER = "Status*"
EXACT_EMPLOYEE_ID_HEADER = "EmployeeId*"
EXACT_EMPLOYEE_ID_HEADER_2 = "EmployeeId"

# Value condition to trigger deletion
TARGET_STATUS_VALUE = "Terminated"


def remove_terminated_employees(
    workbook: openpyxl.Workbook,
    employees_sheet_name: str = "Employees",
) -> openpyxl.Workbook:
    """1. Finds Employees in the 'Employees' sheet where 'Status*' is 'Terminated'

    and removes those records.
    2. Scans all other sheets for 'EmployeeId*' / 'EmployeeId' columns and
       removes related records for those terminated employees using a fast
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
    status_col_idx = None

    # --- STEP 1: Find 'Id*' and 'Status*' column indexes ---
    for col in range(1, emp_sheet.max_column + 1):
        cell_val = emp_sheet.cell(row=1, column=col).value
        if cell_val is not None:
            val_str = str(cell_val).strip()
            if val_str == EXACT_MAIN_ID_HEADER:
                id_col_idx = col
            elif val_str == EXACT_STATUS_HEADER:
                status_col_idx = col

    if not id_col_idx or not status_col_idx:
        print(
            f"⚠️ Sheet '{employees_sheet_name}': Could not find headers"
            f" '{EXACT_MAIN_ID_HEADER}' or '{EXACT_STATUS_HEADER}' on Row 1 ->"
            " Skipped."
        )
        return workbook

    # --- STEP 2: Filter 'Employees' sheet ---
    removed_emp_ids = set()
    emp_rows_to_keep = []

    # Preserve Headers (Row 1) and Guidelines (Row 2)
    for r in range(1, min(3, emp_sheet.max_row + 1)):
        emp_rows_to_keep.append(
            [emp_sheet.cell(row=r, column=c).value for c in range(1, emp_sheet.max_column + 1)]
        )

    # Process Data Rows (Row 3 onwards)
    deleted_emp_count = 0
    for row in range(3, emp_sheet.max_row + 1):
        status_val = emp_sheet.cell(row=row, column=status_col_idx).value
        emp_id_val = emp_sheet.cell(row=row, column=id_col_idx).value

        # Check if row is terminated
        is_terminated = (
            status_val is not None
            and str(status_val).strip().lower() == TARGET_STATUS_VALUE.lower()
        )

        if is_terminated:
            deleted_emp_count += 1
            if emp_id_val is not None:
                removed_emp_ids.add(str(emp_id_val).strip())
        else:
            # Keep non-terminated employee rows
            row_data = [
                emp_sheet.cell(row=row, column=c).value
                for c in range(1, emp_sheet.max_column + 1)
            ]
            emp_rows_to_keep.append(row_data)

    if not removed_emp_ids:
        print(
            f"ℹ️ Sheet '{employees_sheet_name}': No employees found with status"
            f" '{TARGET_STATUS_VALUE}'."
        )
        return workbook

    # Fast overwrite for 'Employees' sheet
    _rebuild_sheet_content(emp_sheet, emp_rows_to_keep)

    print(
        f"🗑️ Sheet '{employees_sheet_name}': Removed {deleted_emp_count}"
        f" '{TARGET_STATUS_VALUE}' record(s). List of removed IDs:"
        f" {removed_emp_ids}"
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
                if val_str in (EXACT_EMPLOYEE_ID_HEADER, EXACT_EMPLOYEE_ID_HEADER_2):
                    emp_id_col_idx = col
                    break

        # Filter child sheet if EmployeeId column exists
        if emp_id_col_idx:
            rows_to_keep = []
            deleted_child_count = 0

            # Preserve Headers (Row 1) and Guidelines (Row 2)
            for r in range(1, min(3, sheet.max_row + 1)):
                rows_to_keep.append(
                    [sheet.cell(row=r, column=c).value for c in range(1, sheet.max_column + 1)]
                )

            # Evaluate Data Rows (Row 3 onwards)
            for row in range(3, sheet.max_row + 1):
                cell_val = sheet.cell(row=row, column=emp_id_col_idx).value
                cell_str = str(cell_val).strip() if cell_val is not None else ""

                if cell_str in removed_emp_ids:
                    deleted_child_count += 1
                else:
                    rows_to_keep.append(
                        [sheet.cell(row=row, column=c).value for c in range(1, sheet.max_column + 1)]
                    )

            if deleted_child_count > 0:
                _rebuild_sheet_content(sheet, rows_to_keep)
                print(
                    f"🧹 Related Sheet '{sheet_name}': Removed"
                    f" {deleted_child_count} row(s) matching terminated IDs."
                )

    return workbook


def _rebuild_sheet_content(worksheet: openpyxl.worksheet.worksheet.Worksheet, rows_data: list):
    """Clears all cells in a worksheet and bulk re-writes preserved row data."""
    worksheet.delete_rows(1, worksheet.max_row)  # Clears all contents fast
    for row in rows_data:
        worksheet.append(row)