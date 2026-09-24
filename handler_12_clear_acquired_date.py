import openpyxl

# Tên cột và giá trị cần kiểm tra khớp chính xác
EXACT_SKILLS_HEADER = "Skill*"
EXACT_ACQUIRED_DATE_HEADER = "Acquired Date"
EXACT_TARGET_SKILL = "Police Check"


def clear_acquired_date_for_police_check(
    workbook: openpyxl.Workbook, sheet_name: str = "EmployeeSkills"
) -> openpyxl.Workbook:
    """Xóa dữ liệu ở cột 'Acquired Date' đối với các dòng có Skill là 'Police Check' trong worksheet 'EmployeeSkills' (So sánh chính xác hoa/thường).

    - Kiểm tra vị trí cột ở Hàng 1.
    - Hàng 2 là Guideline -> Bỏ qua.
    - Xóa dữ liệu từ Hàng 3 trở đi.
    """
    if sheet_name not in workbook.sheetnames:
        print(
            f"⏩ Worksheet '{sheet_name}' không tồn tại trong workbook -> Bỏ qua."
        )
        return workbook

    sheet = workbook[sheet_name]

    skills_col_idx = None
    acquired_date_col_idx = None

    # 1. Tìm vị trí chính xác của cột 'Skills*' và 'Acquired Date' ở Hàng 1
    for col in range(1, sheet.max_column + 1):
        cell_val = sheet.cell(row=1, column=col).value
        if cell_val is not None:
            clean_val = str(cell_val).strip()
            if clean_val == EXACT_SKILLS_HEADER:
                skills_col_idx = col
            elif clean_val == EXACT_ACQUIRED_DATE_HEADER:
                acquired_date_col_idx = col

    # 2. Thực hiện xóa dữ liệu nếu tìm thấy cả 2 cột
    if skills_col_idx and acquired_date_col_idx:
        cleared_count = 0

        # Duyệt từ Hàng 3 trở đi (bỏ qua Hàng 2 Guideline)
        for row in range(3, sheet.max_row + 1):
            skill_val = sheet.cell(row=row, column=skills_col_idx).value

            # So sánh chính xác chuỗi "Police Check"
            if (
                skill_val is not None
                and str(skill_val).strip() == EXACT_TARGET_SKILL
            ):
                date_cell = sheet.cell(row=row, column=acquired_date_col_idx)
                if date_cell.value is not None:
                    date_cell.value = None
                    cleared_count += 1

        print(
            f"🧹 Sheet '{sheet_name}': Đã xóa dữ liệu 'Acquired Date' cho {cleared_count} dòng có Skill = '{EXACT_TARGET_SKILL}'."
        )
    else:
        missing = []
        if not skills_col_idx:
            missing.append(f"'{EXACT_SKILLS_HEADER}'")
        if not acquired_date_col_idx:
            missing.append(f"'{EXACT_ACQUIRED_DATE_HEADER}'")
        print(
            f"⚠️ Sheet '{sheet_name}': Không tìm thấy cột {', '.join(missing)} ở Hàng 1 -> Bỏ qua."
        )

    return workbook