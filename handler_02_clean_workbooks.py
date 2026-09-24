import openpyxl

TARGET_BRANCH_NAME_HEADERS = ["BranchName*"]
def clear_branch_name_column(workbook: openpyxl.Workbook) -> openpyxl.Workbook:
    """Xoá dữ liệu cột 'BranchName*' (hoặc 'Branch name') ở dòng 1 trong tất cả worksheets.

    - Kiểm tra tiêu đề ở Hàng 1.
    - Hàng 2 là Guideline -> Bỏ qua.
    - Xoá nội dung từ Hàng 3 trở đi.
    - Nếu không thấy cột ở Hàng 1 -> Bỏ qua worksheet đó.
    """

    for sheet in workbook.worksheets:
        target_col_idx = None

        # 1. Duyệt các ô ở dòng 1 để tìm cột BranchName*
        for col_idx in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col_idx).value
            if cell_value is not None:
                # Chuẩn hóa chuỗi (xóa khoảng trắng thừa và chuyển về chữ thường)
                clean_value = str(cell_value).strip()
                if clean_value in TARGET_BRANCH_NAME_HEADERS:
                    target_col_idx = col_idx
                    break

        # 2. Nếu tìm thấy cột ở dòng 1, tiến hành xóa dữ liệu từ dòng 3 trở đi
        if target_col_idx:
            cleared_count = 0
            for row_idx in range(3, sheet.max_row + 1):
                cell = sheet.cell(row=row_idx, column=target_col_idx)
                if cell.value is not None:
                    cell.value = None
                    cleared_count += 1

            print(
                f"🧹 Sheet '{sheet.title}': Đã xóa dữ liệu cột '{sheet.cell(row=1, column=target_col_idx).value}' ({cleared_count} dòng, bỏ qua Guideline dòng 2)."
            )
        else:
            print(
                f"⏩ Sheet '{sheet.title}': Không tìm thấy cột 'BranchName*' ở dòng 1 -> Bỏ qua."
            )

    return workbook