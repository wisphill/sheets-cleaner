from typing import List, Optional
import openpyxl



def remove_columns_from_worksheet(
    workbook: openpyxl.Workbook,
    sheet_name: str,
    columns_to_remove: Optional[List[str]] = None,
) -> openpyxl.Workbook:
    """Xóa các cột chỉ định trong worksheet dựa theo tiêu đề ở Hàng 1.

    :param workbook: Đối tượng openpyxl Workbook trong RAM.
    :param sheet_name: Tên sheet cần xử lý
    :param columns_to_remove: Danh sách tên cột cần xóa.
    """
    targets_clean = [col.strip() for col in columns_to_remove]

    # Tìm sheet Contacts (không phân biệt hoa thường)
    target_sheet = None
    for sheet in workbook.worksheets:
        if sheet.title.strip() == sheet_name.strip():
            target_sheet = sheet
            break

    if not target_sheet:
        print(f"⏩ Không tìm thấy worksheet '{sheet_name}' -> Bỏ qua.")
        return workbook

    # 1. Tìm chỉ số các cột cần xóa ở Hàng 1
    cols_to_delete_idx = []
    for col_idx in range(1, target_sheet.max_column + 1):
        cell_value = target_sheet.cell(row=1, column=col_idx).value
        if cell_value is not None:
            clean_value = str(cell_value).strip()
            if clean_value in targets_clean:
                cols_to_delete_idx.append((col_idx, str(cell_value).strip()))

    # 2. Xóa các cột từ phải sang trái (index lớn đến bé) để tránh bị lệch chỉ số
    if cols_to_delete_idx:
        # Sắp xếp giảm dần theo chỉ số cột (col_idx)
        cols_to_delete_idx.sort(key=lambda x: x[0], reverse=True)

        for col_idx, col_name in cols_to_delete_idx:
            target_sheet.delete_cols(col_idx)
            print(
                f"🗑️ Sheet '{target_sheet.title}': Đã xóa cột '{col_name}' (Cột số {col_idx})."
            )
    else:
        print(
            f"⏩ Sheet '{target_sheet.title}': Không tìm thấy cột nào thuộc danh sách cần xóa ở Hàng 1."
        )

    return workbook