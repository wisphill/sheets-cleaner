from typing import List, Optional
import openpyxl

# Danh sách các sheet mặc định cần xóa (ví dụ các sheet nháp, tạm)
DEFAULT_SHEETS_TO_DELETE: List[str] = [
    "TempSheet",
    "Draft",
    "Sheet1",
    "Data_Temp",
]

def delete_sheets(
    workbook: openpyxl.Workbook,
    sheets_to_delete: Optional[List[str]] = None,
) -> openpyxl.Workbook:
    """Xóa danh sách các sheet chỉ định trên đối tượng Workbook đang nằm trong RAM.

    :param workbook: Đối tượng openpyxl Workbook trong RAM.
    :param sheets_to_delete: Danh sách tên các sheet cần xóa. Nếu None, sẽ dùng
    DEFAULT_SHEETS_TO_DELETE.
    :return: Đối tượng Workbook đã xử lý.
    """
    # Nếu không truyền danh sách thì sử dụng danh sách mặc định
    targets = (
        sheets_to_delete
        if sheets_to_delete is not None
        else DEFAULT_SHEETS_TO_DELETE
    )

    for sheet_name in targets:
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
            print(f"✅ Đã xóa sheet '{sheet_name}' trong RAM.")
        else:
            print(f"⚠️ Sheet '{sheet_name}' không tồn tại trong workbook.")

    return workbook