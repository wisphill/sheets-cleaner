import openpyxl

from common.remove_column import (
    remove_columns_from_worksheet
)

# Danh sách các cột cần xóa trong sheet Contacts (đã chuẩn hóa chữ thường)
TARGET_COLUMNS_TO_REMOVE = [
    "Status",
    "Correspondence Method",
]


def remove_columns_from_clientcontacts_worksheet(
    workbook: openpyxl.Workbook
) -> openpyxl.Workbook:
    return remove_columns_from_worksheet(workbook, "ClientContacts", TARGET_COLUMNS_TO_REMOVE)