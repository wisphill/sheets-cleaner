from typing import List, Optional
import openpyxl

from common.remove_column import (
    remove_columns_from_worksheet
)

# Danh sách các cột cần xóa trong sheet Contacts (đã chuẩn hóa chữ thường)
TARGET_CONTACTS_COLUMNS_TO_REMOVE = [
    "Correspondence Method",
    "Import Id",
    "Is Billing Contact",
]


def remove_columns_from_contacts_worksheet(
    workbook: openpyxl.Workbook
) -> openpyxl.Workbook:
    return remove_columns_from_worksheet(workbook, "Contacts", TARGET_CONTACTS_COLUMNS_TO_REMOVE)