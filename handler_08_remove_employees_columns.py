import openpyxl

from common.remove_column import (
    remove_columns_from_worksheet
)

TARGET_COLUMNS_TO_REMOVE = [
    "Medical Status",
    "Password", # not to be confused with “Password*”
    "Tags",
]


def remove_columns_from_employees_worksheet(
    workbook: openpyxl.Workbook
) -> openpyxl.Workbook:
    return remove_columns_from_worksheet(workbook, "Employees", TARGET_COLUMNS_TO_REMOVE)