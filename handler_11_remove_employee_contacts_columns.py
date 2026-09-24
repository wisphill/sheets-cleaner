import openpyxl

from common.remove_column import (
    remove_columns_from_worksheet
)

TARGET_COLUMNS_TO_REMOVE = [
    "Status",
    "Correspondence Method", # not to be confused with “Password*”
    "Is Billing Contact",
]


def remove_columns_from_employee_contacts_worksheet(
    workbook: openpyxl.Workbook
) -> openpyxl.Workbook:
    return remove_columns_from_worksheet(workbook, "EmployeeContacts", TARGET_COLUMNS_TO_REMOVE)