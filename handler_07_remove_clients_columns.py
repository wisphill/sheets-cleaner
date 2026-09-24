import openpyxl

from common.remove_column import (
    remove_columns_from_worksheet
)

TARGET_COLUMNS_TO_REMOVE = [
    "Care Needs",
    "Consent For Future Contacts",
    "Consent to Provide Details",
    "Country of Birth Code",
    "Emergency Response Level",
    "Has Carer",
    "Hcp Statement Password",
    "Is Billing Contact",
    "Is Birth Date An Estimate",
    "Is Using Pseudonym",
    "Send Invoices",
    "Timeframe",
    "Contact Custom Stmt Delivery",
]


def remove_columns_from_clients_worksheet(
    workbook: openpyxl.Workbook
) -> openpyxl.Workbook:
    return remove_columns_from_worksheet(workbook, "Clients", TARGET_COLUMNS_TO_REMOVE)