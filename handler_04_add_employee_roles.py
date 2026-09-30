from typing import Dict, List, Optional
import openpyxl
from openpyxl.styles import Alignment, Font

EXACT_EMP_ID_HEADER = "EmployeeId*"
EXACT_ROLE_HEADER = "Role*"
EXACT_BRANCH_HEADER = "BranchName*"
EXACT_TENANT_HEADER = "Tenant*"
EXACT_MAIN_ID_HEADER = "Id*"

EXACT_EMAIL_HEADER = "EmailAddress*"

# Guidelines theo chuẩn thiết kế spec
GUIDELINE_EMP_ID = (
    "1. Every entry in this column must also be in one of the columns [Id] in"
    " the sheet [Employees]\n2. NOTE: The contents of this sheet will be merged"
    " into any of the sheets [Employees, EmployeeUpdates] and only migrated"
    " when those sheets are migrated. It is not possible to migrate this data"
    " by itself."
)
GUIDELINE_ROLE = (
    "Value must match the description of a role listed in Roles and"
    " Permissions."
)
GUIDELINE_BRANCH = (
    "Branch: Must match Branch Name exactly. Empty values will be migrated to"
    " HQ."
)
GUIDELINE_TENANT = (
    'Tenant: Must match url prefix. Example, for demo3.alayacare.com, cell'
    ' value should be "demo3"'
)

# Danh sách Admin mặc định sử dụng EmailAddress để lookup
DEFAULT_TARGET_USERS = [
    {
        "email": "matt.oliver@hiscsydnorth.com.au",  # Matthew
    },
    {
        "email": "jo.hegney@dovida-snh.com.au",  # Jo Heyney
    },
    {
        "email": "an@mayflyventures.com",  # An
    },
    {
        "email": "david@mayflyventures.com",  # David
    },
]


def create_employee_roles_with_specific_users(
    workbook: openpyxl.Workbook,
    target_users: Optional[List[Dict[str, Optional[str]]]] = None,
    sheet_name: str = "EmployeeRoles",
    employees_sheet_name: str = "Employees",
) -> openpyxl.Workbook:
    """Lookup 'EmailAddress*' trong sheet Employees để lấy 'Id*', sau đó tạo worksheet 'EmployeeRoles' và gán Role cho các user tìm thấy."""
    if target_users is None:
        target_users = DEFAULT_TARGET_USERS

    # 1. Quét map email -> emp_id từ worksheet Employees
    email_to_id_map = {}

    if employees_sheet_name in workbook.sheetnames:
        emp_sheet = workbook[employees_sheet_name]
        id_col_idx = None
        email_col_idx = None

        # Tìm vị trí cột Id* và EmailAddress* ở Hàng 1
        for col in range(1, emp_sheet.max_column + 1):
            cell_val = emp_sheet.cell(row=1, column=col).value
            if cell_val is not None:
                val_str = str(cell_val).strip()
                if val_str == EXACT_MAIN_ID_HEADER:
                    id_col_idx = col
                elif val_str == EXACT_EMAIL_HEADER:
                    email_col_idx = col

        # Tạo dictionary ánh xạ {email: emp_id} (Duyệt từ Hàng 3 trở đi)
        if id_col_idx and email_col_idx:
            for r in range(3, emp_sheet.max_row + 1):
                id_val = emp_sheet.cell(row=r, column=id_col_idx).value
                email_val = emp_sheet.cell(row=r, column=email_col_idx).value

                if (
                    id_val is not None
                    and email_val is not None
                    and str(email_val).strip() != ""
                ):
                    clean_email = str(email_val).strip().lower()
                    clean_id = str(id_val).strip()
                    email_to_id_map[clean_email] = clean_id

    # 2. Làm mới worksheet EmployeeRoles
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]

    sheet = workbook.create_sheet(title=sheet_name)

    # 3. Tạo Hàng 1 (Headers)
    headers = [
        EXACT_EMP_ID_HEADER,
        EXACT_ROLE_HEADER,
        EXACT_BRANCH_HEADER,
        EXACT_TENANT_HEADER,
    ]
    sheet.append(headers)

    # 4. Tạo Hàng 2 (Guidelines)
    guidelines = [
        GUIDELINE_EMP_ID,
        GUIDELINE_ROLE,
        GUIDELINE_BRANCH,
        GUIDELINE_TENANT,
    ]
    sheet.append(guidelines)

    # Apply Style nghiêng + màu xám cho Hàng 2
    italic_gray_font = Font(
        name="Calibri", size=10, italic=True, color="595959"
    )
    wrap_alignment = Alignment(wrap_text=True, vertical="top")

    for col_idx in range(1, len(guidelines) + 1):
        cell = sheet.cell(row=2, column=col_idx)
        cell.font = italic_gray_font
        cell.alignment = wrap_alignment

    # 5. Lookup theo EmailAddress và ghi dòng vào sheet (từ Hàng 3)
    added_count = 0
    skipped_emails = []

    for user in target_users:
        user_email = str(user.get("email", "")).strip().lower()

        # Lookup email trong map
        found_emp_id = email_to_id_map.get(user_email)

        if found_emp_id:
            row_data = [
                found_emp_id,
                user.get("role", "System Administrator"),
                user.get("branch", None),
                user.get("tenant", "raykay"),
            ]
            sheet.append(row_data)
            added_count += 1
            print(
                f"  ✓ Found: {user_email} -> EmployeeId: {found_emp_id} ("
                f" Role: {user.get('role')} )"
            )
        else:
            skipped_emails.append(user_email)

    print(
        f"✨ Sheet '{sheet_name}': Đã hoàn tất gán Role cho {added_count} user."
    )
    if skipped_emails:
        print(
            f"⚠️ Bỏ qua {len(skipped_emails)} email không tìm thấy trong sheet '{employees_sheet_name}': {skipped_emails}"
        )

    return workbook