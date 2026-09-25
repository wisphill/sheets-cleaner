import re
from datetime import date, datetime
from typing import Optional, Union
import openpyxl

from common.remove_column import (
    remove_columns_from_worksheet
)

from common.const import GO_LIVE_DATE

# Tên chính xác tiêu đề cột (So sánh nguyên văn)
EXACT_PROGRAM_HEADERS = ["Program"]
EXACT_FUNDING_METHODOLOGY_HEADER = "Funding Methodology*"

# Pattern Regex bắt định dạng ngày YYYY-MM-DD
DATE_REGEX_PATTERN = re.compile(r"\b\d{4}-\d{2}-\d{2}\b")
TARGET_COLUMNS_TO_REMOVE = [
    "service_is_hidden",
    "Emergency Response Level"
]

# ==============================================================================
# HÀM CON 1: Xử lý mục (c) - Điền 'Support at Home'
# ==============================================================================
def update_funding_methodology_for_sah(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    program_col_idx: int,
    funding_method_col_idx: Optional[int],
) -> int:
    """Nhiệm vụ (c): Nếu cột Program chứa 'SAH' và 'Funding Methodology*' bị
    trống -> Điền 'Support at Home'."""
    if not funding_method_col_idx:
        return 0

    updated_count = 0
    for row in range(3, sheet.max_row + 1):  # Bỏ qua Hàng 2 Guideline
        prog_val = sheet.cell(row=row, column=program_col_idx).value

        if prog_val is not None and "SAH" in str(prog_val):
            funding_cell = sheet.cell(row=row, column=funding_method_col_idx)
            # Nếu ô trống (None hoặc chuỗi rỗng)
            if (
                funding_cell.value is None
                or str(funding_cell.value).strip() == ""
            ):
                funding_cell.value = "Support at Home"
                updated_count += 1

    return updated_count


# ==============================================================================
# HÀM CON 2: Xử lý mục (d) - Re-date chuỗi ngày trong Program
# ==============================================================================
def redate_program_labels_for_sah(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    program_col_idx: int,
    formatted_go_live: str,
) -> int:
    """Nhiệm vụ (d): Tìm chuỗi ngày YYYY-MM-DD trong cột Program của các dịch vụ
    SAH và thay thế bằng go_live_date (Giữ nguyên prefix và suffix như (AT
    Medium), (AT High)...)."""
    redated_count = 0

    for row in range(3, sheet.max_row + 1):  # Bỏ qua Hàng 2 Guideline
        prog_cell = sheet.cell(row=row, column=program_col_idx)
        prog_val = prog_cell.value

        if prog_val is not None:
            prog_str = str(prog_val)

            # Chỉ áp dụng re-date cho các label SAH có chứa định dạng ngày YYYY-MM-DD
            if "SAH" in prog_str and DATE_REGEX_PATTERN.search(prog_str):
                new_prog_str = DATE_REGEX_PATTERN.sub(
                    formatted_go_live, prog_str
                )

                if new_prog_str != prog_str:
                    prog_cell.value = new_prog_str
                    redated_count += 1

    return redated_count


# ==============================================================================
# HÀM CHA: Điều phối chính cho worksheet 'Services'
# ==============================================================================
def clean_services_sheet(
    workbook: openpyxl.Workbook,
    go_live_date: Union[str, datetime, date] = GO_LIVE_DATE,
    sheet_name: str = "Services",
) -> openpyxl.Workbook:
    """Hàm cha: Quản lý và gọi các hàm con dọn dẹp dữ liệu worksheet 'Services'."""
    if sheet_name not in workbook.sheetnames:
        print(
            f"⏩ Worksheet '{sheet_name}' không tồn tại trong workbook -> Bỏ qua."
        )
        return workbook

    # --------------------------------------------------------------------------
    # BƯỚC 1 (Task a): Gọi hàm chung để xóa 2 cột "service_is_hidden" và "Emergency Response Level"
    # --------------------------------------------------------------------------
    workbook = remove_columns_from_worksheet(
        workbook=workbook,
        sheet_name=sheet_name,
        columns_to_remove=TARGET_COLUMNS_TO_REMOVE,
    )
    sheet = workbook[sheet_name]

    # Chuẩn hóa go_live_date thành chuỗi YYYY-MM-DD
    if isinstance(go_live_date, (datetime, date)):
        formatted_go_live = go_live_date.strftime("%Y-%m-%d")
    else:
        formatted_go_live = str(go_live_date).strip()

    program_col_idx = None
    funding_method_col_idx = None

    # Tìm vị trí các cột ở Hàng 1
    for col in range(1, sheet.max_column + 1):
        cell_val = sheet.cell(row=1, column=col).value
        if cell_val is not None:
            val_str = str(cell_val)
            if val_str in EXACT_PROGRAM_HEADERS:
                program_col_idx = col
            elif val_str == EXACT_FUNDING_METHODOLOGY_HEADER:
                funding_method_col_idx = col

    if not program_col_idx:
        print(
            f"⚠️ Sheet '{sheet_name}': Không tìm thấy cột 'Program*' hoặc 'Program' ở Hàng 1 -> Bỏ qua."
        )
        return workbook

    # --- GỌI CÁC HÀM CON ---
    # 2. Gọi hàm con xử lý mục (c)
    updated_funding_count = update_funding_methodology_for_sah(
        sheet=sheet,
        program_col_idx=program_col_idx,
        funding_method_col_idx=funding_method_col_idx,
    )

    # 3. Gọi hàm con xử lý mục (d)
    redated_program_count = redate_program_labels_for_sah(
        sheet=sheet,
        program_col_idx=program_col_idx,
        formatted_go_live=formatted_go_live,
    )

    print(
        f"✨ Sheet '{sheet_name}': "
        f"Đã gán 'Support at Home' cho {updated_funding_count} dòng. "
        f"Đã re-date Program thành '{formatted_go_live}' cho {redated_program_count} dòng."
    )

    return workbook