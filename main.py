import duckdb
import openpyxl

from handler_01_delete_sheets import (
    delete_sheets,
)

from handler_02_clean_workbooks import (
    clear_branch_name_column,
)

from handler_03_update_tenants import (
    update_tenant_column
)

from handler_05_remove_contacts_columns import (
    remove_columns_from_contacts_worksheet
)

from handler_06_remove_client_contacts_columns import (
    remove_columns_from_clientcontacts_worksheet
)

from handler_07_remove_clients_columns import (
    remove_columns_from_clients_worksheet
)

from handler_08_remove_employees_columns import (
    remove_columns_from_employees_worksheet
)

from handler_11_remove_employee_contacts_columns import (
    remove_columns_from_employee_contacts_worksheet
)

from handler_12_clear_acquired_date import (
    clear_acquired_date_for_police_check
)

from handler_15_clean_sah_funding_episodes import (
    clean_sah_funding_episodes
)

from handler_16_clean_sah_funding_episode_contracts import (
    clean_sah_funding_episode_contracts
)

def load_workbook(input_file: str) -> openpyxl.Workbook:
    """Tải file Excel vào bộ nhớ RAM.

    File gốc trên ổ đĩa KHÔNG bị thay đổi.
    """
    return openpyxl.load_workbook(input_file)


def export_to_new_file(workbook: openpyxl.Workbook, output_file: str):
    """Ghi dữ liệu từ RAM ra một file Excel MỚI hoàn toàn."""
    workbook.save(output_file)
    print(f"📁 Đã xuất thành công ra file MỚI: '{output_file}'")


def load_excel_sheet_duckdb(file_path: str, sheet_name: str):
    """Đọc và truy vấn dữ liệu từ một sheet bằng DuckDB."""
    duckdb.install_extension("spatial")
    duckdb.load_extension("spatial")

    query = f"""
        SELECT * 
        FROM st_read('{file_path}', layer='{sheet_name}')
    """
    return duckdb.query(query).df()


# ==========================================
# Main process
# ==========================================
if __name__ == "__main__":
    input_file = "dovida_staging_2026-09-23_010126_Import_Template_v0.0.53.xlsx"
    output_file = "new_data.xlsx"
    
    wb = load_workbook(input_file=input_file)
    wb = delete_sheets(wb)
    wb = clear_branch_name_column(workbook=wb)
    wb = update_tenant_column(workbook=wb)
    wb = remove_columns_from_contacts_worksheet(workbook=wb)
    wb = remove_columns_from_clientcontacts_worksheet(wb)
    wb = remove_columns_from_clients_worksheet(wb)
    wb = remove_columns_from_employees_worksheet(wb)
    wb = remove_columns_from_employee_contacts_worksheet(wb)
    wb = clear_acquired_date_for_police_check(wb)
    wb = clean_sah_funding_episodes(wb)
    wb = clean_sah_funding_episode_contracts(wb)

    # Export data from RAM to the new file
    export_to_new_file(workbook=wb, output_file=output_file)