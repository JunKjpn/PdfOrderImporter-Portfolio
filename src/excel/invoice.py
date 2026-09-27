import shutil
from datetime import date
from pathlib import Path

import win32com
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from src.constants.excel import PROFIT_CALCULATION_SHEET_NAME, SUMMARY_PRODUCT_COLUMN, SUMMARY_START_ROW, \
    SUMMARY_HEADER_ROW, INVOICE_DATE_COLUMN, INVOICE_START_COLUMN, INVOICE_HEADER_ROW, INVOICE_START_ROW
from src.constants.branches import BRANCHES
from src.utils.converter import normalize_date, normalize_amount, normalize_text
from src.excel.finder import find_row_by_value, find_column_by_value, find_branch_columns
from src.excel.invoice_sheet import find_invoice_sheet_name

# =========================================================
# 集計表読込
# =========================================================
def load_invoice_amounts(summary_path: Path) -> dict[str, int | float]:
    """
    集計表の「計算」シートから、
    営業所ごとの請求金額を取得する。
    戻り値例:
    {BRANCH_A: 10000, BRANCH_B: 12000,}
    """
    workbook = load_workbook(filename=summary_path, data_only=True, read_only=True)

    try:
        if PROFIT_CALCULATION_SHEET_NAME not in workbook.sheetnames:
            raise ValueError(
                "集計表に必要なシートが見つかりません。\n\n"
                f"必要なシート：{PROFIT_CALCULATION_SHEET_NAME}"
            )

        worksheet = workbook[PROFIT_CALCULATION_SHEET_NAME]

        total_row = find_row_by_value(worksheet=worksheet, search_value="合計", column_number=SUMMARY_PRODUCT_COLUMN)

        amount_row = total_row + SUMMARY_START_ROW

        branch_columns = find_branch_columns(
            worksheet=worksheet, branch_names=BRANCHES, header_row=total_row + SUMMARY_HEADER_ROW,
        )

        branch_amounts: dict[str, int | float] = {}

        for branch_name, column_number in branch_columns.items():
            cell_value = worksheet.cell(row=amount_row, column=column_number).value
            branch_amounts[branch_name] = normalize_amount(cell_value)

        return branch_amounts

    finally:
        workbook.close()


# =========================================================
# 請求書処理
# =========================================================
def find_invoice_branch_columns(worksheet: Worksheet) -> dict[str, int]:
    """
    B列から「合計」列の直前までを営業所列として取得する。
    """
    total_column = find_column_by_value(
        worksheet=worksheet,
        search_value="合計",
        row_number=INVOICE_HEADER_ROW,
        start_column=INVOICE_START_COLUMN,
    )

    branch_columns: dict[str, int] = {}

    for column_number in range(INVOICE_START_COLUMN, total_column):
        cell_value = worksheet.cell(row=INVOICE_HEADER_ROW, column=column_number).value

        if cell_value is None:
            continue

        branch_name = normalize_text(cell_value)
        if branch_name:
            branch_columns[branch_name] = column_number

    if not branch_columns:
        raise ValueError(
            "請求書に営業所列が見つかりませんでした。\n\n"
            f"対象シート：{worksheet.title}"
        )

    return branch_columns

def find_invoice_date_rows(worksheet: Worksheet) -> dict[date, int]:
    """
    A列の9行目から「小計」行の直前までを確認し、日付と行番号の対応を返す。
    """
    subtotal_row = find_row_by_value(worksheet=worksheet, search_value="小計", column_number=INVOICE_DATE_COLUMN)

    date_rows: dict[date, int] = {}

    for row_number in range(INVOICE_START_ROW, subtotal_row):
        cell_value = worksheet.cell(row=row_number, column=INVOICE_DATE_COLUMN).value
        parsed_date = normalize_date(value=cell_value)
        if parsed_date is not None:
            date_rows[parsed_date] = row_number

    if not date_rows:
        raise ValueError(
            "請求書に日付行が見つかりませんでした。\n\n"
            f"対象シート：{worksheet.title}"
        )

    return date_rows

def transfer_invoice(summary_path: Path, invoice_path: Path, delivery_date: date, output_dir: Path) -> Path:
    """
    集計表から営業所ごとの合計金額を取得し、納品日に対応する請求書シートへ転記する。
    請求書ファイルは上書き保存する。
    Args:
        summary_path: PDF取込処理で作成した集計表のパス。
        invoice_path: 転記対象となる請求書ファイルのパス。
        delivery_date: 転記対象となる納品日。
    Returns:
        上書き保存した請求書ファイルのパス。
    """
    branch_amounts = load_invoice_amounts(summary_path=summary_path)

    read_workbook = load_workbook(filename=invoice_path, data_only=False, keep_links=True)

    try:
        invoice_sheet_name = find_invoice_sheet_name(workbook=read_workbook, delivery_date=delivery_date)

        read_worksheet = read_workbook[invoice_sheet_name]

        invoice_branch_columns = find_invoice_branch_columns(worksheet=read_worksheet)
        invoice_date_rows = find_invoice_date_rows(worksheet=read_worksheet)

    finally:
        read_workbook.close()

    if delivery_date not in invoice_date_rows:
        raise ValueError(
            f"請求書に納品日の行が見つかりませんでした。\n\n対象シート：{invoice_sheet_name}\n納品日：{delivery_date:%Y/%m/%d}"
        )

    target_row = invoice_date_rows[delivery_date]
    output_path = output_dir / f"{invoice_path.stem}_{delivery_date:%Y%m%d}{invoice_path.suffix}"
    shutil.copy2(invoice_path, output_path)

    excel = None
    workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Open(str(output_path.resolve()))
        worksheet = workbook.Worksheets(invoice_sheet_name)
        for branch_name, amount in branch_amounts.items():
            normalized_branch_name = normalize_text(branch_name)

            if normalized_branch_name not in invoice_branch_columns:
                print(f"請求書に営業所がないためスキップ: {branch_name}")
                continue

            target_column = invoice_branch_columns[normalized_branch_name]

            worksheet.Cells(target_row, target_column).Value = amount
        workbook.Save()

        return output_path

    except Exception as error:
        raise RuntimeError(
            "請求書ファイルへの転記または保存に失敗しました。\nExcelで開いている場合は閉じてから、もう一度実行してください。\n\n"
            f"{output_path}"
        ) from error

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)

        if excel is not None:
            excel.EnableEvents = True
            excel.Quit()
