from pathlib import Path

from openpyxl import load_workbook

from src.services.label import create_label_pdf_output
from src.constants.paths import OUTPUT_DIR
from src.constants.excel import SUMMARY_SHEET_NAME, SUMMARY_HEADER_ROW
from src.constants.branches import BRANCHES
from src.excel.calculator import recalculate_workbook
from src.excel.finder import find_branch_columns
from src.excel.delivery import transfer_delivery_note
from src.excel.invoice import transfer_invoice
from src.excel.product_master import load_product_master
from src.excel.summary import create_summary
from src.excel.tray import transfer_tray
from src.pdf.order import load_pdf_orders
from src.pdf.validator import validate_product_names



def run_conversion(
    excel_path: Path,
    invoice_path: Path,
    tray_path: Path,
    delivery_note_path: Path,
    pdf_paths: list[Path],
    label_floor: int | None = None,
    label_start_position: int = 1,
) -> tuple[Path, Path, Path, Path, Path | None]:
    """発注PDFから集計表・請求書・天板並べ早見表・納品書・ラベルPDFを作成する。"""

    # 商品マスタを読み込む
    product_master = load_product_master(path=excel_path)

    # 営業所ごとの列番号を取得する
    workbook = load_workbook(
        filename=excel_path,
        data_only=False,
        read_only=False,
    )

    branch_columns = find_branch_columns(
        worksheet=workbook[SUMMARY_SHEET_NAME],
        branch_names=BRANCHES,
        header_row=SUMMARY_HEADER_ROW,
    )

    # 発注PDFを読み込む
    delivery_date, branch_orders = load_pdf_orders(
        pdf_paths=pdf_paths,
        branch_names=set(branch_columns),
    )

    # 商品名を検証する
    validate_product_names(
        branch_orders=branch_orders,
        product_master=product_master,
    )

    # 集計表を作成する
    summary_path = create_summary(
        excel_path=excel_path,
        sheet_name=SUMMARY_SHEET_NAME,
        product_master=product_master,
        branch_columns=branch_columns,
        branch_orders=branch_orders,
        delivery_date=delivery_date,
        output_dir=OUTPUT_DIR,
    )

    # 集計表を再計算する
    recalculate_workbook(excel_path=summary_path)

    # 請求書へ転記する
    invoice_output_path = transfer_invoice(
        summary_path=summary_path,
        invoice_path=invoice_path,
        delivery_date=delivery_date,
        output_dir=OUTPUT_DIR,
    )

    # 天板並べ早見表へ転記する
    tray_output_path = transfer_tray(
        summary_path=summary_path,
        tray_path=tray_path,
        delivery_date=delivery_date,
        output_dir=OUTPUT_DIR,
    )

    # 納品書へ転記する
    delivery_note_output_path = transfer_delivery_note(
        summary_path=summary_path,
        delivery_note_path=delivery_note_path,
        delivery_date=delivery_date,
        output_dir=OUTPUT_DIR,
    )

    # ラベルPDFを作成する
    label_output_path = None

    if label_floor is not None:
        label_output_path, _ = create_label_pdf_output(
            excel_path=excel_path,
            pdf_paths=pdf_paths,
            floor=label_floor,
            start_position=label_start_position,
            output_dir=OUTPUT_DIR,
        )

    return (
        summary_path,
        invoice_output_path,
        tray_output_path,
        delivery_note_output_path,
        label_output_path,
    )
