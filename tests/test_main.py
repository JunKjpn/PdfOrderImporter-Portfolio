from pathlib import Path
from unittest.mock import patch

from src.constants.paths import OUTPUT_DIR
from src.constants.branches import BRANCHES, BRANCH_A
from src.constants.excel import SUMMARY_HEADER_ROW, SUMMARY_SHEET_NAME
from src.main import run_conversion


def test_run_conversion():
    excel_path = Path("template.xlsx")
    invoice_path = Path("invoice.xlsx")
    tray_path = Path("tray.xlsx")
    delivery_note_path = Path("delivery_note.xlsx")
    pdf_paths = [Path("order.pdf")]
    label_output_path = Path("label_output.pdf")

    product_master = {"商品A": "product"}
    branch_columns = {BRANCH_A: 1}
    delivery_date = "2026/09/20"
    branch_orders = {
        BRANCH_A: {
            "商品A": 10,
        }
    }

    summary_path = Path("summary.xlsx")
    invoice_output_path = Path("invoice_output.xlsx")
    tray_output_path = Path("tray_output.xlsx")
    delivery_note_output_path = Path("delivery_note_output.xlsx")

    with (
        patch("src.main.load_product_master", return_value=product_master) as mock_load_product_master,
        patch("src.main.load_workbook") as mock_load_workbook,
        patch("src.main.find_branch_columns", return_value=branch_columns) as mock_find_branch_columns,
        patch(
            "src.main.load_pdf_orders",
            return_value=(delivery_date, branch_orders),
        ) as mock_load_pdf_orders,
        patch("src.main.validate_product_names") as mock_validate_product_names,
        patch("src.main.create_summary", return_value=summary_path) as mock_create_summary,
        patch("src.main.recalculate_workbook") as mock_recalculate_workbook,
        patch(
            "src.main.transfer_invoice",
            return_value=invoice_output_path,
        ) as mock_transfer_invoice,
        patch(
            "src.main.transfer_tray",
            return_value=tray_output_path,
        ) as mock_transfer_tray,
        patch(
            "src.main.transfer_delivery_note",
            return_value=delivery_note_output_path,
        ) as mock_transfer_delivery_note,
        patch(
            "src.main.create_label_pdf_output",
            return_value=(label_output_path, 10),
        ) as mock_create_label_pdf_output,
    ):
        worksheet = object()
        mock_load_workbook.return_value.__getitem__.return_value = worksheet

        result = run_conversion(
            excel_path=excel_path,
            invoice_path=invoice_path,
            tray_path=tray_path,
            delivery_note_path=delivery_note_path,
            pdf_paths=pdf_paths,
            label_floor=1,
            label_start_position=1,
        )

    assert result == (
        summary_path,
        invoice_output_path,
        tray_output_path,
        delivery_note_output_path,
        label_output_path,
    )

    mock_load_product_master.assert_called_once_with(path=excel_path)

    mock_find_branch_columns.assert_called_once_with(
        worksheet=worksheet,
        branch_names=BRANCHES,
        header_row=SUMMARY_HEADER_ROW,
    )

    mock_load_pdf_orders.assert_called_once_with(
        pdf_paths=pdf_paths,
        branch_names=set(branch_columns),
    )

    mock_validate_product_names.assert_called_once_with(
        branch_orders=branch_orders,
        product_master=product_master,
    )

    mock_create_summary.assert_called_once_with(
        excel_path=excel_path,
        sheet_name=SUMMARY_SHEET_NAME,
        product_master=product_master,
        branch_columns=branch_columns,
        branch_orders=branch_orders,
        delivery_date=delivery_date,
        output_dir=OUTPUT_DIR,
    )

    mock_recalculate_workbook.assert_called_once_with(
        excel_path=summary_path,
    )

    mock_transfer_invoice.assert_called_once_with(
        summary_path=summary_path,
        invoice_path=invoice_path,
        delivery_date=delivery_date,
        output_dir=OUTPUT_DIR,
    )

    mock_transfer_tray.assert_called_once_with(
        summary_path=summary_path,
        tray_path=tray_path,
        delivery_date=delivery_date,
        output_dir=OUTPUT_DIR,
    )

    mock_transfer_delivery_note.assert_called_once_with(
        summary_path=summary_path,
        delivery_note_path=delivery_note_path,
        delivery_date=delivery_date,
        output_dir=OUTPUT_DIR,
    )

    mock_create_label_pdf_output.assert_called_once_with(
        excel_path=excel_path,
        pdf_paths=pdf_paths,
        floor=1,
        start_position=1,
        output_dir=OUTPUT_DIR,
    )