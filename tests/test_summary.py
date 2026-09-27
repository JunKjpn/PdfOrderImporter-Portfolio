from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch

import pandas as pd

from src.constants.paths import OUTPUT_DIR
from src.constants.branches import BRANCH_C, BRANCH_A
from src.constants.excel import (
    BOX_CONVERSION_SHEET_NAME,
    BOX_TABLE_SHEET_NAME,
    BOX_HEADER_ROW,
    BOX_START_ROW,
    BOX_UNITS_PER_BOX_LABEL,
    PROFIT_CALCULATION_SHEET_NAME,
    PROFIT_HEADER_ROW,
    PROFIT_MARGIN_UNIT_PRICE_LABEL,
    PROFIT_START_ROW,
    PROFIT_UNIT_PRICE_LABEL,
    SUMMARY_SHEET_NAME,
    TRAY_HEADER_ROW,
    TRAY_ORDER_MULTIPLIER_LABEL,
    TRAY_SHEET_NAME,
    TRAY_START_ROW,
    TRAY_STORE_SALES_UNITS_LABEL, SUMMARY_FILE_NAME,
)
from src.excel.summary import (
    create_summary,
    load_branch_colors,
    merge_product_orders,
    merge_all_orders,
)
from tests.helpers import create_test_product

PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = PROJECT_ROOT / "templates" / SUMMARY_FILE_NAME


def test_load_branch_colors():
    branch_colors = load_branch_colors(TEMPLATE_PATH)

    assert branch_colors[BRANCH_C] is None

    for branch_name, color in branch_colors.items():
        if color is None:
            continue

        assert len(color) == 3
        assert all(0 <= value <= 255 for value in color)

        print(f"{branch_name}: {color}")


def test_merge_product_orders():
    base_df = pd.DataFrame(
        {
            "商品名": [
                "商品A",
                "商品B",
                "商品C(バリエーション1)",
            ]
        }
    )

    orders = {
        "商品B": 5,
        "商品A": 10,
    }

    result = merge_product_orders(
        base_df=base_df,
        orders=orders,
    )

    expected = pd.DataFrame(
        {
            "商品名": [
                "商品A",
                "商品B",
                "商品C(バリエーション1)",
            ],
            "数量": [10, 5, 0],
        }
    )

    pd.testing.assert_frame_equal(result, expected)


def test_merge_all_orders():
    product_master = {
        "商品A": create_test_product(
            units_per_box=1,
        ),
        "商品B": create_test_product(
            product_id=2,
            name="商品B",
            units_per_box=1,
        ),
        "商品C(バリエーション2)": create_test_product(
            product_id=3,
            name="商品C(バリエーション2)",
            units_per_box=1,
        ),
    }

    branch_orders = {
        BRANCH_C: {
            "商品C(バリエーション2)": 8,
        },
        BRANCH_A: {
            "商品A": 10,
        },
    }

    result = merge_all_orders(
        product_master=product_master,
        branch_orders=branch_orders,
    )

    expected = pd.DataFrame(
        {
            "商品名": [
                "商品A",
                "商品C(バリエーション2)",
            ]
        }
    )

    pd.testing.assert_frame_equal(result, expected)


def test_create_summary_writes_product_master_values():
    product = create_test_product(
        product_id=1,
        name="商品A",
        units_per_box=7,
    )

    product_master = {
        "商品A": product,
    }

    branch_columns = {
        BRANCH_A: 2,
    }

    branch_orders = {
        BRANCH_A: {
            "商品A": 10,
        },
    }

    workbook = MagicMock()
    worksheet = MagicMock()
    workbook.sheetnames = [
        SUMMARY_SHEET_NAME,
        BOX_CONVERSION_SHEET_NAME,
        TRAY_SHEET_NAME,
        PROFIT_CALCULATION_SHEET_NAME,
        BOX_TABLE_SHEET_NAME,
    ]

    workbook.__getitem__.side_effect = {
        BOX_CONVERSION_SHEET_NAME: MagicMock(),
        TRAY_SHEET_NAME: MagicMock(),
        PROFIT_CALCULATION_SHEET_NAME: MagicMock(),
        BOX_TABLE_SHEET_NAME: MagicMock(),
    }.__getitem__

    with (
        patch(
            "src.excel.summary.open_workbook",
            return_value=(workbook, worksheet),
        ),
        patch(
            "src.excel.summary.find_row_by_value",
            return_value=10,
        ),
        patch("src.excel.summary.clear_cell_range"),
        patch("src.excel.summary.write_product_names"),
        patch("src.excel.summary.write_orders"),
        patch(
            "src.excel.summary.write_product_master_value"
        ) as mock_write_master,
        patch(
            "src.excel.summary.write_box_table_units"
        ) as mock_write_box_table,
        patch("src.excel.summary.save_workbook"),
    ):
        result = create_summary(
            excel_path=Path("template.xlsx"),
            sheet_name=SUMMARY_SHEET_NAME,
            product_master=product_master,
            branch_columns=branch_columns,
            branch_orders=branch_orders,
            delivery_date=date(2026, 9, 20),
            output_dir=OUTPUT_DIR,
        )

    assert result == OUTPUT_DIR / "template_20260920.xlsx"

    assert mock_write_master.call_count == 5

    expected_master_calls = [
        {
            "product_attribute": "order_multiplier",
            "search_value": TRAY_ORDER_MULTIPLIER_LABEL,
            "header_row_num": TRAY_HEADER_ROW,
            "data_row_num": TRAY_START_ROW,
        },
        {
            "product_attribute": "store_sales_units",
            "search_value": TRAY_STORE_SALES_UNITS_LABEL,
            "header_row_num": TRAY_HEADER_ROW,
            "data_row_num": TRAY_START_ROW,
        },
        {
            "product_attribute": "unit_price",
            "search_value": PROFIT_UNIT_PRICE_LABEL,
            "header_row_num": PROFIT_HEADER_ROW,
            "data_row_num": PROFIT_START_ROW,
        },
        {
            "product_attribute": "margin_unit_price",
            "search_value": PROFIT_MARGIN_UNIT_PRICE_LABEL,
            "header_row_num": PROFIT_HEADER_ROW,
            "data_row_num": PROFIT_START_ROW,
        },
        {
            "product_attribute": "units_per_box",
            "search_value": BOX_UNITS_PER_BOX_LABEL,
            "header_row_num": BOX_HEADER_ROW,
            "data_row_num": BOX_START_ROW,
        },
    ]

    actual_master_calls = [
        {
            "product_attribute": call.kwargs["product_attribute"],
            "search_value": call.kwargs["search_value"],
            "header_row_num": call.kwargs["header_row_num"],
            "data_row_num": call.kwargs["data_row_num"],
        }
        for call in mock_write_master.call_args_list
    ]

    assert actual_master_calls == expected_master_calls

    assert mock_write_box_table.call_count == 2

    assert {
        call.kwargs["value_column"]
        for call in mock_write_box_table.call_args_list
    } == {2, 4}