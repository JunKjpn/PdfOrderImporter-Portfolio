from pathlib import Path

import pandas as pd

from src.constants.excel import SUMMARY_FILE_NAME
from src.excel.summary import merge_product_orders
from src.excel.product_master import load_product_master


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = PROJECT_ROOT / "templates" / SUMMARY_FILE_NAME

def test_load_product_master():
    products = load_product_master(TEMPLATE_PATH)

    assert len(products) == 19

    france_bread = products["商品A"]

    assert france_bread.product_id == 1
    assert france_bread.units_per_box == 7
    assert france_bread.packaging_floor == 1

    chocolate_croissant = products["商品D"]

    assert chocolate_croissant.product_id == 7
    assert chocolate_croissant.units_per_box == 12
    assert chocolate_croissant.packaging_floor == 2

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