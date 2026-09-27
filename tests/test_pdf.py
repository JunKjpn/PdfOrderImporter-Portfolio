from src.constants.branches import BRANCH_C
from src.pdf.validator import validate_product_names
from src.utils.path import get_app_dir
from src.pdf.parser import extract_order_items


PDF_PATH = get_app_dir() / "sample" / "pdf" / "260922営業所C（サンプル）.pdf"

def test_extract_order_items(monkeypatch):
    lines = [
        "商品A 10",
        "商品B 5",
        "商品A 3",
        "商品C(バリエーション1) 8",
        "NEW商品E 10",
        "商品E 3",
        "合計 100",
    ]

    monkeypatch.setattr(
        "src.pdf.parser.extract_lines",
        lambda pdf_path: lines,
    )

    result = extract_order_items(
        pdf_path=PDF_PATH,
    )

    assert result == {
        "商品A": 13,
        "商品B": 5,
        "商品C(バリエーション1)": 8,
        "商品E": 13,
    }
    assert "合計" not in result


def test_validate_product_names():
    branch_orders = {
        BRANCH_C: {
            "商品A": 10,
            "商品B": 5,
        }
    }

    product_master = {
        "商品A": ...,
        "商品B": ...,
    }

    validate_product_names(
        branch_orders=branch_orders,
        product_master=product_master,
    )

import pytest

def test_validate_product_names_unknown_product():
    branch_orders = {
        BRANCH_C: {
            "商品A": 10,
            "商品C(バリエーション1)": 8,
        }
    }

    product_master = {
        "商品A": ...,
    }

    with pytest.raises(ValueError, match="商品C"):
        validate_product_names(
            branch_orders=branch_orders,
            product_master=product_master,
        )