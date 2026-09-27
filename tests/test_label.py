from pathlib import Path

import pytest

from src.constants.excel import SUMMARY_FILE_NAME
from src.constants.branches import BRANCH_A
from src.services.label import build_label_orders, split_label_order, split_labels_by_floor, place_labels, \
    calculate_next_position, build_label_placements
from src.models.label import LabelOrder, LabelItem
from tests.helpers import create_test_product

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SOURCE_PATH = PROJECT_ROOT / "templates" / SUMMARY_FILE_NAME
SUMMARY_PATH = PROJECT_ROOT / "output" / "集計表_20260922.xlsx"


def test_build_label_orders():
    product = create_test_product()

    branch_orders = {
        "営業所A": {
            "商品A": 20,
        }
    }

    product_master = {
        "商品A": product,
    }

    orders = build_label_orders(
        branch_orders=branch_orders,
        product_master=product_master,
    )

    assert len(orders) == 1
    assert orders[0].product.name == "商品A"
    assert orders[0].branch_name == BRANCH_A
    assert orders[0].quantity == 20

def test_split_label_order():
    product = create_test_product()

    order = LabelOrder(
        product=product,
        branch_name=BRANCH_A,
        quantity=20,
    )

    items = split_label_order(order)

    assert [item.quantity for item in items] == [7, 7, 6]


def test_split_labels_by_floor():
    product_1f = create_test_product()

    product_2f = create_test_product(
        product_id=7,
        name="商品D",
        units_per_box=12,
        packaging_floor=2,
    )

    items = [
        LabelItem(
            product=product_1f,
            branch_name=BRANCH_A,
            quantity=7,
        ),
        LabelItem(
            product=product_2f,
            branch_name=BRANCH_A,
            quantity=12,
        ),
    ]

    result = split_labels_by_floor(items)

    assert len(result[1]) == 1
    assert len(result[2]) == 1
    assert result[1][0].product.name == "商品A"
    assert result[2][0].product.name == "商品D"


def test_place_labels():
    product = create_test_product()

    items = [
        LabelItem(
            product=product,
            branch_name=BRANCH_A,
            quantity=7,
        )
        for _ in range(23)
    ]

    placements = place_labels(
        label_items=items,
        start_position=20,
    )

    assert placements[0].sheet_number == 1
    assert placements[0].position == 20

    assert placements[1].position == 21
    assert placements[2].sheet_number == 2
    assert placements[2].position == 1

    assert placements[-1].sheet_number == 2
    assert placements[-1].position == 21


def test_calculate_next_position():
    assert calculate_next_position(
        start_position=1,
        printed_count=1,
    ) == 2

    assert calculate_next_position(
        start_position=21,
        printed_count=1,
    ) == 1

    assert calculate_next_position(
        start_position=20,
        printed_count=23,
    ) == 1


def test_split_label_order_zero_quantity():
    product = create_test_product()

    order = LabelOrder(
        product=product,
        branch_name=BRANCH_A,
        quantity=0,
    )

    assert split_label_order(order) == []

def test_place_labels_invalid_start_position():
    with pytest.raises(ValueError):
        place_labels([], start_position=0)

    with pytest.raises(ValueError):
        place_labels([], start_position=22)


def test_build_label_placements():
    product = create_test_product()

    branch_orders = {
        BRANCH_A: {
            "商品A": 20,
        }
    }

    product_master = {
        "商品A": product,
    }

    placements = build_label_placements(
        branch_orders=branch_orders,
        product_master=product_master,
        floor=1,
        start_position=1,
    )

    assert len(placements) == 3

    assert placements[0].position == 1
    assert placements[1].position == 2
    assert placements[2].position == 3

    assert [placement.item.quantity for placement in placements] == [
        7,
        7,
        6,
    ]