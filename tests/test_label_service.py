import pytest

from src.constants.branches import BRANCH_A, BRANCH_B
from src.models.label import LabelOrder, LabelItem
from src.services.label import build_label_orders, split_label_order, place_labels, calculate_next_position, \
    split_labels_by_floor
from tests.helpers import create_test_product


def test_build_label_orders():
    product_master = {
        "商品A": create_test_product(),
        "商品D": create_test_product(
            product_id=7,
            name="商品D",
            units_per_box=12,
            packaging_floor=2,
        ),
    }

    branch_orders = {
        BRANCH_A: {
            "商品A": 17,
            "商品D": 24,
        },
        BRANCH_B: {
            "商品A": 14,
        },
    }

    label_orders = build_label_orders(
        branch_orders=branch_orders,
        product_master=product_master,
    )

    assert len(label_orders) == 3

    assert label_orders[0].product.name == "商品A"
    assert label_orders[0].branch_name == BRANCH_A
    assert label_orders[0].quantity == 17
    assert label_orders[0].product.units_per_box == 7
    assert label_orders[0].product.packaging_floor == 1

    assert label_orders[1].product.name == "商品D"
    assert label_orders[1].branch_name == BRANCH_A
    assert label_orders[1].quantity == 24
    assert label_orders[1].product.packaging_floor == 2


def test_build_label_orders_unknown_product():
    product_master = {
        "商品A": create_test_product()
    }

    branch_orders = {
        BRANCH_A: {
            "存在しない商品": 10,
        },
    }

    with pytest.raises(ValueError, match="商品マスタに登録されていない商品"):
        build_label_orders(
            branch_orders=branch_orders,
            product_master=product_master,
        )


def test_split_label_order():
    product = create_test_product()

    order = LabelOrder(
        product=product,
        branch_name=BRANCH_A,
        quantity=17,
    )

    label_items = split_label_order(order)

    assert [item.quantity for item in label_items] == [7, 7, 3]


def test_place_labels():
    product = create_test_product()

    items = [
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 3),
    ]

    placements = place_labels(items)

    assert [placement.position for placement in placements] == [1, 2, 3]


def test_place_labels_from_middle():
    product = create_test_product()

    items = [
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 3),
    ]

    placements = place_labels(
        items,
        start_position=19,
    )

    assert [placement.position for placement in placements] == [19, 20, 21]


def test_place_labels_wraps_to_next_sheet():
    product = create_test_product()

    items = [
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 3),
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 7),
    ]

    placements = place_labels(
        items,
        start_position=19,
    )

    assert [placement.position for placement in placements] == [
        19, 20, 21, 1, 2
    ]


def test_calculate_next_position():
    assert calculate_next_position(1, 5) == 6
    assert calculate_next_position(18, 5) == 2
    assert calculate_next_position(1, 21) == 1
    assert calculate_next_position(5, 21) == 5


def test_calculate_next_position_invalid_start():
    with pytest.raises(ValueError):
        calculate_next_position(0, 5)

    with pytest.raises(ValueError):
        calculate_next_position(22, 5)


def test_calculate_next_position_negative_count():
    with pytest.raises(ValueError):
        calculate_next_position(1, -1)


def test_place_labels_tracks_sheet_number():
    product = create_test_product()

    items = [
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 3),
        LabelItem(product, BRANCH_A, 7),
        LabelItem(product, BRANCH_A, 7),
    ]

    placements = place_labels(items, start_position=19)

    assert [
               (placement.sheet_number, placement.position)
               for placement in placements
           ] == [
               (1, 19),
               (1, 20),
               (1, 21),
               (2, 1),
               (2, 2),
           ]


def test_split_labels_by_floor():
    floor_1 = create_test_product()
    floor_2 = create_test_product(
        product_id=7,
        name="商品D",
        units_per_box=12,
        packaging_floor=2,
    )

    items = [
        LabelItem(floor_1, BRANCH_A, 7),
        LabelItem(floor_2, BRANCH_A, 12),
        LabelItem(floor_1, BRANCH_A, 3),
    ]

    result = split_labels_by_floor(items)

    assert len(result[1]) == 2
    assert len(result[2]) == 1
