from pathlib import Path

from src.constants.branches import BRANCH_A
from src.services.label_pdf import create_test_label_pdf, create_test_sheet_pdf
from tests.helpers import create_test_product


def test_create_test_label_pdf(tmp_path: Path):
    product = create_test_product()

    item = LabelItem(
        product=product,
        branch_name=BRANCH_A,
        quantity=7,
    )

    output_path = tmp_path / "test_label.pdf"

    create_test_label_pdf(
        output_path=output_path,
        item=item,
        delivary_date=date(2026, 8, 19),
        branch_color=(255, 204, 255)
    )

    assert output_path.exists()
    assert output_path.stat().st_size > 0


def test_create_label_sheet_pdf(tmp_path: Path):
    product = create_test_product()

    item = LabelItem(
        product=product,
        branch_name=BRANCH_A,
        quantity=7,
    )

    items = [item] * 21

    branch_colors = {
        BRANCH_A: (255, 192, 0),
    }

    output_path = tmp_path / "test_label_sheet.pdf"

    create_test_sheet_pdf(
        output_path=output_path,
        items=items,
        delivery_date=date(2026, 8, 19),
        branch_colors=branch_colors,
    )

    assert output_path.exists()
    assert output_path.stat().st_size > 0

from datetime import date
from pathlib import Path

from src.models.label import LabelItem, LabelPlacement
from src.models.product import Product
from src.services.label_pdf import create_label_pdf


def test_create_label_pdf(tmp_path: Path):
    product = create_test_product()

    item = LabelItem(
        product=product,
        branch_name=BRANCH_A,
        quantity=7,
    )

    # 23枚 → 1ページ目21枚 + 2ページ目2枚
    placements = [
        LabelPlacement(
            item=item,
            sheet_number=1 if index < 21 else 2,
            position=index + 1 if index < 21 else index - 20,
        )
        for index in range(23)
    ]

    branch_colors = {
        BRANCH_A: (255, 192, 0),
    }

    output_path = tmp_path / "test_label.pdf"

    create_label_pdf(
        output_path=output_path,
        placements=placements,
        delivery_date=date(2026, 8, 19),
        branch_colors=branch_colors,
    )

    assert output_path.exists()
    assert output_path.stat().st_size > 0