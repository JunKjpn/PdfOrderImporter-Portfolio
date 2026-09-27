from datetime import date
from pathlib import Path

from src.constants.excel import SUMMARY_FILE_NAME
from src.constants.paths import OUTPUT_DIR
from src.constants.branches import BRANCH_A
from src.excel.summary import load_branch_colors
from src.models.label import LabelItem
from src.models.product import Product
from src.services.label_pdf import create_test_sheet_pdf

PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = PROJECT_ROOT / "templates" / SUMMARY_FILE_NAME


product = Product(
    product_id=1,
    name="商品A",
    units_per_box=7,
    packaging_floor=1,
)

item = LabelItem(
    product=product,
    branch_name=BRANCH_A,
    quantity=7,
)
items = [item] * 21

branch_colors = load_branch_colors(TEMPLATE_PATH)

create_test_sheet_pdf(
    output_path=OUTPUT_DIR / "test_sheet.pdf",
    items=items,
    delivery_date=date(2026, 8, 19),
    branch_colors=branch_colors,
)