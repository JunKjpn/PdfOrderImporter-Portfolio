from pathlib import Path

from src.constants.excel import SUMMARY_FILE_NAME
from src.constants.paths import OUTPUT_DIR
from src.excel.product_master import load_product_master
from src.excel.summary import load_branch_colors
from src.pdf.order import load_pdf_orders
from src.services.label import build_label_placements
from src.services.label_pdf import create_label_pdf


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = PROJECT_ROOT / "templates"
PDF_DIR = PROJECT_ROOT / "sample/pdf"

product_master = load_product_master(TEMPLATE_DIR / SUMMARY_FILE_NAME)
branch_colors = load_branch_colors(TEMPLATE_DIR / SUMMARY_FILE_NAME)
pdf_paths = list(PDF_DIR.glob("*.pdf"))

delivery_date, branch_orders = load_pdf_orders(
    pdf_paths=pdf_paths,
    product_names=list(product_master),
    branch_names=set(branch_colors),
)

placements_1f = build_label_placements(
    branch_orders=branch_orders,
    product_master=product_master,
    floor=1,
    start_position=1,
)

placements_2f = build_label_placements(
    branch_orders=branch_orders,
    product_master=product_master,
    floor=2,
    start_position=1,
)

create_label_pdf(
    output_path=OUTPUT_DIR / "label_1f.pdf",
    placements=placements_1f,
    delivery_date=delivery_date,
    branch_colors=branch_colors,
)

create_label_pdf(
    output_path=OUTPUT_DIR / "label_2f.pdf",
    placements=placements_2f,
    delivery_date=delivery_date,
    branch_colors=branch_colors,
)