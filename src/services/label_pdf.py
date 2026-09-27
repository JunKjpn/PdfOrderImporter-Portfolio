from datetime import date
from pathlib import Path

from reportlab.lib.colors import Color
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from src.models.label import LabelItem, LabelPlacement

FONT_PATH = r"C:\Windows\Fonts\YuGothM.ttc"
pdfmetrics.registerFont(TTFont("YuGothM", FONT_PATH))

MM_TO_PT = 72 / 25.4

A4_WIDTH = 210.0 * MM_TO_PT
A4_HEIGHT = 297.0 * MM_TO_PT

LABEL_COLUMNS = 3
LABEL_ROWS = 7
LABEL_COUNT = LABEL_COLUMNS * LABEL_ROWS
LABEL_WIDTH = 70.0 * MM_TO_PT
LABEL_HEIGHT = 42.4 * MM_TO_PT

PRODUCT_X = 5.0 * MM_TO_PT
PRODUCT_Y = 32.0 * MM_TO_PT

BRANCH_X = 5.0 * MM_TO_PT
BRANCH_Y = 23.0 * MM_TO_PT

DATE_X = 5.0 * MM_TO_PT
DATE_Y = 14.0 * MM_TO_PT

QUANTITY_X = 5.0 * MM_TO_PT
QUANTITY_Y = 5.0 * MM_TO_PT

PRODUCT_FONT_SIZE = 18
PRODUCT_MIN_FONT_SIZE = 10
PRODUCT_MAX_WIDTH = 60.0 * MM_TO_PT


def draw_label(
    pdf: canvas.Canvas,
    item: LabelItem,
    delivaty_date: date,
    branch_color: tuple[int, int, int] | None,
    x: float,
    y: float,
) -> None:
    """
    1枚のラベルをPDF上に描画する。
    x, y はラベル左下の座標。
    """

    # 商品名
    product_name = item.product.name
    font_size = PRODUCT_FONT_SIZE
    while (
            font_size > PRODUCT_MIN_FONT_SIZE
            and pdfmetrics.stringWidth(product_name, "YuGothM", font_size) > PRODUCT_MAX_WIDTH
    ):
        font_size -= 1
    pdf.setFont("YuGothM", font_size)
    pdf.drawString(x + PRODUCT_X, y + PRODUCT_Y, product_name)

    # 営業所名
    if branch_color is not None:
        pdf.setFillColor(Color(branch_color[0] / 255, branch_color[1] / 255, branch_color[2] / 255))
        pdf.rect(x + BRANCH_X, y + (BRANCH_Y - 3), 60 * MM_TO_PT, 6 * MM_TO_PT, fill=1, stroke=0)
        pdf.setFillColorRGB(0, 0, 0)

    branch_text = f"営業所 ： {item.branch_name}"
    pdf.setFont("YuGothM", 14)
    pdf.drawString(x + BRANCH_X, y + BRANCH_Y, branch_text)

    # 販売日
    sales_date_text = f"日付け ： {delivaty_date.strftime("%Y/%m/%d")}"
    pdf.setFont("YuGothM", 14)
    pdf.drawString(x + DATE_X, y + DATE_Y, sales_date_text)

    # 入り数
    quantity_text = f"入り数（   {item.quantity}   ）個"
    pdf.setFont("YuGothM", 14)
    pdf.drawString(x + QUANTITY_X, y + QUANTITY_Y, quantity_text)


def create_test_label_pdf(
    output_path: Path,
    item: LabelItem,
    delivary_date: date,
    branch_color: tuple[int, int, int] | None,
) -> None:
    """
    1枚のラベルを確認するためのテストPDFを生成する。
    """

    pdf = canvas.Canvas(str(output_path), pagesize=(LABEL_WIDTH, LABEL_HEIGHT))
    draw_label(pdf=pdf, item=item, delivaty_date=delivary_date, branch_color=branch_color, x=0, y=0)
    pdf.save()


def create_test_sheet_pdf(
    output_path: Path,
    items: list[LabelItem],
    delivery_date: date,
    branch_colors: dict[str, tuple[int, int, int] | None],
) -> None:
    """
    A4の21面ラベルPDFを生成する。

    items:
        ラベルに印刷するLabelItemの一覧。
        最大21枚まで。
    """

    if len(items) > LABEL_COUNT:
        raise ValueError(
            f"1ページに配置できるラベルは最大{LABEL_COUNT}枚です。\n\n"
            f"対象枚数：{len(items)}"
        )

    pdf = canvas.Canvas(str(output_path), pagesize=(A4_WIDTH, A4_HEIGHT))

    for index, item in enumerate(items):
        column = index % LABEL_COLUMNS
        row = index // LABEL_COLUMNS

        x = column * LABEL_WIDTH
        y = A4_HEIGHT - (row + 1) * LABEL_HEIGHT

        branch_color = branch_colors.get(item.branch_name)

        draw_label(
            pdf=pdf,
            item=item,
            delivaty_date=delivery_date,
            branch_color=branch_color,
            x=x,
            y=y,
        )

    pdf.save()

def get_label_position(position: int) -> tuple[float, float]:
    """
    21面シート上の位置番号から、PDF上のx,y座標を取得する。
    positionは1～21。
    """
    if not 1 <= position <= 21:
        raise ValueError(f"ラベル位置は1～21で指定してください：{position}")

    index = position - 1

    column = index % LABEL_COLUMNS
    row = index // LABEL_COLUMNS

    x = column * LABEL_WIDTH
    y = A4_HEIGHT - (row + 1) * LABEL_HEIGHT

    return x, y

def create_label_pdf(
    output_path: Path,
    placements: list[LabelPlacement],
    delivery_date: date,
    branch_colors: dict[str, tuple[int, int, int] | None],
) -> None:
    """
    LabelPlacementをA4 21面PDFとして出力する。
    """
    if not placements:
        raise ValueError("印刷対象となるラベルがありません。")

    pdf = canvas.Canvas(str(output_path), pagesize=(A4_WIDTH, A4_HEIGHT))

    current_sheet_number = 1


    for placement in placements:
        if placement.sheet_number != current_sheet_number:
            pdf.showPage()
            current_sheet_number = placement.sheet_number

        x, y = get_label_position(placement.position)

        branch_color = branch_colors.get(placement.item.branch_name)

        draw_label(
            pdf=pdf,
            item=placement.item,
            delivaty_date=delivery_date,
            branch_color=branch_color,
            x=x,
            y=y,
        )

    pdf.save()

def cleanup_old_label_pdfs(output_dir: Path, floor: int, keep_count: int = 3) -> None:
    """
    指定階のラベルPDFがkeep_countを超えた場合、
    最も古いPDFを削除する。
    """
    pdf_paths = list(output_dir.glob(f"ラベル_{floor}F_*.pdf"))
    if len(pdf_paths) <= keep_count:
        return
    pdf_paths.sort(key=lambda path: path.stem.rsplit("_", 1)[-1])
    for pdf_path in pdf_paths[:-keep_count]:
        pdf_path.unlink()