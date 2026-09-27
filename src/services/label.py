from pathlib import Path

from src.excel.product_master import load_product_master
from src.excel.summary import load_branch_colors
from src.pdf.order import load_pdf_orders
from src.services.label_pdf import create_label_pdf
from src.models.label import LabelOrder, LabelItem, LabelPlacement
from src.models.product import Product


LABEL_POSITIONS = 21

def build_label_orders(
    branch_orders: dict[str, dict[str, int]],
    product_master: dict[str, Product],
) -> list[LabelOrder]:
    """
    PDFから取得した営業所別・商品別の注文数量と商品マスタを紐付けて、ラベル発行用データを作成する。
    """
    label_orders: list[LabelOrder] = []

    for branch_name, product_orders in branch_orders.items():
        for product_name, quantity in product_orders.items():
            if quantity <= 0:
                continue

            if product_name not in product_master:
                raise ValueError(
                    "商品マスタに登録されていない商品があります。\n\n"
                    f"商品名：{product_name}"
                )

            product = product_master[product_name]

            label_orders.append(
                LabelOrder(
                    product=product,
                    branch_name=branch_name,
                    quantity=quantity,
                )
            )

    return label_orders


def split_label_order(order: LabelOrder) -> list[LabelItem]:
    """
    1件の受注データを、ラベル1枚ごとの数量に分割する。
    """
    units_per_box = order.product.units_per_box
    quantity = order.quantity

    if quantity <= 0:
        return []

    full_labels, remainder = divmod(quantity, units_per_box)

    label_items = [
        LabelItem(product=order.product, branch_name=order.branch_name, quantity=units_per_box)
        for _ in range(full_labels)
    ]

    if remainder > 0:
        label_items.append(
            LabelItem(product=order.product, branch_name=order.branch_name, quantity=remainder)
        )

    return label_items


def place_labels(label_items: list[LabelItem], start_position: int = 1) -> list[LabelPlacement]:
    """
    ラベルを21面シート上に順番に配置する。
    """
    if not 1 <= start_position <= 21:
        raise ValueError(f"ラベル開始位置は1～21で指定してください。: {start_position}")

    placements: list[LabelPlacement] = []
    position = start_position
    sheet_number = 1

    for item in label_items:
        placements.append(
            LabelPlacement(item=item, sheet_number=sheet_number, position=position)
        )

        position += 1
        if position > LABEL_POSITIONS:
            position = 1
            sheet_number += 1

    return placements


def calculate_next_position(start_position: int, printed_count: int) -> int:
    """
    印刷後の次回ラベル開始位置を計算する。
    """
    if not 1 <= start_position <= LABEL_POSITIONS:
        raise ValueError(f"開始位置は1～{LABEL_POSITIONS}で指定してください。")

    if printed_count < 0:
        raise ValueError("印刷枚数は0以上で指定してください。")

    return (start_position - 1 + printed_count) % LABEL_POSITIONS + 1


def split_labels_by_floor(label_items: list[LabelItem]) -> dict[int, list[LabelItem]]:
    """
    ラベルを包装階ごとに分ける。
    """
    labels_by_floor: dict[int, list[LabelItem]] = {
        1: [],
        2: [],
    }

    for item in label_items:
        floor = item.product.packaging_floor

        if floor not in labels_by_floor:
            raise ValueError(f"ラベル包装階が不正です：{floor}")

        labels_by_floor[floor].append(item)

    return labels_by_floor


def build_label_placements(
    branch_orders: dict[str, dict[str, int]],
    product_master: dict[str, Product],
    floor: int,
    start_position: int = 1,
) -> list[LabelPlacement]:
    """
    営業所別・商品別の受注数量から、指定した包装階のラベル配置情報を作成する。
    """

    label_orders = build_label_orders(
        branch_orders=branch_orders,
        product_master=product_master,
    )

    label_items: list[LabelItem] = []

    for order in label_orders:
        items = split_label_order(order)

        for item in items:
            if item.product.packaging_floor == floor:
                label_items.append(item)

    return place_labels(
        label_items=label_items,
        start_position=start_position,
    )

def create_label_pdf_output(
        excel_path: Path,
        pdf_paths: list[Path],
        floor: int,
        start_position: int,
        output_dir: Path,
) -> tuple[Path, int]:
    """発注PDFから指定階のラベルPDFを作成する。"""

    label_product_master = load_product_master(path=excel_path)

    branch_colors = load_branch_colors(excel_path)

    delivery_date, branch_orders = load_pdf_orders(
        pdf_paths=pdf_paths,
        branch_names=set(branch_colors),
    )

    placements = build_label_placements(
        branch_orders=branch_orders,
        product_master=label_product_master,
        floor=floor,
        start_position=start_position,
    )

    output_path = output_dir / f"ラベル_{floor}F_{delivery_date:%Y%m%d}.pdf"

    create_label_pdf(
        output_path=output_path,
        placements=placements,
        delivery_date=delivery_date,
        branch_colors=branch_colors,
    )

    return output_path, len(placements)