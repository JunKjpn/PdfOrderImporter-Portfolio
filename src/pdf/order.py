from datetime import date
from pathlib import Path

from src.pdf.filename import parse_filename
from src.pdf.parser import extract_order_items


def load_pdf_orders(pdf_paths: list[Path], branch_names: set[str]) -> tuple[date, dict[str, dict[str, int]]]:
    """
    複数の発注PDFを読み込み、納品日と営業所別の発注数量を取得する。
    """
    base_delivery_date = None
    branch_orders: dict[str, dict[str, int]] = {}

    for pdf_path in pdf_paths:
        delivery_date, branch_name = parse_filename(pdf_path)

        if base_delivery_date is None:
            base_delivery_date = delivery_date

        elif delivery_date != base_delivery_date:
            raise ValueError(
                f"納品日が異なるPDFが含まれています。\n\n基準日：{base_delivery_date}\n対象ファイル：{pdf_path.name}\n"
                f"対象日：{delivery_date}"
            )

        if branch_name not in branch_names:
            raise ValueError(
                f"選択したExcelに営業所の列が見つかりません。\n\n営業所：{branch_name}\nPDFファイル：{pdf_path.name}"
            )

        orders = extract_order_items(pdf_path=pdf_path)

        if branch_name not in branch_orders:
            branch_orders[branch_name] = {}

        for product_name, quantity in orders.items():
            current_quantity = branch_orders[branch_name].get(product_name,0)

            branch_orders[branch_name][product_name] = current_quantity + quantity

    if base_delivery_date is None:
        raise ValueError("PDFから納品日を取得できませんでした。")

    return base_delivery_date, branch_orders