from src.models.product import Product


def validate_product_names(branch_orders: dict[str, dict[str, int]], product_master: dict[str, Product]) -> None:
    """
    PDFから取得した商品名が商品マスタに存在するか確認する。
    """
    for orders in branch_orders.values():
        for product_name in orders:
            if product_name not in product_master:
                raise ValueError(
                    f"商品マスタに登録されていない商品が発注PDFに含まれています。\n\n"
                    f"商品名：{product_name}"
                )