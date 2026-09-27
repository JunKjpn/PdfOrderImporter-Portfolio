from src.models.product import Product


def create_test_product(
    product_id: int = 1,
    name: str = "商品A",
    units_per_box: int = 7,
    packaging_floor: int = 1,
) -> Product:
    return Product(
        product_id=product_id,
        name=name,
        units_per_box=units_per_box,
        packaging_floor=packaging_floor,
        order_multiplier=1,
        store_sales_units=1,
        unit_price=100,
        margin_unit_price=50.0,
    )