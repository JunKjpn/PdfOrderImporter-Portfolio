from dataclasses import dataclass


@dataclass(frozen=True)
class Product:
    product_id: int
    name: str
    units_per_box: int
    packaging_floor: int
    order_multiplier: int
    store_sales_units: int
    unit_price: int
    margin_unit_price: float