from dataclasses import dataclass

from src.models.product import Product


@dataclass(frozen=True)
class LabelOrder:
    product: Product
    branch_name: str
    quantity: int


@dataclass(frozen=True)
class LabelItem:
    product: Product
    branch_name: str
    quantity: int


@dataclass(frozen=True)
class LabelPlacement:
    item: LabelItem
    sheet_number: int
    position: int


@dataclass(frozen=True)
class LabelPrintState:
    floor: int
    next_position: int
