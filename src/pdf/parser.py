import re
from pathlib import Path

import pdfplumber

from src.constants.products import PRODUCT_ALIASES
from src.utils.converter import normalize_text

IGNORED_PRODUCT_NAMES = {
    "合計",
}


def extract_lines(pdf_path: Path) -> list[str]:
    """PDF全ページからテキストを抽出し、行単位で返す。"""
    lines: list[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            if not text:
                continue

            lines.extend(text.splitlines())

    return lines


def extract_order_items(pdf_path: Path) -> dict[str, int]:
    """
    PDFから商品名と数量を抽出する。
    数量が空欄の商品は0とする。
    同じ商品が複数回出た場合は数量を加算する。
    """
    lines = extract_lines(pdf_path)
    orders: dict[str, int] = {}

    for line in lines:
        stripped_line = line.strip()

        match = re.match(r"^(.+?)\s+(\d+)(?:\s|$)", stripped_line)

        if match is None:
            continue

        product_name = normalize_text(match.group(1))

        if product_name in IGNORED_PRODUCT_NAMES:
            continue

        # PDF上の商品名を商品マスタ上の商品名へ変換
        product_name = PRODUCT_ALIASES.get(
            product_name,
            product_name,
        )

        quantity = int(match.group(2))

        orders[product_name] = orders.get(product_name, 0) + quantity

    return orders