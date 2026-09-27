from pathlib import Path

from src.pdf.parser import extract_order_items

PDF_PATH = Path(__file__).parents[1] / "sample/260922営業所C（サンプル）.pdf"

def test_extract_order_items_vegetable_roll():
    result = extract_order_items(pdf_path=PDF_PATH)

    assert result["商品C(バリエーション2)"] == 9