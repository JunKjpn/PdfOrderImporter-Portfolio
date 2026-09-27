import re
from datetime import date, datetime

import unicodedata

CELL_REFERENCE_PATTERN = re.compile(r"^=\$?([A-Z]{1,3})\$?(\d+)$", re.IGNORECASE)

def normalize_amount(value: object) -> int | float:
    """
    Excelから取得した金額を数値へ変換する。
    空欄は0として扱う。
    """
    if value is None or value == "":
        return 0

    if isinstance(value, bool):
        raise ValueError(f"金額欄に不正な値があります：{value}")

    if isinstance(value, (int, float)):
        return value

    value_text = str(value).replace(",", "").replace("￥", "").replace("¥", "").strip()

    if not value_text:
        return 0

    try:
        number = float(value_text)
    except ValueError as error:
        raise ValueError(f"金額を数値として読み取れませんでした。\n\n対象値：{value}") from error

    if number.is_integer():
        return int(number)

    return number

def normalize_date(value: object) -> date | None:
    """
    日付セルをdate型へ変換する。
    """
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    return None

def normalize_int(value: object) -> int | None:
    """
    Excelから取得した値を正の整数へ変換する。
    """
    if value is None or value == "":
        return None

    normalized_value = normalize_text(value)

    try:
        number = int(normalized_value)
    except ValueError as error:
        raise ValueError(f"整数として読み取れない値です。\n\n対象値：{value}") from error

    return number

def normalize_float(value: object) -> float:
    """値をfloatへ正規化する。"""
    if value is None:
        raise ValueError("数値が未入力です。")

    if isinstance(value, bool):
        raise ValueError(f"不正な数値です：{value}")

    if isinstance(value, (int, float)):
        return float(value)

    normalized_value = normalize_text(value)

    if not normalized_value:
        raise ValueError("数値が未入力です。")

    try:
        return float(normalized_value)
    except ValueError as error:
        raise ValueError(
            f"数値として解釈できません：{value}"
        ) from error

def normalize_positive_float(value: object) -> float:
    """正のfloatへ正規化する。"""
    normalized_value = normalize_float(value)

    if normalized_value <= 0:
        raise ValueError(
            f"正の数値を指定してください：{value}"
        )

    return normalized_value

def normalize_text(text: object) -> str:
    """
    文字列比較用に文字列を正規化する。
    """
    if text is None:
        return ""

    text = unicodedata.normalize("NFKC", str(text))
    text = text.replace("\u00a0", " ").replace("\u3000", " ").replace(" ", "")
    text = text.replace("　", "").replace("（", "(").replace("）", ")").strip()
    return text
