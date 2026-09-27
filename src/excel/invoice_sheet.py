import re
from datetime import date

from openpyxl.workbook.workbook import Workbook

# 請求書側のシート名パターン
INVOICE_SHEET_PATTERN = re.compile(
    r"^R(?P<era_year>\d+)"
    r"[.．]"
    r"(?P<start_month>\d+)月"
    r"(?P<start_day>\d+)日?"
    r"\s*[~～〜\-－]\s*"
    r"(?:(?P<end_month>\d+)月)?"
    r"(?P<end_day>\d+)日?$"
)


def parse_invoice_sheet_period(sheet_name: str) -> tuple[date, date] | None:
    """
    請求書のシート名から対象期間を取得する。

    例:
        R8.7月10~7月20
        R8.7月21～31
    """
    normalized_name = sheet_name.strip()

    match = INVOICE_SHEET_PATTERN.fullmatch(normalized_name)

    if match is None:
        return None

    era_year = int(match.group("era_year"))

    start_month = int(match.group("start_month"))
    start_day = int(match.group("start_day"))

    end_month_text = match.group("end_month")
    end_month = (int(end_month_text) if end_month_text else start_month)

    end_day = int(match.group("end_day"))

    # 令和元年は2019年なので、令和年 + 2018
    western_year = era_year + 2018

    try:
        start_date = date(western_year, start_month, start_day)

        end_year = western_year

        # 12月末～1月初など、年をまたぐ場合
        if end_month < start_month:
            end_year += 1

        end_date = date(end_year, end_month, end_day)

    except ValueError as error:
        raise ValueError(
            "請求書のシート名に不正な日付があります。\n\n"
            f"シート名：{sheet_name}"
        ) from error

    return start_date, end_date

def find_invoice_sheet_name(workbook: Workbook, delivery_date: date) -> str:
    """
    納品日が対象期間内に入る請求書シートを取得する。
    """
    matched_sheet_names: list[str] = []

    for sheet_name in workbook.sheetnames:
        period = parse_invoice_sheet_period(sheet_name=sheet_name)

        # 対象形式ではないシートは無視
        if period is None:
            continue

        start_date, end_date = period

        if start_date <= delivery_date <= end_date:
            matched_sheet_names.append(sheet_name)

    if not matched_sheet_names:
        raise ValueError(
            "納品日に対応する請求書シートが見つかりません。\n\n"
            f"納品日：{delivery_date:%Y/%m/%d}\n"
            "シート名が次の形式になっているか確認してください。\n"
            "例：R8.7月10~7月20"
        )

    if len(matched_sheet_names) > 1:
        raise ValueError(
            "納品日を含む請求書シートが複数見つかりました。\n\n"
            f"納品日：{delivery_date:%Y/%m/%d}\n"
            f"対象シート：{', '.join(matched_sheet_names)}"
        )

    return matched_sheet_names[0]