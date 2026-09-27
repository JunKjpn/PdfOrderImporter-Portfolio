from datetime import date
from pathlib import Path
import xml.etree.ElementTree as ET

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEX

from src.constants.excel import TRAY_ORDER_MULTIPLIER_LABEL, TRAY_HEADER_ROW, \
    TRAY_STORE_SALES_UNITS_LABEL, PROFIT_UNIT_PRICE_LABEL, PROFIT_HEADER_ROW, PROFIT_START_ROW, \
    PROFIT_MARGIN_UNIT_PRICE_LABEL, TRAY_START_ROW, TRAY_SHEET_NAME, PROFIT_CALCULATION_SHEET_NAME, \
    BOX_CONVERSION_SHEET_NAME, BOX_UNITS_PER_BOX_LABEL, BOX_HEADER_ROW, BOX_START_ROW, BOX_TABLE_SHEET_NAME, \
    SUMMARY_SHEET_NAME, SUMMARY_HEADER_ROW, SUMMARY_PRODUCT_COLUMN, SUMMARY_START_ROW, PRODUCT_MASTER_SHEET_NAME
from src.models.product import Product
from src.constants.branches import BRANCHES
from src.excel.finder import find_branch_columns, find_row_by_value
from src.excel.writer import open_workbook, write_orders, save_workbook, write_product_names, clear_cell_range, \
    write_product_master_value, write_box_table_units
from src.excel.product_master import product_master_to_dataframe



def load_branch_colors(summary_path: Path) -> dict[str, tuple[int, int, int] | None]:
    """
    集計表の営業所ヘッダーからラベル用の色を取得する。
    戻り値:
    {
        "営業所A": (255, 192, 0),
        "営業所B": (146, 208, 80),
        ...
    }
    """
    workbook = load_workbook(filename=summary_path, data_only=False, read_only=False)

    try:
        if SUMMARY_SHEET_NAME not in workbook.sheetnames:
            raise ValueError(
                f"集計表に必要なシートが見つかりません。\n\n"
                f"必要なシート：{SUMMARY_SHEET_NAME}"
            )

        worksheet = workbook[SUMMARY_SHEET_NAME]

        branch_colors: dict[str, tuple[int, int, int] | None] = {}

        branch_columns = find_branch_columns(
            worksheet=worksheet,
            branch_names=BRANCHES,
            header_row=SUMMARY_HEADER_ROW,
        )

        for branch_name, column_number in branch_columns.items():
            cell = worksheet.cell(
                row=SUMMARY_HEADER_ROW,
                column=column_number,
            )

            fill = cell.fill

            if fill.fill_type is None:
                branch_colors[branch_name] = None
                continue

            color = fill.fgColor

            if color.type == "rgb":
                rgb = color.rgb

                if rgb is None:
                    raise ValueError(
                        f"営業所の色を取得できませんでした。\n\n"
                        f"営業所：{branch_name}"
                    )

                # ARGBの場合は先頭のアルファ値を除く
                if len(rgb) == 8:
                    rgb = rgb[2:]

                branch_colors[branch_name] = (
                    int(rgb[0:2], 16),
                    int(rgb[2:4], 16),
                    int(rgb[4:6], 16),
                )

            elif color.type == "indexed":
                rgb = COLOR_INDEX[color.indexed]

                branch_colors[branch_name] = (
                    int(rgb[2:4], 16),
                    int(rgb[4:6], 16),
                    int(rgb[6:8], 16),
                )

            elif color.type == "theme":
                branch_colors[branch_name] = resolve_theme_color(
                    workbook=workbook,
                    theme_index=color.theme,
                    tint=color.tint or 0.0,
                )

            else:
                raise ValueError(
                    f"営業所の色形式に対応していません。\n\n"
                    f"営業所：{branch_name}\n"
                    f"色形式：{color.type}"
                )

        return branch_colors

    finally:
        workbook.close()

def apply_tint(red: int, green: int, blue: int, tint: float) -> tuple[int, int, int]:
    """
    Excelのtint値をRGBに反映する。
    """
    def adjust(value: int) -> int:
        if tint < 0:
            value = value * (1.0 + tint)
        else:
            value = value * (1.0 - tint) + 255.0 * tint

        return max(0, min(255, round(value)))

    return (
        adjust(red),
        adjust(green),
        adjust(blue),
    )

def resolve_theme_color(workbook, theme_index: int, tint: float) -> tuple[int, int, int]:
    """
    ExcelのテーマカラーをRGBへ変換する。
    """
    if workbook.loaded_theme is None:
        raise ValueError("Excelファイルにテーマ情報がありません。")

    root = ET.fromstring(workbook.loaded_theme)

    namespace = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }

    color_scheme = root.find(".//a:clrScheme", namespace)

    if color_scheme is None:
        raise ValueError("Excelファイルからカラーテーマを取得できませんでした。")

    theme_colors = []

    for child in color_scheme:
        color_element = next(iter(child), None)

        if color_element is None:
            continue

        value = color_element.get("lastClr") or color_element.get("val")

        if value is None:
            raise ValueError(
                f"テーマカラーのRGB値を取得できませんでした：{child.tag}"
            )

        theme_colors.append(value)

    if theme_index < 0 or theme_index >= len(theme_colors):
        raise ValueError(
            f"テーマカラー番号が範囲外です：{theme_index}"
        )

    rgb = theme_colors[theme_index]

    red = int(rgb[0:2], 16)
    green = int(rgb[2:4], 16)
    blue = int(rgb[4:6], 16)

    return apply_tint(red, green, blue, tint)

def orders_to_dataframe(orders: dict[str, int]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"商品名": product_name, "数量": quantity}
            for product_name, quantity in orders.items()
        ]
    )

def merge_product_orders(base_df: pd.DataFrame, orders: dict[str, int]) -> pd.DataFrame:
    """基準となる商品一覧に営業所の発注数量を結合する。注文がない商品は数量を0とする。"""
    orders_df = orders_to_dataframe(orders)
    merged_df = pd.merge(base_df[["商品名"]], orders_df, on="商品名", how="left")
    merged_df["数量"] = merged_df["数量"].fillna(0).astype(int)
    return merged_df

def merge_all_orders(product_master: dict[str, Product], branch_orders: dict[str, dict[str, int]]) -> pd.DataFrame:
    """全営業所で今回注文された商品を、商品マスタの順番に整列する。"""
    all_orders: dict[str, int] = {}

    for orders in branch_orders.values():
        for product_name in orders:
            all_orders[product_name] = 0

    product_master_df = product_master_to_dataframe(product_master)
    orders_df = orders_to_dataframe(all_orders)

    return pd.merge(product_master_df, orders_df, on="商品名", how="inner")[["商品名"]]

def create_summary(
    excel_path: Path,
    sheet_name: str,
    product_master: dict[str, Product],
    branch_columns: dict[str, int],
    branch_orders: dict[str, dict[str, int]],
    delivery_date: date,
    output_dir: Path,
) -> Path:
    """発注内容を集計表へ転記し、納品日を付けた別名ファイルとして保存する。"""
    # 選択されたExcelを開いて、各worksheetを取得
    workbook, worksheet = open_workbook(
        template_path=excel_path,
        sheet_name=sheet_name,
    )
    box_worksheet = workbook[BOX_CONVERSION_SHEET_NAME]
    tray_worksheet = workbook[TRAY_SHEET_NAME]
    profit_worksheet = workbook[PROFIT_CALCULATION_SHEET_NAME]
    box_table_worksheet = workbook[BOX_TABLE_SHEET_NAME]

    try:
        # 今回注文された商品の一覧を商品マスタ順で取得
        merged_all_df = merge_all_orders(
            product_master=product_master,
            branch_orders=branch_orders,
        )

        # 今回注文された商品名の一覧を取得
        product_names = merged_all_df["商品名"].tolist()

        # シート内の合計が記入されている行番号を取得
        total_row = find_row_by_value(
            worksheet=worksheet,
            search_value="合計",
            column_number=SUMMARY_PRODUCT_COLUMN,
            start_row=SUMMARY_START_ROW,
        )

        # 商品名列をクリア
        clear_cell_range(
            worksheet=worksheet,
            start_row=SUMMARY_START_ROW,
            end_row=total_row,
            start_column=SUMMARY_PRODUCT_COLUMN,
            end_column=SUMMARY_PRODUCT_COLUMN,
        )

        # 商品マスタ順の商品名をA列へ書き込む
        write_product_names(
            worksheet=worksheet,
            merged_df=merged_all_df,
            start_row=SUMMARY_START_ROW,
            product_column=SUMMARY_PRODUCT_COLUMN,
        )

        # 営業所ごとに数量を書き込む
        for branch_name, orders in branch_orders.items():
            branch_merged_df = merge_product_orders(
                base_df=merged_all_df,
                orders=orders,
            )

            write_orders(
                worksheet=worksheet,
                merged_df=branch_merged_df,
                start_row=SUMMARY_START_ROW,
                quantity_column=branch_columns[branch_name],
            )

        # 天板換算表へ商品マスタ情報を書き込む
        write_product_master_value(
            worksheet=tray_worksheet,
            product_master=product_master,
            product_names=product_names,
            search_value=TRAY_ORDER_MULTIPLIER_LABEL,
            header_row_num=TRAY_HEADER_ROW,
            data_row_num=TRAY_START_ROW,
            product_attribute="order_multiplier",
        )

        write_product_master_value(
            worksheet=tray_worksheet,
            product_master=product_master,
            product_names=product_names,
            search_value=TRAY_STORE_SALES_UNITS_LABEL,
            header_row_num=TRAY_HEADER_ROW,
            data_row_num=TRAY_START_ROW,
            product_attribute="store_sales_units",
        )

        # 計算シートへ商品マスタ情報を書き込む
        write_product_master_value(
            worksheet=profit_worksheet,
            product_master=product_master,
            product_names=product_names,
            search_value=PROFIT_UNIT_PRICE_LABEL,
            header_row_num=PROFIT_HEADER_ROW,
            data_row_num=PROFIT_START_ROW,
            product_attribute="unit_price",
        )

        write_product_master_value(
            worksheet=profit_worksheet,
            product_master=product_master,
            product_names=product_names,
            search_value=PROFIT_MARGIN_UNIT_PRICE_LABEL,
            header_row_num=PROFIT_HEADER_ROW,
            data_row_num=PROFIT_START_ROW,
            product_attribute="margin_unit_price",
        )

        write_product_master_value(
            worksheet=box_worksheet,
            product_master=product_master,
            product_names=product_names,
            search_value=BOX_UNITS_PER_BOX_LABEL,
            header_row_num=BOX_HEADER_ROW,
            data_row_num=BOX_START_ROW,
            product_attribute="units_per_box",
        )

        write_box_table_units(
            worksheet=box_table_worksheet,
            product_master=product_master,
            product_names=product_names[:27],
            value_column=2,
            start_row=3,
        )

        write_box_table_units(
            worksheet=box_table_worksheet,
            product_master=product_master,
            product_names=product_names[27:],
            value_column=4,
            start_row=3,
        )

        # 納品日をA1へ入力
        worksheet["A1"] = delivery_date

        # 元のExcelと同じフォルダへ別名保存
        date_suffix = delivery_date.strftime("%Y%m%d")
        output_path = (
                output_dir / f"{excel_path.stem}_{date_suffix}{excel_path.suffix}"
        )

        # 商品マスタシートは出力から除外
        if PRODUCT_MASTER_SHEET_NAME in workbook.sheetnames:
            del workbook[PRODUCT_MASTER_SHEET_NAME]

        save_workbook(
            workbook=workbook,
            output_path=output_path,
        )
        return output_path

    except Exception:
        workbook.close()
        raise
