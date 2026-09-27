from pathlib import Path

import win32com.client


def recalculate_workbook(excel_path: Path) -> None:
    """
    Excel本体でブックを開き、 数式を再計算して上書き保存する。
    """
    excel = None
    workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(excel_path.resolve()))
        excel.CalculateFullRebuild()
        workbook.Save()

    except Exception as error:
        raise RuntimeError(
            "集計表の数式再計算に失敗しました。\nExcelがインストールされているか、対象ファイルが開かれていないか確認してください。\n\n"
            f"{excel_path}"
        ) from error

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)

        if excel is not None:
            excel.Quit()