from pathlib import Path

import win32api


def print_pdf(pdf_path: Path) -> None:
    """
    PDFをWindowsの既定プリンターへ送信する。
    """
    if not pdf_path.exists():
        raise FileNotFoundError(f"印刷対象のPDFが見つかりません。\n\n{pdf_path}")

    try:
        win32api.ShellExecute(0,"print", str(pdf_path),None,".",0)


    except Exception as error:
        raise RuntimeError(
            f"PDFの印刷処理を開始できませんでした。\n\n"
            f"PDF：{pdf_path}\n"
            f"エラー：{error}"
        ) from error