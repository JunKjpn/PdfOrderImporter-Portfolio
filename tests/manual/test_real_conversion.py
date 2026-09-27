import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

from src.main import run_conversion


EXCEL_PATH = BASE_DIR / "templates" / "集計表.xlsx"
INVOICE_PATH = BASE_DIR / "templates" / "請求書.xlsx"
TRAY_PATH = BASE_DIR / "templates" / "天板並べ早見表.xlsm"
DELIVERY_NOTE_PATH = BASE_DIR / "templates" / "納品書.xlsx"

PDF_DIR = BASE_DIR / "sample"


def main() -> None:
    """実際のPDF・Excelファイルを使用して発注処理を実行する。"""

    pdf_paths = sorted(PDF_DIR.glob("*.pdf"))

    if not pdf_paths:
        raise FileNotFoundError(
            f"PDFファイルが見つかりません: {PDF_DIR}"
        )

    print("=== 実機テスト開始 ===")
    print(f"PDFファイル: {len(pdf_paths)}件")
    print(f"集計表: {EXCEL_PATH}")
    print(f"請求書: {INVOICE_PATH}")
    print(f"天板並べ早見表: {TRAY_PATH}")
    print(f"納品書: {DELIVERY_NOTE_PATH}")
    print()

    result_paths = run_conversion(
        excel_path=EXCEL_PATH,
        invoice_path=INVOICE_PATH,
        tray_path=TRAY_PATH,
        delivery_note_path=DELIVERY_NOTE_PATH,
        pdf_paths=pdf_paths,
        label_floor=1,
        label_start_position=1,
    )

    print()
    print("=== 実機テスト完了 ===")
    print()
    print("生成されたファイル:")

    for path in result_paths:
        print(f"  {path}")


if __name__ == "__main__":
    main()