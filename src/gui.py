import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from src.constants.paths import OUTPUT_DIR, CONFIG_DIR
from src.constants.app import APP_NAME, APP_VERSION
from src.models.label import LabelPrintState
from src.config.label_print_state import load_label_print_state, save_label_print_state
from src.config.app_config import load_config, save_config
from src.services.label import calculate_next_position, create_label_pdf_output
from src.services.label_pdf import cleanup_old_label_pdfs
from src.main import run_conversion as execute_conversion

# =========================================================
# 固定設定
# =========================================================

PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR.mkdir(exist_ok=True)
CONFIG_DIR.mkdir(exist_ok=True)


# =========================================================
# GUI初期設定
# =========================================================
root = tk.Tk()

root.title(f"{APP_NAME} v{APP_VERSION}")
root.geometry("650x350")
root.resizable(False, False)

selected_pdf_paths: list[Path] = []
config = load_config()
excel_path_text = tk.StringVar(value=config.get("excel_path", ""))
invoice_path_text = tk.StringVar(value=config.get("invoice_path", ""))
tray_path_text = tk.StringVar(value=config.get("tray_path", ""))
delivery_note_path_text = tk.StringVar(value=config.get("delivery_note_path", ""))
pdf_files_text = tk.StringVar(value="PDFファイルが選択されていません")
status_text = tk.StringVar(value="待機中")


# =========================================================
# ファイル選択
# =========================================================
def select_excel_file() -> None:
    """発注数量を書き込むExcelファイルを選択する。"""
    selected_path = filedialog.askopenfilename(
        title="発注書ファイルを選択",
        filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")]
    )
    if not selected_path:
        return
    excel_path_text.set(selected_path)
    status_text.set("Excel選択済み")

def select_transfer_file(target_var: tk.StringVar, title: str, filetypes: list[tuple[str, str]]) -> None:
    """転記先Excelファイルを選択する。"""
    selected_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if not selected_path:
        return
    target_var.set(selected_path)
    status_text.set("転記先ファイル選択済み")

def get_transfer_file_path(path_var: tk.StringVar, display_name: str, allowed_suffixes: set[str]) -> Path | None:
    """転記先Excelファイルを確認して返す。"""
    path_value = path_var.get().strip()
    if not path_value:
        messagebox.showwarning("入力不足",f"{display_name}を選択してください。")
        return None
    file_path = Path(path_value)
    if not file_path.exists():
        message = f"{display_name}が見つかりません。\nファイルを再選択してください。\n\n{file_path}"
        messagebox.showerror("ファイルエラー", message)
        return None

    if file_path.suffix.lower() not in allowed_suffixes:
        allowed_text = "、".join(sorted(allowed_suffixes))
        message = f"{display_name}のファイル形式が正しくありません。\n対応形式：{allowed_text}"
        messagebox.showerror("ファイル形式エラー", message)
        return None

    return file_path

def select_pdf_files() -> None:
    """処理対象のPDFファイルを複数選択する。"""
    file_types = [("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")]
    selected_paths = filedialog.askopenfilenames(title="発注PDFを選択", filetypes=file_types)
    if not selected_paths:
        return
    selected_pdf_paths.clear()
    selected_pdf_paths.extend(Path(path) for path in selected_paths)
    if len(selected_pdf_paths) == 1:
        display_text = selected_pdf_paths[0].name
    else:
        display_text = f"{len(selected_pdf_paths)}件選択：{selected_pdf_paths[0].name} ほか"
    pdf_files_text.set(display_text)
    status_text.set("PDF選択済み")

# =========================================================
# 入力チェック
# =========================================================
def get_selected_excel_path() -> Path | None:
    """選択されたExcelファイルを確認して返す。"""
    excel_path_value = excel_path_text.get()

    if not excel_path_value or excel_path_value == "入力するExcelファイルが選択されていません":
        messagebox.showwarning("入力不足","入力するExcelファイルを選択してください。")
        return None
    excel_path = Path(excel_path_value)
    message = f"選択されたExcelファイルが見つかりません。\n\n{excel_path}"
    if not excel_path.exists():
        messagebox.showerror("Excelエラー", message)
        return None

    if excel_path.suffix.lower() != ".xlsx":
        messagebox.showerror("Excelエラー","拡張子が.xlsxのExcelファイルを選択してください。")
        return None
    return excel_path

def validate_pdf_files() -> bool:
    """選択されたPDFファイルを確認する。"""
    if not selected_pdf_paths:
        messagebox.showwarning("入力不足","処理するPDFファイルを選択してください。")
        return False

    missing_pdf_paths = [path for path in selected_pdf_paths if not path.exists()]
    if missing_pdf_paths:
        message = f"選択されたPDFファイルが見つかりません。\n\n{missing_pdf_paths[0]}"
        messagebox.showerror("PDFエラー", message)
        return False
    return True

# =========================================================
# 変換処理
# =========================================================
def update_status(message: str) -> None:
    """画面下部の処理状況を更新する。"""
    status_text.set(message)
    root.update_idletasks()

def run_conversion() -> None:
    """選択されたファイルを使用して発注処理を実行する。"""
    excel_path = get_selected_excel_path()
    if excel_path is None:
        return

    if not validate_pdf_files():
        return

    invoice_path = get_transfer_file_path(
        path_var=invoice_path_text,
        display_name="請求書",
        allowed_suffixes={".xlsx"},
    )
    if invoice_path is None:
        return

    tray_path = get_transfer_file_path(
        path_var=tray_path_text,
        display_name="天板並べ早見表",
        allowed_suffixes={".xlsm"},
    )
    if tray_path is None:
        return

    delivery_note_path = get_transfer_file_path(
        path_var=delivery_note_path_text,
        display_name="納品書",
        allowed_suffixes={".xlsx"},
    )
    if delivery_note_path is None:
        return

    save_config(
        {
            "excel_path": str(excel_path),
            "invoice_path": str(invoice_path),
            "tray_path": str(tray_path),
            "delivery_note_path": str(delivery_note_path),
        }
    )

    try:
        update_status("発注データを処理しています...")

        execute_conversion(
            excel_path=excel_path,
            invoice_path=invoice_path,
            tray_path=tray_path,
            delivery_note_path=delivery_note_path,
            pdf_paths=selected_pdf_paths,
        )

        update_status("処理完了")
        show_label_print_window()

    except Exception as error:
        update_status("エラー")
        messagebox.showerror(
            "処理エラー",
            str(error),
        )

def run_label_print(floor: int, start_position: str, window: tk.Toplevel) -> None:
    """選択された発注PDFから指定階のラベルPDFを生成する。"""

    if not validate_pdf_files():
        return

    excel_path = get_selected_excel_path()
    if excel_path is None:
        return

    try:
        start_position_value = int(start_position)

    except ValueError:
        messagebox.showerror(title="入力エラー", message="開始位置は1～21の整数で入力してください。", parent=window)
        return

    if not 1 <= start_position_value <= 21:
        messagebox.showerror(title="入力エラー", message="開始位置は1～21で指定してください。", parent=window)
        return

    try:
        output_path, printed_count = create_label_pdf_output(
            excel_path=excel_path,
            pdf_paths=selected_pdf_paths,
            floor=floor,
            start_position=start_position_value,
            output_dir=OUTPUT_DIR,
        )

        next_position = calculate_next_position(
            start_position=start_position_value,
            printed_count=printed_count,
        )

        save_label_print_state(
            LabelPrintState(
                floor=floor,
                next_position=next_position,
            )
        )

        cleanup_old_label_pdfs(
            output_dir=OUTPUT_DIR,
            floor=floor,
            keep_count=3,
        )

        page_count = (printed_count + 20) // 21

        messagebox.showinfo(
            "ラベル発行完了",
            (
                f"{floor}Fラベルを作成しました。\n\n"
                f"PDF：{output_path.name}\n"
                f"ラベル枚数：{printed_count}枚\n"
                f"印刷ページ：約{page_count}枚"
            ),
            parent=window,
        )

    except Exception as error:
        messagebox.showerror(
            "ラベル発行エラー",
            str(error),
            parent=window,
        )

def show_label_print_window() -> None:
    """1F/2Fのラベル発行画面を表示する。"""

    window = tk.Toplevel(root)
    window.title("ラベル発行")
    window.grab_set()

    # =====================================================
    # 1F
    # =====================================================
    start_position_1f_var = tk.StringVar()
    state_1f = load_label_print_state(1)
    start_position_1f_var.set(str(state_1f.next_position))
    tk.Label(window, text="1F 開始位置").grid(row=0, column=0, padx=10, pady=10)
    tk.Entry(window, textvariable=start_position_1f_var, width=8).grid(row=0, column=1, padx=5)
    tk.Button(
        window,
        text="1Fラベル発行",
        width=15,
        command=lambda: run_label_print(
            floor=1,
            start_position=start_position_1f_var.get(),
            window=window,
        ),
    ).grid(
        row=0,
        column=2,
        padx=10,
        pady=10,
    )

    # =====================================================
    # 2F
    # =====================================================
    start_position_2f_var = tk.StringVar()
    state_2f = load_label_print_state(2)
    start_position_2f_var.set(str(state_2f.next_position))
    tk.Label(window, text="2F 開始位置").grid(row=1, column=0, padx=10, pady=10)
    tk.Entry(window, textvariable=start_position_2f_var, width=8).grid(row=1, column=1, padx=5)
    tk.Button(
        window,
        text="2Fラベル発行",
        width=15,
        command=lambda: run_label_print(
            floor=2,
            start_position=start_position_2f_var.get(),
            window=window,
        ),
    ).grid(
        row=1,
        column=2,
        padx=10,
        pady=10,
    )

# =========================================================
# GUI配置
# =========================================================
tk.Label(root, text="発注書").grid(row=0, column=0, padx=10, pady=15, sticky="e")
tk.Entry(root, textvariable=excel_path_text, width=65, state="readonly").grid(row=0, column=1, padx=5)
tk.Button(root, text="参照", width=10, command=select_excel_file).grid(row=0, column=2, padx=10)

invoice_title = "請求書ファイルを選択"
xlsx_types = [("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")]
tk.Label(root, text="請求書").grid(row=1, column=0, padx=10, pady=10, sticky="e")
tk.Entry(root, textvariable=invoice_path_text, width=65, state="readonly").grid(row=1, column=1, padx=5)
tk.Button(
    root, text="参照", width=10, command=lambda: select_transfer_file(
        target_var=invoice_path_text, title=invoice_title, filetypes=xlsx_types
    ),
).grid(row=1, column=2, padx=10)

tray_text = "天板並べ早見表を選択"
xlsm_types = [("マクロ有効Excel", "*.xlsm"), ("すべてのファイル", "*.*")]
tk.Label(root, text="天板並べ早見表").grid(row=2, column=0, padx=10, pady=10, sticky="e")
tk.Entry(root, textvariable=tray_path_text, width=65, state="readonly").grid(row=2, column=1, padx=5)
tk.Button(
    root, text="参照", width=10, command=lambda: select_transfer_file(
        target_var=tray_path_text, title=tray_text, filetypes=xlsm_types
    ),
).grid(row=2, column=2, padx=10)

delivery_note_text = "納品書ファイルを選択"
tk.Label(root, text="納品書").grid(row=3, column=0, padx=10, pady=10, sticky="e")
tk.Entry(root, textvariable=delivery_note_path_text, width=65, state="readonly").grid(row=3, column=1, padx=5)
tk.Button(root, text="参照", width=10, command=lambda: select_transfer_file(
    target_var=delivery_note_path_text, title=delivery_note_text, filetypes=xlsx_types
),
).grid(row=3, column=2, padx=10)

tk.Label(root, text="発注表PDF").grid(row=4, column=0, padx=10, pady=15, sticky="e")
tk.Entry(root, textvariable=pdf_files_text, width=65, state="readonly").grid(row=4, column=1, padx=5)
tk.Button(root, text="複数選択", width=10, command=select_pdf_files).grid(row=4, column=2, padx=10)

tk.Button(root, text="PDF読み込み", width=20, height=2, command=run_conversion).grid(row=5, column=1, pady=20)
tk.Label(root, textvariable=status_text).grid(row=9, column=1, pady=5)

root.mainloop()