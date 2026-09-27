import json

from src.constants.paths import CONFIG_PATH

DEFAULT_CONFIG = {
    "excel_path": "",
    "invoice_path": "",
    "tray_path": "",
    "label_path": "",
    "delivery_note_path": "",
}


def load_config() -> dict:
    """
    config.jsonを読み込む。
    ファイルがない場合や壊れている場合は初期値を返す。
    """
    if not CONFIG_PATH.exists():
        return DEFAULT_CONFIG.copy()

    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as file:
            loaded_config = json.load(file)

        # 設定項目が不足していてもエラーにならないようにする
        config = DEFAULT_CONFIG.copy()
        config.update(loaded_config)

        return config

    except (OSError, json.JSONDecodeError, TypeError):
        return DEFAULT_CONFIG.copy()


def save_config(config: dict) -> None:
    """
    辞書をconfig.jsonへ保存する。
    """
    try:
        with CONFIG_PATH.open("w", encoding="utf-8") as file:
            json.dump(
                config,
                file,
                ensure_ascii=False,
                indent=4,
            )

    except OSError as error:
        raise RuntimeError(
            f"設定ファイルの保存に失敗しました。\n{CONFIG_PATH}"
        ) from error
    