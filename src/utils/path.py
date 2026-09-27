import sys
from pathlib import Path


def get_app_dir() -> Path:
    """
    開発中は実行ファイルのあるプロジェクトフォルダ、exe化後はexeが置かれているフォルダを取得する。
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parents[2]