from datetime import datetime
from pathlib import Path
import re

from src.constants.branches import BRANCHES


def parse_filename(pdf_path: Path) -> tuple[datetime, str]:
    """
    ファイル名から日付と営業所を取得する。
    Returns:
        ("2026/07/10", "営業所C")
    """
    filename = pdf_path.stem
    pattern = rf"(\d{{6}})({'|'.join(BRANCHES)})"
    match = re.match(pattern, filename)

    if match is None:
        raise ValueError(
            f"ファイル名の形式が正しくありません: {filename}"
        )

    delivery_date = datetime.strptime(match.group(1), "%y%m%d").date()
    branch_name = match.group(2)

    return delivery_date, branch_name