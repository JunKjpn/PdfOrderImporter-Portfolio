import json

from src.constants.paths import CONFIG_DIR, STATE_PATH
from src.models.label import LabelPrintState


def load_label_print_state(floor: int) -> LabelPrintState:
    """
    指定階のラベル印刷状態を読み込む。
    """

    if floor not in (1, 2):
        raise ValueError(f"包装階は1または2で指定してください：{floor}")

    if not STATE_PATH.exists():
        return LabelPrintState(
            floor=floor,
            next_position=1,
        )

    with STATE_PATH.open("r", encoding="utf-8") as file:
        data = json.load(file)

    floor_data = data.get(f"{floor}F")

    if floor_data is None:
        return LabelPrintState(
            floor=floor,
            next_position=1,
        )

    return LabelPrintState(
        floor=floor,
        next_position=int(floor_data["next_position"]),
    )


def save_label_print_state(state: LabelPrintState) -> None:
    """
    指定階のラベル印刷状態を保存する。
    """

    if state.floor not in (1, 2):
        raise ValueError(
            f"包装階は1または2で指定してください：{state.floor}"
        )

    if not 1 <= state.next_position <= 21:
        raise ValueError(
            f"ラベル位置は1～21で指定してください：{state.next_position}"
        )

    CONFIG_DIR.mkdir(parents=True, exist_ok=True)

    data = {}

    if STATE_PATH.exists():
        with STATE_PATH.open("r", encoding="utf-8") as file:
            data = json.load(file)

    data[f"{state.floor}F"] = {
        "floor": state.floor,
        "next_position": state.next_position,
    }

    with STATE_PATH.open("w", encoding="utf-8") as file:
        json.dump(
            data,
            file,
            ensure_ascii=False,
            indent=4,
        )