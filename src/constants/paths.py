from src.utils.path import get_app_dir


CONFIG_DIR = get_app_dir() / "config"
CONFIG_PATH = CONFIG_DIR / "config.json"
STATE_PATH = CONFIG_DIR / "label_print_state.json"
OUTPUT_DIR = get_app_dir() / "output"