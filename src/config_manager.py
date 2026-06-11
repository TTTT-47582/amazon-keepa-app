"""スクリーニング条件の設定ファイルを読み書きするモジュール"""

import yaml
from pathlib import Path

# 設定ファイルのパス（このファイルの2階層上の config/）
_CONFIG_PATH = Path(__file__).parent.parent / "config" / "screening_filters.yaml"


def load_filters() -> dict:
    """YAML ファイルからスクリーニング条件を読み込む"""
    with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def save_filters(config: dict) -> None:
    """スクリーニング条件を YAML ファイルに書き込む"""
    with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
