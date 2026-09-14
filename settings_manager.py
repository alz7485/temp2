import json
import os
from pathlib import Path


SETTINGS_VERSION = 1


class SettingsManager:
    def __init__(self, base_dir: Path):
        self.base_dir = Path(base_dir)
        self.settings_path = self.base_dir / "builder_settings.json"

    def exists(self) -> bool:
        return self.settings_path.exists()

    def load(self) -> dict:
        if not self.settings_path.exists():
            return {
                "version": SETTINGS_VERSION,
                "stores": [],
            }

        with open(self.settings_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        if not isinstance(data, dict):
            raise ValueError("builder_settings.json の形式が不正です。")

        stores = data.get("stores", [])

        if not isinstance(stores, list):
            raise ValueError("builder_settings.json の stores が不正です。")

        return data

    def save(self, stores):
        data = {
            "version": SETTINGS_VERSION,
            "stores": stores,
        }

        tmp_path = self.settings_path.with_suffix(".json.tmp")

        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2,
            )
            f.flush()

            try:
                os.fsync(f.fileno())
            except Exception:
                pass

        os.replace(tmp_path, self.settings_path)
