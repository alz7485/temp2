import json, os
from pathlib import Path
SETTINGS_VERSION = 1
class SettingsManager:
    def __init__(self, base_dir: Path):
        self.base_dir = Path(base_dir); self.settings_path = self.base_dir/"builder_settings.json"
    def exists(self): return self.settings_path.exists()
    def load(self):
        if not self.exists(): return {"version": SETTINGS_VERSION, "stores": []}
        with self.settings_path.open("r", encoding="utf-8") as f: data=json.load(f)
        if not isinstance(data, dict) or not isinstance(data.get("stores", []), list):
            raise ValueError("builder_settings.json の形式が不正です。")
        return data
    def save(self, stores):
        tmp=self.settings_path.with_suffix(".json.tmp")
        with tmp.open("w", encoding="utf-8") as f:
            json.dump({"version":SETTINGS_VERSION,"stores":stores}, f, ensure_ascii=False, indent=2); f.flush()
            try: os.fsync(f.fileno())
            except Exception: pass
        os.replace(tmp, self.settings_path)
