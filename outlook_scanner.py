import sys
import traceback
from pathlib import Path
from datetime import datetime

import pythoncom
import win32com.client
import win32timezone  # noqa: F401


DELETED_NAMES = {
    "deleted items",
    "deleted item",
    "削除済みアイテム",
    "削除済みアイテム フォルダー",
}

DRAFT_NAMES = {
    "drafts",
    "下書き",
}

OUTBOX_NAMES = {
    "outbox",
    "送信トレイ",
}

JUNK_NAMES = {
    "junk email",
    "junk e-mail",
    "迷惑メール",
}

RSS_NAMES = {
    "rss feeds",
    "rss フィード",
}

SYNC_NAMES = {
    "sync issues",
    "同期の問題",
}

SEARCH_FOLDER_NAMES = {
    "search folders",
    "検索フォルダー",
    "検索フォルダ",
}

NO_RECURSE_NAMES = (
    SYNC_NAMES
    | RSS_NAMES
    | SEARCH_FOLDER_NAMES
)


class ScanCancelled(Exception):
    pass


class OutlookScanner:
    def __init__(
        self,
        cancel_callback=None,
        progress_callback=None,
    ):
        self.cancel_callback = cancel_callback
        self.progress_callback = progress_callback

        self.application = None
        self.namespace = None

        self.log_path = self._get_base_dir() / "scanner_debug.log"

    def _get_base_dir(self):
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent

    def _log(self, text):
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(f"{timestamp}  {text}\n")
        except Exception:
            pass

    def _check_cancel(self):
        if self.cancel_callback and self.cancel_callback():
            raise ScanCancelled()

    def _progress(self, text):
        if self.progress_callback:
            try:
                self.progress_callback(str(text))
            except Exception:
                pass

    def _safe_str(self, getter, default=""):
        try:
            value = getter()
            if value is None:
                return default
            return str(value)
        except Exception:
            return default

    def _normalize_name(self, name):
        return str(name).strip().casefold()

    def _classify_folder(self, display_name):
        name = self._normalize_name(display_name)

        if name in DELETED_NAMES:
            return "deleted_items"
        if name in DRAFT_NAMES:
            return "drafts"
        if name in OUTBOX_NAMES:
            return "outbox"
        if name in JUNK_NAMES:
            return "junk"
        if name in RSS_NAMES:
            return "rss"
        if name in SYNC_NAMES:
            return "sync_issues"
        if name in SEARCH_FOLDER_NAMES:
            return "search_folders"

        return ""

    def _is_hard_disabled(self, special_type):
        return special_type == "deleted_items"

    def _default_enabled(self, special_type):
        if special_type in {
            "deleted_items",
            "drafts",
            "outbox",
            "junk",
            "rss",
            "sync_issues",
            "search_folders",
        }:
            return False

        return True

    def _should_recurse(self, display_name, special_type):
        name = self._normalize_name(display_name)

        if name in NO_RECURSE_NAMES:
            return False

        if special_type in {
            "sync_issues",
            "rss",
            "search_folders",
        }:
            return False

        return True

    def scan(self):
        try:
            if self.log_path.exists():
                self.log_path.unlink()
        except Exception:
            pass

        self._log("===== Outlook scan start =====")

        pythoncom.CoInitialize()

        try:
            self._check_cancel()

            self._progress("Outlookへ接続しています...")

            self._log("BEFORE Dispatch Outlook.Application")
            self.application = win32com.client.Dispatch("Outlook.Application")
            self._log("AFTER Dispatch Outlook.Application")

            self._log("BEFORE GetNamespace(MAPI)")
            self.namespace = self.application.GetNamespace("MAPI")
            self._log("AFTER GetNamespace(MAPI)")

            self._check_cancel()

            self._log("BEFORE namespace.Stores")
            stores_collection = self.namespace.Stores
            self._log("AFTER namespace.Stores")

            try:
                store_count = int(stores_collection.Count)
            except Exception:
                store_count = 0

            self._log(f"Store count = {store_count}")

            result = []

            for store_index in range(1, store_count + 1):
                self._check_cancel()

                self._log("----------------------------------------")
                self._log(f"STORE {store_index}/{store_count}")

                self._log(f"BEFORE Stores.Item({store_index})")

                try:
                    store = stores_collection.Item(store_index)
                except Exception:
                    self._log(traceback.format_exc())
                    continue

                self._log(f"AFTER Stores.Item({store_index})")

                store_name = self._safe_str(
                    lambda: store.DisplayName,
                    f"Store {store_index}",
                )

                self._log(f"Store name = {store_name}")

                self._progress(
                    f"Outlook解析中 {store_index} / {store_count}\n"
                    f"{store_name}"
                )

                self._log("BEFORE store.StoreID")
                store_id = self._safe_str(lambda: store.StoreID)
                self._log("AFTER store.StoreID")

                if not store_id:
                    self._log("StoreID empty -> skip")
                    continue

                self._log(f"BEFORE GetRootFolder [{store_name}]")

                try:
                    root = store.GetRootFolder()
                except Exception:
                    error_text = traceback.format_exc()
                    self._log(error_text)

                    result.append(
                        {
                            "node_type": "store",
                            "display_name": store_name,
                            "store_id": store_id,
                            "folder_entry_id": "",
                            "folder_path": store_name,
                            "parent_folder_entry_id": "",
                            "enabled": True,
                            "hard_disabled": False,
                            "special_type": "",
                            "available": True,
                            "is_new": False,
                            "effective_enabled": True,
                            "scan_error": error_text,
                            "children": [],
                        }
                    )
                    continue

                self._log(f"AFTER GetRootFolder [{store_name}]")

                store_node = {
                    "node_type": "store",
                    "display_name": store_name,
                    "store_id": store_id,
                    "folder_entry_id": "",
                    "folder_path": store_name,
                    "parent_folder_entry_id": "",
                    "enabled": True,
                    "hard_disabled": False,
                    "special_type": "",
                    "available": True,
                    "is_new": False,
                    "effective_enabled": True,
                    "scan_error": "",
                    "children": [],
                }

                try:
                    store_node["children"] = self._scan_children(
                        parent_folder=root,
                        store_id=store_id,
                        parent_path=store_name,
                        depth=0,
                    )
                except ScanCancelled:
                    raise
                except Exception:
                    store_node["scan_error"] = traceback.format_exc()
                    self._log(store_node["scan_error"])

                result.append(store_node)

            self._log("===== Outlook scan completed =====")
            return result

        finally:
            self.application = None
            self.namespace = None
            pythoncom.CoUninitialize()

    def _scan_children(
        self,
        parent_folder,
        store_id,
        parent_path,
        depth,
    ):
        result = []

        self._check_cancel()

        if depth > 100:
            self._log(f"MAX DEPTH reached: {parent_path}")
            return result

        parent_entry_id = self._safe_str(
            lambda: parent_folder.EntryID
        )

        self._log(f"BEFORE .Folders : {parent_path}")

        try:
            folders = parent_folder.Folders
        except Exception:
            self._log(
                f"ERROR .Folders : {parent_path}\n"
                f"{traceback.format_exc()}"
            )
            return result

        self._log(f"AFTER .Folders : {parent_path}")

        self._log(f"BEFORE .Folders.Count : {parent_path}")

        try:
            count = int(folders.Count)
        except Exception:
            self._log(
                f"ERROR .Folders.Count : {parent_path}\n"
                f"{traceback.format_exc()}"
            )
            return result

        self._log(
            f"AFTER .Folders.Count={count} : {parent_path}"
        )

        for i in range(1, count + 1):
            self._check_cancel()

            self._log(
                f"BEFORE Folders.Item({i}) : {parent_path}"
            )

            try:
                folder = folders.Item(i)
            except Exception:
                self._log(
                    f"ERROR Folders.Item({i}) : {parent_path}\n"
                    f"{traceback.format_exc()}"
                )
                continue

            self._log(
                f"AFTER Folders.Item({i}) : {parent_path}"
            )

            display_name = self._safe_str(
                lambda: folder.Name,
                "(名称取得失敗)",
            )

            folder_path = f"{parent_path}\\{display_name}"

            self._progress(folder_path)
            self._log(f"FOLDER START : {folder_path}")

            self._log(f"BEFORE EntryID : {folder_path}")
            entry_id = self._safe_str(lambda: folder.EntryID)
            self._log(f"AFTER EntryID : {folder_path}")

            if not entry_id:
                self._log(
                    f"EntryID empty -> skip : {folder_path}"
                )
                continue

            special_type = self._classify_folder(display_name)
            hard_disabled = self._is_hard_disabled(special_type)
            enabled = self._default_enabled(special_type)

            if hard_disabled:
                enabled = False

            node = {
                "node_type": "folder",
                "display_name": display_name,
                "store_id": store_id,
                "folder_entry_id": entry_id,
                "folder_path": folder_path,
                "parent_folder_entry_id": parent_entry_id,
                "enabled": enabled,
                "hard_disabled": hard_disabled,
                "special_type": special_type,
                "available": True,
                "is_new": False,
                "effective_enabled": enabled,
                "scan_error": "",
                "children": [],
            }

            if self._should_recurse(
                display_name,
                special_type,
            ):
                self._log(f"RECURSE START : {folder_path}")

                try:
                    node["children"] = self._scan_children(
                        parent_folder=folder,
                        store_id=store_id,
                        parent_path=folder_path,
                        depth=depth + 1,
                    )
                except ScanCancelled:
                    raise
                except Exception:
                    node["scan_error"] = traceback.format_exc()
                    self._log(
                        f"RECURSE ERROR : {folder_path}\n"
                        f"{node['scan_error']}"
                    )

                self._log(f"RECURSE END : {folder_path}")

            else:
                self._log(f"NO RECURSE : {folder_path}")

            result.append(node)
            self._log(f"FOLDER END : {folder_path}")

        return result
