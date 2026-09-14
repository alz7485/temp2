from copy import deepcopy

from PySide6.QtCore import (
    QObject,
    QThread,
    Qt,
    Signal,
)
from PySide6.QtGui import (
    QBrush,
    QColor,
    QCloseEvent,
)
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from outlook_scanner import (
    OutlookScanner,
    ScanCancelled,
)


ROLE_NODE = Qt.ItemDataRole.UserRole


class ScanWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)
    cancelled = Signal()
    progress = Signal(str)

    def __init__(self):
        super().__init__()
        self._cancel_requested = False

    def request_cancel(self):
        self._cancel_requested = True

    def is_cancelled(self):
        return self._cancel_requested

    def emit_progress(self, text):
        self.progress.emit(str(text))

    def run(self):
        try:
            scanner = OutlookScanner(
                cancel_callback=self.is_cancelled,
                progress_callback=self.emit_progress,
            )

            stores = scanner.scan()

            if self._cancel_requested:
                self.cancelled.emit()
                return

            self.finished.emit(stores)

        except ScanCancelled:
            self.cancelled.emit()

        except Exception:
            import traceback
            self.failed.emit(traceback.format_exc())


class MainWindow(QMainWindow):
    def __init__(
        self,
        settings_manager,
    ):
        super().__init__()

        self.settings_manager = settings_manager
        self.stores = []

        self.scan_thread = None
        self.scan_worker = None
        self.scan_running = False

        self._scan_initial = False
        self._scan_result = None
        self._scan_error = None
        self._scan_was_cancelled = False

        self.dirty = False
        self._updating_tree = False
        self._close_after_scan = False

        self.setWindowTitle("Outlook DB Builder - 設定")
        self.resize(900, 760)

        self._build_ui()
        self._load_initial_data()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)

        header_layout = QHBoxLayout()

        title_label = QLabel("Outlookフォルダ")
        title_label.setStyleSheet(
            "font-size: 16px;"
            "font-weight: bold;"
        )

        header_layout.addWidget(title_label)
        header_layout.addStretch()

        self.rescan_button = QPushButton("Outlookを再解析")
        self.rescan_button.clicked.connect(
            self._on_rescan_clicked
        )

        header_layout.addWidget(self.rescan_button)
        main_layout.addLayout(header_layout)

        explanation = QLabel(
            "チェックを付けたフォルダを検索DBの対象にします。\n"
            "親フォルダをOFFにすると、子フォルダの設定は保持したまま"
            "一時的に無効になります。"
        )
        explanation.setWordWrap(True)
        main_layout.addWidget(explanation)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Outlookフォルダ"])
        self.tree.setAlternatingRowColors(True)
        self.tree.itemChanged.connect(
            self._on_item_changed
        )

        main_layout.addWidget(self.tree, 1)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        main_layout.addWidget(self.status_label)

        bottom_layout = QHBoxLayout()

        self.cancel_button = QPushButton("解析キャンセル")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(
            self._cancel_scan
        )
        bottom_layout.addWidget(self.cancel_button)

        bottom_layout.addStretch()

        self.save_button = QPushButton("設定保存")
        self.save_button.clicked.connect(
            self._save_settings
        )
        bottom_layout.addWidget(self.save_button)

        self.close_button = QPushButton("閉じる")
        self.close_button.clicked.connect(self.close)
        bottom_layout.addWidget(self.close_button)

        main_layout.addLayout(bottom_layout)

    def _load_initial_data(self):
        if self.settings_manager.exists():
            try:
                data = self.settings_manager.load()
                self.stores = data.get("stores", [])

                self._normalize_loaded_nodes(
                    self.stores
                )
                self._populate_tree()

                self.status_label.setText(
                    "保存済み設定を読み込みました。\n"
                    "Outlookへの接続は行っていません。"
                )

                self._set_dirty(False)

            except Exception as e:
                QMessageBox.warning(
                    self,
                    "設定読込エラー",
                    "builder_settings.json を読み込めませんでした。\n\n"
                    f"{e}\n\n"
                    "Outlookを再解析します。",
                )

                self._start_scan(
                    initial_scan=True
                )

        else:
            self._start_scan(
                initial_scan=True
            )

    def _normalize_loaded_nodes(
        self,
        nodes,
    ):
        for node in nodes:
            node.setdefault("available", True)
            node.setdefault("is_new", False)
            node.setdefault("hard_disabled", False)
            node.setdefault("special_type", "")
            node.setdefault("scan_error", "")
            node.setdefault("effective_enabled", False)

            children = node.setdefault(
                "children",
                [],
            )

            self._normalize_loaded_nodes(
                children
            )

    def _on_rescan_clicked(self):
        if self.scan_running:
            return

        self._start_scan(
            initial_scan=False
        )

    def _start_scan(
        self,
        initial_scan=False,
    ):
        if self.scan_running:
            return

        self.scan_running = True

        self._scan_initial = bool(
            initial_scan
        )
        self._scan_result = None
        self._scan_error = None
        self._scan_was_cancelled = False

        self.rescan_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.cancel_button.setEnabled(True)

        self.status_label.setText(
            "Outlookを解析しています..."
        )

        self.scan_thread = QThread(self)
        self.scan_worker = ScanWorker()

        self.scan_worker.moveToThread(
            self.scan_thread
        )

        self.scan_thread.started.connect(
            self.scan_worker.run
        )

        self.scan_worker.progress.connect(
            self._on_scan_progress
        )

        self.scan_worker.finished.connect(
            self._store_scan_result
        )
        self.scan_worker.failed.connect(
            self._store_scan_error
        )
        self.scan_worker.cancelled.connect(
            self._store_scan_cancelled
        )

        self.scan_worker.finished.connect(
            self.scan_thread.quit
        )
        self.scan_worker.failed.connect(
            self.scan_thread.quit
        )
        self.scan_worker.cancelled.connect(
            self.scan_thread.quit
        )

        self.scan_thread.finished.connect(
            self._on_scan_thread_finished
        )

        self.scan_thread.finished.connect(
            self.scan_worker.deleteLater
        )
        self.scan_thread.finished.connect(
            self.scan_thread.deleteLater
        )

        self.scan_thread.start()

    def _store_scan_result(
        self,
        stores,
    ):
        self._scan_result = stores

    def _store_scan_error(
        self,
        error_text,
    ):
        self._scan_error = error_text

    def _store_scan_cancelled(self):
        self._scan_was_cancelled = True

    def _on_scan_progress(
        self,
        text,
    ):
        self.status_label.setText(text)

    def _cancel_scan(self):
        if self.scan_worker is not None:
            self.scan_worker.request_cancel()

            self.status_label.setText(
                "解析のキャンセルを要求しました..."
            )

            self.cancel_button.setEnabled(False)

    def _on_scan_thread_finished(self):
        self.scan_running = False

        self.rescan_button.setEnabled(True)
        self.save_button.setEnabled(True)
        self.cancel_button.setEnabled(False)

        scan_result = self._scan_result
        scan_error = self._scan_error
        was_cancelled = self._scan_was_cancelled
        was_initial = self._scan_initial

        self.scan_worker = None
        self.scan_thread = None

        self._scan_result = None
        self._scan_error = None
        self._scan_was_cancelled = False

        if was_cancelled:
            self.status_label.setText(
                "Outlook解析をキャンセルしました。"
            )
            self._finish_close_if_requested()
            return

        if scan_error:
            self.status_label.setText(
                "Outlook解析に失敗しました。"
            )

            QMessageBox.critical(
                self,
                "Outlook解析エラー",
                scan_error,
            )

            self._finish_close_if_requested()
            return

        if scan_result is None:
            self.status_label.setText(
                "Outlook解析結果を取得できませんでした。"
            )
            self._finish_close_if_requested()
            return

        self.status_label.setText(
            "解析結果を画面に反映しています..."
        )

        try:
            if was_initial:
                self.stores = scan_result
                self._clear_new_flags(
                    self.stores
                )
            else:
                self.stores = self._merge_rescan(
                    self.stores,
                    scan_result,
                )

            self._populate_tree()

            if was_initial:
                self.status_label.setText(
                    "Outlook解析が完了しました。\n"
                    "対象フォルダを確認して"
                    "［設定保存］を押してください。"
                )
            else:
                self.status_label.setText(
                    "Outlookの再解析が完了しました。\n"
                    "NEW は新しく見つかったフォルダです。"
                )

            self._set_dirty(True)

        except Exception:
            import traceback

            error_text = traceback.format_exc()

            self.status_label.setText(
                "解析結果の表示中にエラーが発生しました。"
            )

            QMessageBox.critical(
                self,
                "ツリー表示エラー",
                error_text,
            )

        self._finish_close_if_requested()

    def _finish_close_if_requested(self):
        if self._close_after_scan:
            self._close_after_scan = False
            self.close()

    def _identity(
        self,
        node,
    ):
        if node.get("node_type") == "store":
            return (
                "store",
                node.get("store_id", ""),
            )

        return (
            "folder",
            node.get("store_id", ""),
            node.get("folder_entry_id", ""),
        )

    def _build_index(
        self,
        nodes,
        output=None,
    ):
        if output is None:
            output = {}

        for node in nodes:
            output[
                self._identity(node)
            ] = node

            self._build_index(
                node.get("children", []),
                output,
            )

        return output

    def _merge_rescan(
        self,
        old_stores,
        new_stores,
    ):
        old_index = self._build_index(
            old_stores
        )

        for store in new_stores:
            store_identity = self._identity(
                store
            )

            old_store = old_index.get(
                store_identity
            )

            if old_store is None:
                self._mark_new_subtree(
                    store
                )
            else:
                self._merge_existing_node(
                    store,
                    old_index,
                )

        new_index = self._build_index(
            new_stores
        )
        old_index = self._build_index(
            old_stores
        )

        for identity, old_node in list(
            old_index.items()
        ):
            if identity in new_index:
                continue

            parent_identity = self._parent_identity(
                old_node
            )

            if (
                parent_identity
                and parent_identity not in new_index
                and parent_identity in old_index
            ):
                continue

            missing_copy = self._clone_unavailable(
                old_node,
                new_index,
            )

            if missing_copy is None:
                continue

            if old_node.get("node_type") == "store":
                new_stores.append(
                    missing_copy
                )
            else:
                parent_node = new_index.get(
                    parent_identity
                )

                if parent_node is not None:
                    parent_node.setdefault(
                        "children",
                        [],
                    ).append(
                        missing_copy
                    )

            new_index = self._build_index(
                new_stores
            )

        return new_stores

    def _merge_existing_node(
        self,
        node,
        old_index,
    ):
        identity = self._identity(node)
        old_node = old_index.get(identity)

        if old_node is None:
            self._mark_new_subtree(node)
            return

        node["enabled"] = bool(
            old_node.get("enabled", False)
        )
        node["available"] = True
        node["is_new"] = False

        if node.get("hard_disabled", False):
            node["enabled"] = False

        for child in node.get("children", []):
            child_identity = self._identity(
                child
            )

            if child_identity in old_index:
                self._merge_existing_node(
                    child,
                    old_index,
                )
            else:
                self._mark_new_subtree(
                    child
                )

    def _mark_new_subtree(
        self,
        node,
    ):
        node["enabled"] = False
        node["available"] = True
        node["is_new"] = True
        node["effective_enabled"] = False

        for child in node.get(
            "children",
            [],
        ):
            self._mark_new_subtree(
                child
            )

    def _clone_unavailable(
        self,
        old_node,
        current_new_index,
    ):
        identity = self._identity(
            old_node
        )

        if identity in current_new_index:
            return None

        clone = deepcopy(
            old_node
        )

        clone["available"] = False
        clone["is_new"] = False
        clone["effective_enabled"] = False

        clone_children = []

        for old_child in old_node.get(
            "children",
            [],
        ):
            child_clone = self._clone_unavailable(
                old_child,
                current_new_index,
            )

            if child_clone is not None:
                clone_children.append(
                    child_clone
                )

        clone["children"] = clone_children

        return clone

    def _parent_identity(
        self,
        node,
    ):
        if node.get("node_type") == "store":
            return None

        store_id = node.get(
            "store_id",
            "",
        )
        parent_entry_id = node.get(
            "parent_folder_entry_id",
            "",
        )

        if parent_entry_id:
            return (
                "folder",
                store_id,
                parent_entry_id,
            )

        return (
            "store",
            store_id,
        )

    def _clear_new_flags(
        self,
        nodes,
    ):
        for node in nodes:
            node["is_new"] = False

            self._clear_new_flags(
                node.get("children", [])
            )

    def _populate_tree(self):
        self._updating_tree = True

        try:
            self.tree.clear()

            for store in self.stores:
                item = self._create_tree_item(
                    store
                )

                self.tree.addTopLevelItem(
                    item
                )

                self._add_children(
                    item,
                    store,
                )

                item.setExpanded(True)

            self._refresh_effective_states()

        finally:
            self._updating_tree = False

    def _create_tree_item(
        self,
        node,
    ):
        text = node.get(
            "display_name",
            "(名称なし)",
        )

        if node.get("is_new", False):
            text += "    NEW"

        if not node.get("available", True):
            text = "⚠ " + text + "    利用不可"

        item = QTreeWidgetItem([text])

        item.setData(
            0,
            ROLE_NODE,
            node,
        )

        enabled = bool(
            node.get("enabled", False)
        )

        if node.get("hard_disabled", False):
            enabled = False
            node["enabled"] = False

        item.setCheckState(
            0,
            (
                Qt.CheckState.Checked
                if enabled
                else Qt.CheckState.Unchecked
            ),
        )

        flags = item.flags()
        flags |= Qt.ItemFlag.ItemIsUserCheckable

        if node.get("hard_disabled", False):
            flags &= ~Qt.ItemFlag.ItemIsUserCheckable

            item.setToolTip(
                0,
                "削除済みアイテムは"
                "検索DBの対象にできません。",
            )

        if not node.get("available", True):
            flags &= ~Qt.ItemFlag.ItemIsUserCheckable

            item.setToolTip(
                0,
                "前回の設定には存在しますが、"
                "今回のOutlook解析では"
                "確認できませんでした。",
            )

        item.setFlags(flags)

        return item

    def _add_children(
        self,
        parent_item,
        parent_node,
    ):
        for child_node in parent_node.get(
            "children",
            [],
        ):
            child_item = self._create_tree_item(
                child_node
            )

            parent_item.addChild(
                child_item
            )

            self._add_children(
                child_item,
                child_node,
            )

    def _on_item_changed(
        self,
        item,
        column,
    ):
        if self._updating_tree:
            return

        if column != 0:
            return

        node = item.data(
            0,
            ROLE_NODE,
        )

        if not node:
            return

        if node.get("hard_disabled", False):
            return

        if not node.get("available", True):
            return

        checked = (
            item.checkState(0)
            == Qt.CheckState.Checked
        )

        node["enabled"] = bool(checked)

        self._refresh_effective_states()
        self._set_dirty(True)

    def _refresh_effective_states(self):
        previous_state = self._updating_tree
        self._updating_tree = True

        try:
            for i in range(
                self.tree.topLevelItemCount()
            ):
                item = self.tree.topLevelItem(i)

                self._update_effective_state(
                    item=item,
                    parent_effective=True,
                )

        finally:
            self._updating_tree = previous_state

    def _update_effective_state(
        self,
        item,
        parent_effective,
    ):
        node = item.data(
            0,
            ROLE_NODE,
        )

        if not node:
            return

        available = bool(
            node.get("available", True)
        )
        hard_disabled = bool(
            node.get("hard_disabled", False)
        )
        own_enabled = bool(
            node.get("enabled", False)
        )

        effective = (
            parent_effective
            and own_enabled
            and available
            and not hard_disabled
        )

        node["effective_enabled"] = effective

        if (
            not available
            or hard_disabled
            or not parent_effective
        ):
            item.setForeground(
                0,
                QBrush(
                    QColor("#888888")
                ),
            )
        else:
            item.setForeground(
                0,
                QBrush(),
            )

        for i in range(
            item.childCount()
        ):
            child_item = item.child(i)

            self._update_effective_state(
                item=child_item,
                parent_effective=effective,
            )

    def _save_settings(self):
        if self.scan_running:
            return

        try:
            self._clear_new_flags(
                self.stores
            )

            self.settings_manager.save(
                self.stores
            )

            self._populate_tree()
            self._set_dirty(False)

            self.status_label.setText(
                "設定を保存しました。\n"
                "builder_settings.json"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "保存エラー",
                "設定を保存できませんでした。\n\n"
                f"{e}",
            )

    def _set_dirty(
        self,
        dirty,
    ):
        self.dirty = bool(dirty)

        title = "Outlook DB Builder - 設定"

        if self.dirty:
            title += " *"

        self.setWindowTitle(title)

    def closeEvent(
        self,
        event: QCloseEvent,
    ):
        if self.scan_running:
            answer = QMessageBox.question(
                self,
                "Outlook解析中",
                "Outlookの解析中です。\n"
                "解析をキャンセルして終了しますか？",
                (
                    QMessageBox.StandardButton.Yes
                    | QMessageBox.StandardButton.No
                ),
                QMessageBox.StandardButton.No,
            )

            if (
                answer
                != QMessageBox.StandardButton.Yes
            ):
                event.ignore()
                return

            self._close_after_scan = True
            self._cancel_scan()

            event.ignore()
            return

        if self.dirty:
            answer = QMessageBox.question(
                self,
                "未保存の変更",
                "設定が保存されていません。\n"
                "変更を破棄して終了しますか？",
                (
                    QMessageBox.StandardButton.Yes
                    | QMessageBox.StandardButton.No
                ),
                QMessageBox.StandardButton.No,
            )

            if (
                answer
                != QMessageBox.StandardButton.Yes
            ):
                event.ignore()
                return

        event.accept()
