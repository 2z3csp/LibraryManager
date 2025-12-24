# library_manager_mock.py
# PySide6 mock for folder-based document versioning on OneDrive-synced folders.
# - Register folders (G1)
# - Show registered folders list (G0): name/link, latest rev, last edited date, editor
# - Show files in selected folder
# - Update (create next rev, move old to _History, write metadata + memo)
# - Replace (take external file, store as next rev, move old to _History, write metadata + memo)
#
# Notes:
# - Keeps your existing folder structure: "latest files" live in the registered folder root.
# - Uses "_History" folder under the registered folder.
# - Stores metadata under "_Meta" folder under the registered folder.
#
# Tested targets: Windows, OneDrive local sync folder.

from __future__ import annotations

import json
import os
import re
import shutil
import sys
import datetime as dt
import getpass
from dataclasses import dataclass
from typing import Optional, Tuple, Dict, Any, List

from PySide6.QtCore import Qt, QSize, QTimer
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QToolBar,
    QVBoxLayout,
    QHBoxLayout,
    QSplitter,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QDialog,
    QLabel,
    QFormLayout,
    QPlainTextEdit,
    QDialogButtonBox,
    QHeaderView,
    QAbstractItemView,
    QGroupBox,
    QTreeWidget,
    QTreeWidgetItem,
    QMenu,
    QComboBox,
    QSpinBox,
)

REV_RE = re.compile(
    r"^(?P<base>.+)_rev(?P<A>\d+)\.(?P<B>\d+)\.(?P<C>\d+)_(?P<date>\d{8})$",
    re.IGNORECASE,
)

TEMP_FILE_RE = re.compile(r"^~\$")  # Office temporary files


def now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def today_yyyymmdd() -> str:
    return dt.date.today().strftime("%Y%m%d")


def user_name() -> str:
    return getpass.getuser()


def appdata_dir() -> str:
    base = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "TeamLibraryManagerMock")
    os.makedirs(path, exist_ok=True)
    return path


REGISTRY_PATH = os.path.join(appdata_dir(), "registry.json")
SETTINGS_PATH = os.path.join(appdata_dir(), "settings.json")
DEFAULT_MEMO_TIMEOUT_MIN = 30


def load_registry() -> List[Dict[str, str]]:
    if not os.path.exists(REGISTRY_PATH):
        return []
    try:
        with open(REGISTRY_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            # normalize
            out = []
            for x in data:
                if isinstance(x, dict) and "name" in x and "path" in x:
                    out.append({
                        "name": str(x["name"]),
                        "path": str(x["path"]),
                        "main_category": str(x.get("main_category", "")),
                        "sub_category": str(x.get("sub_category", "")),
                    })
            return out
        return []
    except Exception:
        return []


def save_registry(items: List[Dict[str, str]]) -> None:
    with open(REGISTRY_PATH, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)


def load_settings() -> Dict[str, Any]:
    defaults = {"memo_timeout_min": DEFAULT_MEMO_TIMEOUT_MIN}
    if not os.path.exists(SETTINGS_PATH):
        return defaults
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return defaults
        out = defaults.copy()
        out.update(data)
        return out
    except Exception:
        return defaults


def save_settings(settings: Dict[str, Any]) -> None:
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def ensure_folder_structure(folder_path: str) -> Tuple[str, str]:
    """
    Returns (history_dir, meta_dir)
    """
    history_dir = os.path.join(folder_path, "_History")
    meta_dir = os.path.join(folder_path, "_Meta")
    os.makedirs(history_dir, exist_ok=True)
    os.makedirs(meta_dir, exist_ok=True)
    return history_dir, meta_dir


def meta_path_for_folder(folder_path: str) -> str:
    _, meta_dir = ensure_folder_structure(folder_path)
    return os.path.join(meta_dir, "docmeta.json")


def load_meta(folder_path: str) -> Dict[str, Any]:
    p = meta_path_for_folder(folder_path)
    if not os.path.exists(p):
        return {"documents": {}}
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {"documents": {}}
        if "documents" not in data or not isinstance(data["documents"], dict):
            data["documents"] = {}
        return data
    except Exception:
        return {"documents": {}}


def save_meta(folder_path: str, meta: Dict[str, Any]) -> None:
    p = meta_path_for_folder(folder_path)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


def split_name_ext(filename: str) -> Tuple[str, str]:
    base, ext = os.path.splitext(filename)
    return base, ext


def parse_rev_from_filename(filename: str) -> Tuple[str, Optional[Tuple[int, int, int, int]], str]:
    """
    Returns (doc_base_name, version_tuple_or_None, rev_str_or_empty)
    version tuple: (A, B, C, YYYYMMDD as int)
    """
    base, ext = split_name_ext(filename)
    m = REV_RE.match(base)
    if not m:
        return base + ext, None, ""
    doc_base = m.group("base") + ext
    A = int(m.group("A"))
    B = int(m.group("B"))
    C = int(m.group("C"))
    d = int(m.group("date"))
    rev_str = f"rev{A}.{B}.{C}_{m.group('date')}"
    return doc_base, (A, B, C, d), rev_str


def format_rev(A: int, B: int, C: int, yyyymmdd: str) -> str:
    return f"rev{A}.{B}.{C}_{yyyymmdd}"


def next_rev_from_current(current_rev: str) -> Tuple[int, int, int]:
    """
    current_rev like 'rev1.2.3_20251223' or ''
    Returns (A,B,C) for the next revision, defaulting to (1,0,1) if none.
    Behavior: keep A,B; increment C by 1. If no current_rev, start A=1,B=0,C=1.
    """
    if not current_rev:
        return 1, 0, 1
    m = re.match(r"^rev(\d+)\.(\d+)\.(\d+)_\d{8}$", current_rev, re.IGNORECASE)
    if not m:
        return 1, 0, 1
    A, B, C = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return A, B, C + 1


def safe_list_files(folder_path: str) -> List[str]:
    files = []
    try:
        for name in os.listdir(folder_path):
            full = os.path.join(folder_path, name)
            if os.path.isdir(full):
                # skip management folders
                if name.lower() in {"_history", "_meta"}:
                    continue
                continue
            if TEMP_FILE_RE.match(name):
                continue
            files.append(name)
    except Exception:
        pass
    return files


def is_file_locked(path: str) -> bool:
    try:
        with open(path, "r+b") as f:
            if os.name == "nt":
                import msvcrt

                try:
                    msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                    msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                except OSError:
                    return True
        return False
    except OSError:
        return True


@dataclass
class FileRow:
    filename: str
    doc_key: str  # base name without rev (with extension)
    rev: str
    updated_at: str
    updated_by: str
    memo: str


def scan_folder(folder_path: str) -> Tuple[Dict[str, Any], List[FileRow]]:
    """
    Reads meta if exists and also scans real files.
    Returns (meta, file_rows_for_view).
    """
    meta = load_meta(folder_path)
    docs: Dict[str, Any] = meta.get("documents", {})
    files = safe_list_files(folder_path)

    # Build latest candidate per doc_key from filesystem (in case meta is missing/outdated)
    latest_by_doc: Dict[str, Tuple[Optional[Tuple[int, int, int, int]], str]] = {}
    for fn in files:
        doc_key, ver_tuple, rev_str = parse_rev_from_filename(fn)
        prev = latest_by_doc.get(doc_key)
        if prev is None:
            latest_by_doc[doc_key] = (ver_tuple, fn)
        else:
            prev_tuple, prev_fn = prev
            # compare: version tuple if available, else keep existing if it has tuple
            if prev_tuple is None and ver_tuple is not None:
                latest_by_doc[doc_key] = (ver_tuple, fn)
            elif prev_tuple is not None and ver_tuple is not None:
                if ver_tuple > prev_tuple:
                    latest_by_doc[doc_key] = (ver_tuple, fn)
            elif prev_tuple is None and ver_tuple is None:
                # fallback: lexicographic
                if fn > prev_fn:
                    latest_by_doc[doc_key] = (ver_tuple, fn)

    # Merge into meta as "observed current" if meta lacks it
    changed = False
    for doc_key, (_t, fn) in latest_by_doc.items():
        if doc_key not in docs:
            docs[doc_key] = {
                "title": doc_key,
                "current_file": fn,
                "current_rev": parse_rev_from_filename(fn)[2],
                "updated_at": "",
                "updated_by": "",
                "last_memo": "",
                "history": [],
            }
            changed = True
        else:
            # If current_file missing or no longer exists, refresh from scan
            cur_fn = docs[doc_key].get("current_file", "")
            if not cur_fn or cur_fn not in files:
                docs[doc_key]["current_file"] = fn
                docs[doc_key]["current_rev"] = parse_rev_from_filename(fn)[2]
                changed = True

    meta["documents"] = docs
    if changed:
        save_meta(folder_path, meta)

    # Build view rows (latest per doc)
    rows: List[FileRow] = []
    for doc_key, info in docs.items():
        cur_fn = info.get("current_file", "")
        if not cur_fn:
            continue
        rows.append(
            FileRow(
                filename=cur_fn,
                doc_key=doc_key,
                rev=info.get("current_rev", ""),
                updated_at=info.get("updated_at", ""),
                updated_by=info.get("updated_by", ""),
                memo=info.get("last_memo", ""),
            )
        )
    # Sort by updated_at desc then filename
    rows.sort(key=lambda r: (r.updated_at or "", r.filename), reverse=True)
    return meta, rows


class RegisterDialog(QDialog):
    def __init__(
        self,
        parent: QWidget | None = None,
        initial_name: str = "",
        initial_path: str = "",
        initial_main_category: str = "",
        initial_sub_category: str = "",
        main_category_options: Optional[List[str]] = None,
        sub_category_options: Optional[List[str]] = None,
        ok_label: str = "登録",
    ):
        super().__init__(parent)
        self.setWindowTitle("登録（G1）")
        self.setMinimumWidth(520)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.path_edit = QLineEdit()
        self.path_btn = QPushButton("参照...")
        self.path_btn.clicked.connect(self.pick_folder)

        path_row = QHBoxLayout()
        path_row.addWidget(self.path_edit, 1)
        path_row.addWidget(self.path_btn)

        self.name_edit = QLineEdit()
        self.main_category_edit = QComboBox()
        self.main_category_edit.setEditable(True)
        self.sub_category_edit = QComboBox()
        self.sub_category_edit.setEditable(True)
        if self.main_category_edit.lineEdit():
            self.main_category_edit.lineEdit().setPlaceholderText("例：案件　記入 or 選択")
        if self.sub_category_edit.lineEdit():
            self.sub_category_edit.lineEdit().setPlaceholderText("例：図面　記入 or 選択")

        if main_category_options:
            self.main_category_edit.addItems(main_category_options)
        if sub_category_options:
            self.sub_category_edit.addItems(sub_category_options)
        self.main_category_edit.setCurrentIndex(-1)
        self.sub_category_edit.setCurrentIndex(-1)

        form.addRow("登録フォルダパス", path_row)
        form.addRow("登録名", self.name_edit)
        form.addRow("メインカテゴリ", self.main_category_edit)
        form.addRow("サブカテゴリ", self.sub_category_edit)

        if initial_path:
            self.path_edit.setText(initial_path)
        if initial_name:
            self.name_edit.setText(initial_name)
        if initial_main_category:
            self.main_category_edit.setCurrentText(initial_main_category)
        if initial_sub_category:
            self.sub_category_edit.setCurrentText(initial_sub_category)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText(ok_label)
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def pick_folder(self):
        path = QFileDialog.getExistingDirectory(self, "登録するフォルダを選択")
        if path:
            self.path_edit.setText(path)
            if not self.name_edit.text().strip():
                self.name_edit.setText(os.path.basename(path))

    def get_values(self) -> Tuple[str, str, str, str]:
        return (
            self.name_edit.text().strip(),
            self.path_edit.text().strip(),
            self.main_category_edit.currentText().strip(),
            self.sub_category_edit.currentText().strip(),
        )


class MemoDialog(QDialog):
    def __init__(self, title: str, subtitle: str, default_memo: str = "", parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(560)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(subtitle))

        self.memo = QPlainTextEdit()
        self.memo.setPlaceholderText("作業メモ（空欄可）")
        self.memo.setPlainText(default_memo)
        layout.addWidget(self.memo, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("実行")
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_memo(self) -> str:
        return self.memo.toPlainText().strip()


class ReplaceDialog(QDialog):
    def __init__(self, target_filename: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("差し替え")
        self.setMinimumWidth(640)

        self.target_filename = target_filename

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"差し替え対象（現行）: {target_filename}"))

        self.file_edit = QLineEdit()
        self.file_btn = QPushButton("ファイル選択...")
        self.file_btn.clicked.connect(self.pick_file)

        row = QHBoxLayout()
        row.addWidget(self.file_edit, 1)
        row.addWidget(self.file_btn)
        layout.addLayout(row)

        self.memo = QPlainTextEdit()
        self.memo.setPlaceholderText("作業メモ（空欄可） 例：外注成果品差し替え、図面反映 等")
        layout.addWidget(self.memo, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("差し替え実行")
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "差し替えるファイルを選択")
        if path:
            self.file_edit.setText(path)

    def get_values(self) -> Tuple[str, str]:
        return self.file_edit.text().strip(), self.memo.toPlainText().strip()


class OptionsDialog(QDialog):
    def __init__(self, timeout_min: int, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("オプション")
        self.setMinimumWidth(360)

        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(1, 240)
        self.timeout_spin.setValue(timeout_min)
        form.addRow("メモ入力タイムアウト（分）", self.timeout_spin)
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("保存")
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_timeout_min(self) -> int:
        return int(self.timeout_spin.value())


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("チーム間 図書管理モック（G0）")
        self.resize(1300, 760)

        self.registry = load_registry()
        self.settings = load_settings()
        self.memo_timeout_min = int(self.settings.get("memo_timeout_min", DEFAULT_MEMO_TIMEOUT_MIN))

        root = QWidget()
        self.setCentralWidget(root)
        root_layout = QVBoxLayout(root)

        toolbar = QToolBar("ツール")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        act_options = QAction("オプション", self)
        act_options.triggered.connect(self.on_options)
        toolbar.addAction(act_options)

        # Top controls
        top = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("検索（登録名でフィルタ）")
        self.search.textChanged.connect(self.refresh_folder_table)

        btn_register = QPushButton("登録")
        btn_register.clicked.connect(self.on_register)

        btn_update = QPushButton("更新")
        btn_update.clicked.connect(self.on_update)

        btn_replace = QPushButton("差し替え")
        btn_replace.clicked.connect(self.on_replace)

        btn_rescan = QPushButton("再スキャン")
        btn_rescan.clicked.connect(self.on_rescan)

        top.addWidget(self.search, 1)
        top.addWidget(btn_register)
        top.addWidget(btn_update)
        top.addWidget(btn_replace)
        top.addWidget(btn_rescan)

        root_layout.addLayout(top)

        splitter = QSplitter(Qt.Horizontal)
        root_layout.addWidget(splitter, 1)

        # Left: category tree
        tree_box = QWidget()
        tree_layout = QVBoxLayout(tree_box)
        tree_layout.setContentsMargins(0, 0, 0, 0)

        self.category_tree = QTreeWidget()
        self.category_tree.setHeaderHidden(True)
        self.category_tree.itemSelectionChanged.connect(self.on_category_tree_selected)
        self.category_tree.itemDoubleClicked.connect(self.on_category_tree_double_clicked)
        self.category_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.category_tree.customContextMenuRequested.connect(self.on_category_tree_context_menu)
        tree_layout.addWidget(self.category_tree)
        splitter.addWidget(tree_box)

        # Left-middle: registered folders table
        left_box = QWidget()
        left_layout = QVBoxLayout(left_box)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.folders_table = QTableWidget(0, 2)
        self.folders_table.setHorizontalHeaderLabels(["登録名（ダブルクリックで開く）", "最終更新日"])
        self.folders_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.folders_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.folders_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.folders_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.folders_table.itemSelectionChanged.connect(self.on_folder_selected)
        self.folders_table.itemDoubleClicked.connect(self.on_folder_double_clicked)
        self.folders_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.folders_table.customContextMenuRequested.connect(self.on_folders_table_context_menu)

        left_layout.addWidget(self.folders_table)
        splitter.addWidget(left_box)

        # Middle: files table
        mid_box = QWidget()
        mid_layout = QVBoxLayout(mid_box)
        mid_layout.setContentsMargins(0, 0, 0, 0)

        self.files_table = QTableWidget(0, 5)
        self.files_table.setHorizontalHeaderLabels(["ファイル（最新）", "rev", "更新日", "更新者", "DocKey"])
        self.files_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.files_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.files_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.files_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.files_table.itemSelectionChanged.connect(self.on_file_selected)

        mid_layout.addWidget(self.files_table)
        splitter.addWidget(mid_box)

        # Right: memo + history
        right_box = QWidget()
        right_layout = QVBoxLayout(right_box)
        right_layout.setContentsMargins(0, 0, 0, 0)

        memo_group = QGroupBox("最新メモ")
        memo_layout = QVBoxLayout(memo_group)
        self.memo_view = QPlainTextEdit()
        self.memo_view.setReadOnly(True)
        self.memo_view.setPlaceholderText("ここに最新メモを表示")
        memo_layout.addWidget(self.memo_view)
        right_layout.addWidget(memo_group, 1)

        hist_group = QGroupBox("履歴（更新日・人・メモ）")
        hist_layout = QVBoxLayout(hist_group)
        self.hist_table = QTableWidget(0, 3)
        self.hist_table.setHorizontalHeaderLabels(["更新日時", "更新者", "メモ"])
        self.hist_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.hist_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.hist_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        hist_layout.addWidget(self.hist_table)
        right_layout.addWidget(hist_group, 1)

        splitter.addWidget(right_box)
        splitter.setSizes([260, 440, 520, 360])

        # State
        self.current_folder: Optional[Dict[str, str]] = None
        self.current_meta: Optional[Dict[str, Any]] = None
        self.current_file_rows: List[FileRow] = []
        self.selected_main_category: str = ""
        self.selected_sub_category: str = ""
        self.watch_timer = None
        self.watch_target_path = ""
        self.watch_folder_path = ""
        self.watch_doc_key = ""
        self.watch_started_at: Optional[dt.datetime] = None

        self.refresh_folder_table()
        self.refresh_category_tree()

    # ---------- UI helpers ----------
    def info(self, msg: str):
        QMessageBox.information(self, "情報", msg)

    def warn(self, msg: str):
        QMessageBox.warning(self, "注意", msg)

    def ask(self, msg: str) -> bool:
        return QMessageBox.question(self, "確認", msg, QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes

    def selected_folder_index(self) -> int:
        sel = self.folders_table.selectionModel().selectedRows()
        return sel[0].row() if sel else -1

    def selected_file_index(self) -> int:
        sel = self.files_table.selectionModel().selectedRows()
        return sel[0].row() if sel else -1

    def open_in_explorer(self, path: str):
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as e:
            self.warn(f"エクスプローラーで開けませんでした: {e}")

    def category_fallback(self, value: str, fallback: str) -> str:
        return value.strip() or fallback

    def category_options(self) -> Tuple[List[str], List[str]]:
        main_categories = {
            self.category_fallback(item.get("main_category", ""), "未分類")
            for item in self.registry
        }
        sub_categories = {
            self.category_fallback(item.get("sub_category", ""), "未分類")
            for item in self.registry
        }
        return sorted(main_categories), sorted(sub_categories)

    def registry_index_by_path(self, path: str) -> int:
        for idx, item in enumerate(self.registry):
            if os.path.normcase(item["path"]) == os.path.normcase(path):
                return idx
        return -1

    def edit_registered_folder(self, path: str):
        idx = self.registry_index_by_path(path)
        if idx < 0:
            self.warn("登録情報が見つかりません。")
            return
        item = self.registry[idx]
        main_options, sub_options = self.category_options()

        dlg = RegisterDialog(
            self,
            initial_name=item.get("name", ""),
            initial_path=item.get("path", ""),
            initial_main_category=item.get("main_category", ""),
            initial_sub_category=item.get("sub_category", ""),
            main_category_options=main_options,
            sub_category_options=sub_options,
            ok_label="更新",
        )
        if dlg.exec() != QDialog.Accepted:
            return

        name, new_path, main_category, sub_category = dlg.get_values()
        if not name or not new_path or not main_category or not sub_category:
            self.warn("登録名・フォルダパス・メインカテゴリ・サブカテゴリを入力してください。")
            return
        if not os.path.isdir(new_path):
            self.warn("フォルダが存在しません。")
            return

        old_path = item["path"]
        if os.path.normcase(new_path) != os.path.normcase(old_path):
            if self.registry_index_by_path(new_path) >= 0:
                self.warn("このフォルダは既に登録されています。")
                return
            new_history = os.path.join(new_path, "_History")
            new_meta = os.path.join(new_path, "_Meta")
            if os.path.exists(new_history) or os.path.exists(new_meta):
                self.warn("移動先に _History または _Meta が既に存在します。更新をキャンセルしました。")
                return
            old_history = os.path.join(old_path, "_History")
            old_meta = os.path.join(old_path, "_Meta")
            try:
                if os.path.exists(old_history):
                    shutil.move(old_history, new_history)
                if os.path.exists(old_meta):
                    shutil.move(old_meta, new_meta)
            except Exception as e:
                self.warn(f"_History/_Meta の移動に失敗しました: {e}")
                return

        ensure_folder_structure(new_path)

        item["name"] = name
        item["path"] = new_path
        item["main_category"] = main_category
        item["sub_category"] = sub_category
        self.registry[idx] = item
        save_registry(self.registry)
        self.refresh_folder_table()
        self.refresh_category_tree()
        if self.current_folder and os.path.normcase(self.current_folder["path"]) == os.path.normcase(old_path):
            self.current_folder = {"name": name, "path": new_path}
            self.refresh_files_table()
        self.info("更新しました。")

    def delete_registered_folder(self, path: str):
        idx = self.registry_index_by_path(path)
        if idx < 0:
            self.warn("登録情報が見つかりません。")
            return
        if not self.ask("この登録を削除しますか？（_History/_Meta は削除しません）"):
            return
        self.registry.pop(idx)
        save_registry(self.registry)
        if self.current_folder and os.path.normcase(self.current_folder["path"]) == os.path.normcase(path):
            self.current_folder = None
            self.current_meta = None
            self.current_file_rows = []
        self.refresh_folder_table()
        self.refresh_category_tree()
        self.refresh_files_table()
        self.info("削除しました。")

    # ---------- refresh ----------
    def refresh_folder_table(self):
        query = self.search.text().strip().lower()
        items = []
        for x in self.registry:
            if query not in x["name"].lower():
                continue
            main_cat = self.category_fallback(x.get("main_category", ""), "未分類")
            sub_cat = self.category_fallback(x.get("sub_category", ""), "未分類")
            if self.selected_main_category and main_cat != self.selected_main_category:
                continue
            if self.selected_sub_category and sub_cat != self.selected_sub_category:
                continue
            items.append(x)

        self.folders_table.setRowCount(0)

        for item in items:
            name = item["name"]
            path = item["path"]
            last_date = ""

            if os.path.isdir(path):
                try:
                    files = safe_list_files(path)
                    latest_mtime = None
                    for filename in files:
                        file_path = os.path.join(path, filename)
                        try:
                            mtime = os.path.getmtime(file_path)
                        except Exception:
                            continue
                        if latest_mtime is None or mtime > latest_mtime:
                            latest_mtime = mtime
                    if latest_mtime is not None:
                        last_date = dt.datetime.fromtimestamp(latest_mtime).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    pass

            r = self.folders_table.rowCount()
            self.folders_table.insertRow(r)

            it_name = QTableWidgetItem(name)
            it_name.setToolTip(path)
            it_name.setData(Qt.UserRole, path)

            self.folders_table.setItem(r, 0, it_name)
            self.folders_table.setItem(r, 1, QTableWidgetItem(last_date))

        self.folders_table.repaint()

    def refresh_category_tree(self):
        self.category_tree.clear()
        categories: Dict[str, Dict[str, List[Dict[str, str]]]] = {}
        for item in self.registry:
            main_cat = self.category_fallback(item.get("main_category", ""), "未分類")
            sub_cat = self.category_fallback(item.get("sub_category", ""), "未分類")
            categories.setdefault(main_cat, {}).setdefault(sub_cat, []).append(item)

        for main_cat in sorted(categories.keys()):
            main_item = QTreeWidgetItem([main_cat])
            main_item.setData(0, Qt.UserRole, {"type": "main", "main": main_cat})
            self.category_tree.addTopLevelItem(main_item)
            for sub_cat in sorted(categories[main_cat].keys()):
                sub_item = QTreeWidgetItem([sub_cat])
                sub_item.setData(0, Qt.UserRole, {"type": "sub", "main": main_cat, "sub": sub_cat})
                main_item.addChild(sub_item)
                for folder in sorted(categories[main_cat][sub_cat], key=lambda x: x["name"].lower()):
                    folder_item = QTreeWidgetItem([folder["name"]])
                    folder_item.setToolTip(0, folder["path"])
                    folder_item.setData(0, Qt.UserRole, {
                        "type": "folder",
                        "main": main_cat,
                        "sub": sub_cat,
                        "path": folder["path"],
                    })
                    sub_item.addChild(folder_item)

        self.category_tree.expandAll()

    def select_folder_in_table(self, path: str):
        for row in range(self.folders_table.rowCount()):
            item = self.folders_table.item(row, 0)
            if item and item.data(Qt.UserRole) == path:
                self.folders_table.selectRow(row)
                break

    def refresh_files_table(self):
        self.files_table.setRowCount(0)
        self.memo_view.setPlainText("")
        self.hist_table.setRowCount(0)

        if not self.current_folder:
            return

        folder_path = self.current_folder["path"]
        if not os.path.isdir(folder_path):
            self.warn("登録フォルダが見つかりません。パスを確認してください。")
            return

        meta, rows = scan_folder(folder_path)
        self.current_meta = meta
        self.current_file_rows = rows

        for row in rows:
            r = self.files_table.rowCount()
            self.files_table.insertRow(r)

            it_fn = QTableWidgetItem(row.filename)
            it_fn.setToolTip(os.path.join(folder_path, row.filename))
            self.files_table.setItem(r, 0, it_fn)
            self.files_table.setItem(r, 1, QTableWidgetItem(row.rev))
            self.files_table.setItem(r, 2, QTableWidgetItem((row.updated_at or "").replace("T", " ")))
            self.files_table.setItem(r, 3, QTableWidgetItem(row.updated_by or ""))
            self.files_table.setItem(r, 4, QTableWidgetItem(row.doc_key))

        # hide DocKey column by default (can be useful for debugging)
        self.files_table.setColumnHidden(4, True)

    def refresh_right_pane_for_doc(self, doc_key: str):
        self.memo_view.setPlainText("")
        self.hist_table.setRowCount(0)

        if not self.current_meta:
            return
        docs = self.current_meta.get("documents", {})
        info = docs.get(doc_key)
        if not isinstance(info, dict):
            return

        self.memo_view.setPlainText(info.get("last_memo", "") or "")

        history = info.get("history", [])
        if not isinstance(history, list):
            return

        # show newest first
        items = list(history)
        items.reverse()

        for h in items:
            r = self.hist_table.rowCount()
            self.hist_table.insertRow(r)
            self.hist_table.setItem(r, 0, QTableWidgetItem((h.get("updated_at", "") or "").replace("T", " ")))
            self.hist_table.setItem(r, 1, QTableWidgetItem(h.get("updated_by", "") or ""))
            self.hist_table.setItem(r, 2, QTableWidgetItem(h.get("memo", "") or ""))

    # ---------- events ----------
    def on_folders_table_context_menu(self, pos):
        item = self.folders_table.itemAt(pos)
        if not item:
            return
        row = item.row()
        path_item = self.folders_table.item(row, 0)
        if not path_item:
            return
        path = path_item.data(Qt.UserRole)
        if not path:
            return

        menu = QMenu(self)
        act_edit = menu.addAction("編集")
        act_delete = menu.addAction("削除")
        action = menu.exec(self.folders_table.viewport().mapToGlobal(pos))
        if action == act_edit:
            self.edit_registered_folder(path)
        elif action == act_delete:
            self.delete_registered_folder(path)

    def on_category_tree_context_menu(self, pos):
        item = self.category_tree.itemAt(pos)
        if not item:
            return
        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")

        menu = QMenu(self)
        act_register = None
        act_edit = None
        act_delete = None

        if item_type in {"main", "sub"}:
            act_register = menu.addAction("登録")
        elif item_type == "folder":
            act_edit = menu.addAction("編集")
            act_delete = menu.addAction("削除")
        else:
            return

        action = menu.exec(self.category_tree.viewport().mapToGlobal(pos))
        if action == act_register:
            initial_main = data.get("main", "")
            initial_sub = data.get("sub", "") if item_type == "sub" else ""
            self.open_register_dialog(initial_main_category=initial_main, initial_sub_category=initial_sub)
        elif action == act_edit:
            path = data.get("path")
            if path:
                self.edit_registered_folder(path)
        elif action == act_delete:
            path = data.get("path")
            if path:
                self.delete_registered_folder(path)

    def open_register_dialog(self, initial_main_category: str = "", initial_sub_category: str = ""):
        main_options, sub_options = self.category_options()
        dlg = RegisterDialog(
            self,
            initial_main_category=initial_main_category,
            initial_sub_category=initial_sub_category,
            main_category_options=main_options,
            sub_category_options=sub_options,
        )
        if dlg.exec() != QDialog.Accepted:
            return
        name, path, main_category, sub_category = dlg.get_values()
        if not name or not path or not main_category or not sub_category:
            self.warn("登録名・フォルダパス・メインカテゴリ・サブカテゴリを入力してください。")
            return
        if not os.path.isdir(path):
            self.warn("フォルダが存在しません。")
            return

        # Ensure _History/_Meta
        ensure_folder_structure(path)

        # Prevent duplicates by path
        for x in self.registry:
            if os.path.normcase(x["path"]) == os.path.normcase(path):
                self.warn("このフォルダは既に登録されています。")
                return

        self.registry.append({
            "name": name,
            "path": path,
            "main_category": main_category,
            "sub_category": sub_category,
        })
        save_registry(self.registry)
        self.refresh_folder_table()
        self.refresh_category_tree()
        self.info("登録しました。")

    def on_register(self):
        self.open_register_dialog()

    def on_folder_selected(self):
        idx = self.selected_folder_index()
        if idx < 0:
            self.current_folder = None
            self.current_meta = None
            self.current_file_rows = []
            self.files_table.setRowCount(0)
            return

        # Determine which registry item matches current filtered table row:
        # we stored path in tooltip/userrole for the name item
        it = self.folders_table.item(idx, 0)
        if not it:
            return
        path = it.data(Qt.UserRole)
        name = it.text()
        self.current_folder = {"name": name, "path": path}
        self.refresh_files_table()

    def on_folder_double_clicked(self, item: QTableWidgetItem):
        if item.column() != 0:
            return
        path = item.data(Qt.UserRole)
        if path and os.path.isdir(path):
            self.open_in_explorer(path)

    def on_category_tree_selected(self):
        items = self.category_tree.selectedItems()
        if not items:
            return
        data = items[0].data(0, Qt.UserRole) or {}
        item_type = data.get("type")
        if item_type == "main":
            self.selected_main_category = data.get("main", "")
            self.selected_sub_category = ""
            self.refresh_folder_table()
        elif item_type == "sub":
            self.selected_main_category = data.get("main", "")
            self.selected_sub_category = data.get("sub", "")
            self.refresh_folder_table()
        elif item_type == "folder":
            self.selected_main_category = data.get("main", "")
            self.selected_sub_category = data.get("sub", "")
            self.refresh_folder_table()
            self.select_folder_in_table(data.get("path", ""))

    def on_category_tree_double_clicked(self, item: QTreeWidgetItem):
        data = item.data(0, Qt.UserRole) or {}
        if data.get("type") != "folder":
            return
        path = data.get("path")
        if path and os.path.isdir(path):
            self.open_in_explorer(path)

    def on_file_selected(self):
        idx = self.selected_file_index()
        if idx < 0:
            return
        if idx >= len(self.current_file_rows):
            return
        row = self.current_file_rows[idx]
        self.refresh_right_pane_for_doc(row.doc_key)

    def on_rescan(self):
        self.refresh_folder_table()
        self.refresh_files_table()
        self.refresh_category_tree()

    def on_options(self):
        dlg = OptionsDialog(self.memo_timeout_min, self)
        if dlg.exec() != QDialog.Accepted:
            return
        self.memo_timeout_min = dlg.get_timeout_min()
        self.settings["memo_timeout_min"] = self.memo_timeout_min
        save_settings(self.settings)
        self.info("設定を保存しました。")

    # ---------- core operations ----------
    def _get_selected_doc(self) -> Optional[Tuple[str, str, Dict[str, Any]]]:
        """
        Returns (folder_path, doc_key, doc_info)
        """
        if not self.current_folder or not self.current_meta:
            self.warn("フォルダを選択してください。")
            return None
        file_idx = self.selected_file_index()
        if file_idx < 0:
            self.warn("対象ファイル（最新）を選択してください。")
            return None
        if file_idx >= len(self.current_file_rows):
            return None
        row = self.current_file_rows[file_idx]
        doc_key = row.doc_key
        docs = self.current_meta.get("documents", {})
        info = docs.get(doc_key)
        if not isinstance(info, dict):
            self.warn("メタデータが見つかりません。再スキャンしてください。")
            return None
        return self.current_folder["path"], doc_key, info

    def on_update(self):
        sel = self._get_selected_doc()
        if not sel:
            return
        folder_path, doc_key, info = sel

        cur_fn = info.get("current_file", "")
        cur_rev = info.get("current_rev", "")
        if not cur_fn:
            self.warn("現行ファイルが不明です。")
            return

        base_name, ext = split_name_ext(cur_fn)
        doc_base, ver_tuple, _rev_str = parse_rev_from_filename(cur_fn)
        # doc_base includes ext; we want base without ext for naming
        doc_base_no_ext, _ = split_name_ext(doc_base)

        A, B, C = next_rev_from_current(cur_rev)
        new_rev = format_rev(A, B, C, today_yyyymmdd())
        new_fn = f"{doc_base_no_ext}_{new_rev}{ext}"

        history_dir, _ = ensure_folder_structure(folder_path)

        cur_path = os.path.join(folder_path, cur_fn)
        new_path = os.path.join(folder_path, new_fn)

        if not os.path.exists(cur_path):
            self.warn("現行ファイルが見つかりません。")
            return
        if os.path.exists(new_path):
            self.warn("新規ファイル名が既に存在します。再度実行してください（Cが進みます）。")
            return

        try:
            # 1) copy current -> new
            shutil.copy2(cur_path, new_path)

            # 2) move old current -> _History
            hist_name = cur_fn
            hist_path = os.path.join(history_dir, hist_name)
            # avoid collision in _History
            if os.path.exists(hist_path):
                ts = dt.datetime.now().strftime("%Y%m%d%H%M%S")
                hist_name = f"{split_name_ext(cur_fn)[0]}_{ts}{split_name_ext(cur_fn)[1]}"
                hist_path = os.path.join(history_dir, hist_name)
            shutil.move(cur_path, hist_path)

            # 3) update meta
            meta = load_meta(folder_path)
            docs = meta.get("documents", {})
            if doc_key not in docs:
                docs[doc_key] = {
                    "title": doc_key,
                    "current_file": new_fn,
                    "current_rev": new_rev,
                    "updated_at": now_iso(),
                    "updated_by": user_name(),
                    "last_memo": "",
                    "history": [],
                }
            else:
                d = docs[doc_key]
                # push previous current into history list
                prev_entry = {
                    "rev": cur_rev,
                    "file": hist_name,
                    "updated_at": d.get("updated_at", ""),
                    "updated_by": d.get("updated_by", ""),
                    "memo": d.get("last_memo", ""),
                }
                d.setdefault("history", [])
                if isinstance(d["history"], list):
                    d["history"].append(prev_entry)

                d["current_file"] = new_fn
                d["current_rev"] = new_rev
                d["updated_at"] = now_iso()
                d["updated_by"] = user_name()
                d["last_memo"] = ""

            meta["documents"] = docs
            save_meta(folder_path, meta)

            # Refresh
            self.refresh_files_table()
            self.refresh_folder_table()

            # Open the new file for convenience
            try:
                os.startfile(new_path)  # type: ignore[attr-defined]
            except Exception:
                pass
            self.start_file_lock_watch(new_path, folder_path, doc_key)

        except Exception as e:
            self.warn(f"更新に失敗しました: {e}")

    def on_replace(self):
        sel = self._get_selected_doc()
        if not sel:
            return
        folder_path, doc_key, info = sel

        cur_fn = info.get("current_file", "")
        cur_rev = info.get("current_rev", "")
        if not cur_fn:
            self.warn("現行ファイルが不明です。")
            return

        dlg = ReplaceDialog(cur_fn, self)
        if dlg.exec() != QDialog.Accepted:
            return
        incoming_path, memo = dlg.get_values()
        if not incoming_path or not os.path.exists(incoming_path):
            self.warn("差し替えるファイルを選択してください。")
            return

        history_dir, _ = ensure_folder_structure(folder_path)

        cur_path = os.path.join(folder_path, cur_fn)
        if not os.path.exists(cur_path):
            self.warn("現行ファイルが見つかりません。")
            return

        # Determine next rev and destination filename (keep doc base)
        doc_base, _t, _r = parse_rev_from_filename(cur_fn)
        doc_base_no_ext, ext = split_name_ext(doc_base)

        A, B, C = next_rev_from_current(cur_rev)
        new_rev = format_rev(A, B, C, today_yyyymmdd())
        dest_fn = f"{doc_base_no_ext}_{new_rev}{ext}"
        dest_path = os.path.join(folder_path, dest_fn)

        if os.path.exists(dest_path):
            self.warn("新規ファイル名が既に存在します。再度実行してください（Cが進みます）。")
            return

        try:
            # 1) move old current -> _History
            hist_name = cur_fn
            hist_path = os.path.join(history_dir, hist_name)
            if os.path.exists(hist_path):
                ts = dt.datetime.now().strftime("%Y%m%d%H%M%S")
                hist_name = f"{split_name_ext(cur_fn)[0]}_{ts}{split_name_ext(cur_fn)[1]}"
                hist_path = os.path.join(history_dir, hist_name)
            shutil.move(cur_path, hist_path)

            # 2) copy incoming -> dest
            shutil.copy2(incoming_path, dest_path)

            # 3) update meta
            meta = load_meta(folder_path)
            docs = meta.get("documents", {})
            d = docs.get(doc_key, {
                "title": doc_key,
                "current_file": dest_fn,
                "current_rev": new_rev,
                "updated_at": now_iso(),
                "updated_by": user_name(),
                "last_memo": memo,
                "history": [],
            })

            prev_entry = {
                "rev": cur_rev,
                "file": hist_name,
                "updated_at": d.get("updated_at", ""),
                "updated_by": d.get("updated_by", ""),
                "memo": d.get("last_memo", ""),
            }
            d.setdefault("history", [])
            if isinstance(d["history"], list):
                d["history"].append(prev_entry)

            d["current_file"] = dest_fn
            d["current_rev"] = new_rev
            d["updated_at"] = now_iso()
            d["updated_by"] = user_name()
            d["last_memo"] = memo

            docs[doc_key] = d
            meta["documents"] = docs
            save_meta(folder_path, meta)

            # Refresh
            self.refresh_files_table()
            self.refresh_folder_table()

        except Exception as e:
            self.warn(f"差し替えに失敗しました: {e}")

    def start_file_lock_watch(self, file_path: str, folder_path: str, doc_key: str):
        if self.watch_timer is None:
            self.watch_timer = QTimer(self)
            self.watch_timer.setInterval(5000)
            self.watch_timer.timeout.connect(self.on_watch_timer)
        self.watch_timer.stop()
        self.watch_target_path = file_path
        self.watch_folder_path = folder_path
        self.watch_doc_key = doc_key
        self.watch_started_at = dt.datetime.now()
        self.watch_timer.start()

    def on_watch_timer(self):
        if not self.watch_target_path:
            self.watch_timer.stop()
            return
        if not os.path.exists(self.watch_target_path):
            self.watch_timer.stop()
            return
        if self.watch_started_at is None:
            self.watch_started_at = dt.datetime.now()
        elapsed = (dt.datetime.now() - self.watch_started_at).total_seconds()
        if elapsed >= self.memo_timeout_min * 60:
            self.watch_timer.stop()
            self.prompt_memo_timeout()
            return
        if not is_file_locked(self.watch_target_path):
            self.watch_timer.stop()
            self.prompt_memo_input()

    def prompt_memo_timeout(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("メモ入力")
        msg.setText("一定時間が経過しました。作業メモを入力しますか？")
        btn_input = msg.addButton("入力する", QMessageBox.AcceptRole)
        btn_later = msg.addButton("後で", QMessageBox.RejectRole)
        msg.exec()
        if msg.clickedButton() == btn_input:
            self.prompt_memo_input()
        elif msg.clickedButton() == btn_later:
            self.info("後でを選択しました。作業メモは手動入力となります。")

    def prompt_memo_input(self):
        if not self.watch_folder_path or not self.watch_doc_key:
            return
        subtitle = f"対象ファイル: {os.path.basename(self.watch_target_path)}"
        dlg = MemoDialog("作業メモ入力", subtitle, "", self)
        if dlg.exec() != QDialog.Accepted:
            return
        memo = dlg.get_memo()
        if not memo:
            return
        meta = load_meta(self.watch_folder_path)
        docs = meta.get("documents", {})
        doc = docs.get(self.watch_doc_key)
        if isinstance(doc, dict):
            doc["last_memo"] = memo
            docs[self.watch_doc_key] = doc
            meta["documents"] = docs
            save_meta(self.watch_folder_path, meta)
            self.refresh_files_table()
            self.refresh_folder_table()
            self.refresh_right_pane_for_doc(self.watch_doc_key)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
