# library_manager_mock.py
# PySide6 mock for folder-based document versioning on OneDrive-synced folders.
# - Register folders (G1)
# - Show registered folders list (G0): name, last edited date
# - Show files in selected folder
# - Update (create next rev, move old to _History, write metadata + memo)
# - Replace (take external file, store as next rev, move old to _History, write metadata + memo)
#
# Notes:
# - Keeps your existing folder structure: "latest files" live in the registered folder root.
# - Uses "_History" folder under the registered folder.
# - Stores metadata as a hidden ".libra_meta.json" file under the registered folder.
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

from PySide6.QtCore import Qt, QSize, QTimer, Signal
from PySide6.QtGui import QAction, QBrush, QColor, QIcon, QPalette
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
    QRadioButton,
    QCheckBox,
)

REV_RE = re.compile(
    r"^(?P<base>.+)_rev(?P<A>\d+)\.(?P<B>\d+)\.(?P<C>\d+)_(?P<date>\d{8})$",
    re.IGNORECASE,
)

TEMP_FILE_RE = re.compile(r"^~\$")  # Office temporary files
META_FILENAME = ".libra_meta.json"
LEGACY_META_DIR = "_Meta"
LEGACY_META_FILENAME = "docmeta.json"


def now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def today_yyyymmdd() -> str:
    return dt.date.today().strftime("%Y%m%d")


def user_name() -> str:
    return getpass.getuser()


def appdata_dir() -> str:
    base = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "Libra")
    os.makedirs(path, exist_ok=True)
    return path


REGISTRY_PATH = os.path.join(appdata_dir(), "registry.json")
SETTINGS_PATH = os.path.join(appdata_dir(), "settings.json")
USER_CHECKS_PATH = os.path.join(appdata_dir(), "user_checks.json")
DEFAULT_MEMO_TIMEOUT_MIN = 30
UNCHECKED_COLOR = QColor("#C0504D")
NEW_FOLDER_BG_COLOR_LIGHT = QColor("#FFF2CC")
NEW_FOLDER_BG_COLOR_DARK = QColor("#4C3B00")
CATEGORY_PATH_SEP = "\u001f"
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
ASSET_DIR = os.path.join(ROOT_DIR, "assets")
APP_ICON_PATH = os.path.join(ASSET_DIR, "icons", "Libra.ico")
DEFAULT_VERSION_RULES = {
    "major": "",
    "minor": "",
    "patch": "",
}


def load_app_icon() -> Optional[QIcon]:
    if os.path.exists(APP_ICON_PATH):
        return QIcon(APP_ICON_PATH)
    return None


def normalize_category_path(values: List[str]) -> List[str]:
    return [value.strip() for value in values if value.strip()]


def categories_from_item(item: Dict[str, Any]) -> List[str]:
    raw_categories = item.get("categories")
    if isinstance(raw_categories, list):
        return normalize_category_path([str(x) for x in raw_categories])
    legacy = [
        str(item.get("main_category", "")),
        str(item.get("sub_category", "")),
    ]
    return normalize_category_path(legacy)


def load_registry() -> List[Dict[str, Any]]:
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
                    categories = categories_from_item(x)
                    out.append({
                        "name": str(x["name"]),
                        "path": str(x["path"]),
                        "categories": categories,
                    })
            return out
        return []
    except Exception:
        return []


def save_registry(items: List[Dict[str, Any]]) -> None:
    with open(REGISTRY_PATH, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)


def load_settings() -> Dict[str, Any]:
    defaults = {
        "memo_timeout_min": DEFAULT_MEMO_TIMEOUT_MIN,
        "category_order": {"categories": {}, "folder": {}},
        "archived_categories": [],
        "category_folder_paths": {},
        "folder_subfolder_counts": {},
        "ignore_types": {
            "shortcut": True,
            "bak": True,
            "log": True,
            "dwl": True,
            "dwl2": True,
            "ini": True,
        },
        "version_rules": DEFAULT_VERSION_RULES,
    }
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


def load_user_checks() -> Dict[str, Dict[str, bool]]:
    if not os.path.exists(USER_CHECKS_PATH):
        return {}
    try:
        with open(USER_CHECKS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("folders"), dict):
            return data["folders"]
        if isinstance(data, dict):
            return {k: v for k, v in data.items() if isinstance(v, dict)}
    except Exception:
        return {}
    return {}


def save_user_checks(checks: Dict[str, Dict[str, bool]]) -> None:
    with open(USER_CHECKS_PATH, "w", encoding="utf-8") as f:
        json.dump({"folders": checks}, f, ensure_ascii=False, indent=2)


def ensure_history_dir(folder_path: str) -> str:
    history_dir = os.path.join(folder_path, "_History")
    os.makedirs(history_dir, exist_ok=True)
    return history_dir


def legacy_meta_path_for_folder(folder_path: str) -> str:
    return os.path.join(folder_path, LEGACY_META_DIR, LEGACY_META_FILENAME)


def meta_path_for_folder(folder_path: str) -> str:
    return os.path.join(folder_path, META_FILENAME)


def set_hidden_on_windows(path: str) -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        FILE_ATTRIBUTE_HIDDEN = 0x02
        attrs = ctypes.windll.kernel32.GetFileAttributesW(path)
        if attrs == -1:
            return
        if attrs & FILE_ATTRIBUTE_HIDDEN:
            return
        ctypes.windll.kernel32.SetFileAttributesW(path, attrs | FILE_ATTRIBUTE_HIDDEN)
    except Exception:
        return


def load_meta(folder_path: str) -> Dict[str, Any]:
    p = meta_path_for_folder(folder_path)
    legacy_path = legacy_meta_path_for_folder(folder_path)
    if not os.path.exists(p):
        if not os.path.exists(legacy_path):
            return {"documents": {}}
        p = legacy_path
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
    set_hidden_on_windows(p)


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


def display_rev(rev: str) -> str:
    if not rev:
        return ""
    return rev.split("_", 1)[0]


def format_version_numbers(A: int, B: int, C: int) -> str:
    return f"{A}.{B}.{C}"


def parse_rev_numbers(current_rev: str) -> Optional[Tuple[int, int, int]]:
    if not current_rev:
        return None
    m = re.match(r"^rev(\d+)\.(\d+)\.(\d+)_\d{8}$", current_rev, re.IGNORECASE)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2)), int(m.group(3))


def next_rev_from_current(current_rev: str) -> Tuple[int, int, int]:
    """
    current_rev like 'rev1.2.3_20251223' or ''
    Returns (A,B,C) for the next revision, defaulting to (0,0,1) if none.
    Behavior: keep A,B; increment C by 1. If no current_rev, start A=0,B=0,C=1.
    """
    if not current_rev:
        return 0, 0, 1
    m = re.match(r"^rev(\d+)\.(\d+)\.(\d+)_\d{8}$", current_rev, re.IGNORECASE)
    if not m:
        return 0, 0, 1
    A, B, C = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return A, B, C + 1


def next_rev_with_bump(current_rev: str, bump: str) -> Tuple[int, int, int]:
    """
    bump: "major" | "minor" | "patch"
    """
    parsed = parse_rev_numbers(current_rev)
    if parsed is None:
        A, B, C = 0, 0, 0
    else:
        A, B, C = parsed
    if bump == "major":
        return A + 1, 0, 0
    if bump == "minor":
        return A, B + 1, 0
    return A, B, C + 1


def parse_rev_from_rev_string(rev_str: str) -> Optional[Tuple[int, int, int]]:
    m = re.match(r"^rev(\d+)\.(\d+)\.(\d+)_\d{8}$", rev_str, re.IGNORECASE)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2)), int(m.group(3))


def should_select_patch(latest: Tuple[int, int, int], candidate: Tuple[int, int, int]) -> bool:
    latest_A, latest_B, _latest_C = latest
    A, B, _C = candidate
    if (A, B) != (latest_A, latest_B):
        return False
    return candidate < latest


def should_select_minor(latest: Tuple[int, int, int], candidate: Tuple[int, int, int]) -> bool:
    latest_A, latest_B, _latest_C = latest
    A, B, _C = candidate
    if A != latest_A or B >= latest_B:
        return False
    return candidate < latest


def normalize_ignore_types(ignore_types: Optional[Dict[str, Any]]) -> Dict[str, bool]:
    defaults = {
        "shortcut": True,
        "bak": True,
        "log": True,
        "dwl": True,
        "dwl2": True,
        "ini": True,
    }
    if not isinstance(ignore_types, dict):
        return defaults
    merged = defaults.copy()
    for key in defaults.keys():
        if key in ignore_types:
            merged[key] = bool(ignore_types[key])
    return merged


def normalize_version_rules(version_rules: Optional[Dict[str, Any]]) -> Dict[str, str]:
    if not isinstance(version_rules, dict):
        return DEFAULT_VERSION_RULES.copy()
    normalized = DEFAULT_VERSION_RULES.copy()
    for key in normalized:
        if key in version_rules:
            normalized[key] = str(version_rules[key]).strip()
    return normalized


def should_ignore_file(filename: str, ignore_types: Dict[str, bool]) -> bool:
    _, ext = os.path.splitext(filename)
    ext = ext.lower()
    if ignore_types.get("shortcut") and ext == ".lnk":
        return True
    if ignore_types.get("bak") and ext == ".bak":
        return True
    if ignore_types.get("log") and ext == ".log":
        return True
    if ignore_types.get("dwl") and ext == ".dwl":
        return True
    if ignore_types.get("dwl2") and ext == ".dwl2":
        return True
    if ignore_types.get("ini") and ext == ".ini":
        return True
    return False


def safe_list_files(folder_path: str, ignore_types: Optional[Dict[str, Any]] = None) -> List[str]:
    ignore_flags = normalize_ignore_types(ignore_types)
    files = []
    try:
        for name in os.listdir(folder_path):
            full = os.path.join(folder_path, name)
            if os.path.isdir(full):
                # skip management folders
                if name.lower() in {"_history", LEGACY_META_DIR.lower()}:
                    continue
                continue
            if name.lower() == META_FILENAME.lower():
                continue
            if TEMP_FILE_RE.match(name):
                continue
            if should_ignore_file(name, ignore_flags):
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


class CategoryTreeWidget(QTreeWidget):
    order_changed = Signal()

    def dropEvent(self, event):  # type: ignore[override]
        items = self.selectedItems()
        if not items:
            return super().dropEvent(event)
        dragged_parent = items[0].parent()
        pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
        target_item = self.itemAt(pos)
        if target_item is None:
            if dragged_parent is not None:
                event.ignore()
                return
        else:
            if target_item.parent() is not dragged_parent:
                event.ignore()
                return
        super().dropEvent(event)
        self.order_changed.emit()


def scan_folder(
    folder_path: str,
    ignore_types: Optional[Dict[str, Any]] = None,
) -> Tuple[Dict[str, Any], List[FileRow]]:
    """
    Reads meta if exists and also scans real files.
    Returns (meta, file_rows_for_view).
    """
    meta = load_meta(folder_path)
    docs: Dict[str, Any] = meta.get("documents", {})
    files = safe_list_files(folder_path, ignore_types)

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

    # Remove docs whose current file no longer exists and are not present in scan
    stale_keys = []
    for doc_key, info in docs.items():
        cur_fn = info.get("current_file", "") if isinstance(info, dict) else ""
        if not cur_fn:
            stale_keys.append(doc_key)
            continue
        if cur_fn not in files and doc_key not in latest_by_doc:
            stale_keys.append(doc_key)
    if stale_keys:
        for key in stale_keys:
            docs.pop(key, None)
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
        initial_categories: Optional[List[str]] = None,
        category_options: Optional[List[List[str]]] = None,
        ok_label: str = "登録",
        subfolder_count: Optional[int] = None,
    ):
        super().__init__(parent)
        self.setWindowTitle("登録")
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
        self.category_edits: List[QComboBox] = []
        self.category_form = QFormLayout()
        self.category_form.setContentsMargins(0, 0, 0, 0)
        self.category_options = category_options or []
        self.initial_categories = initial_categories or []
        self.add_category_button = QPushButton("+")
        self.add_category_button.setToolTip("下の階層のカテゴリを追加")
        self.add_category_button.clicked.connect(self.add_category_row)

        form.addRow("登録フォルダパス", path_row)
        form.addRow("登録名", self.name_edit)
        category_box = QVBoxLayout()
        category_box.addLayout(self.category_form)
        button_row = QHBoxLayout()
        button_row.addStretch(1)
        button_row.addWidget(self.add_category_button)
        category_box.addLayout(button_row)
        form.addRow("カテゴリ階層", category_box)
        if subfolder_count is not None:
            count_label = QLabel(str(subfolder_count))
            form.addRow("配下フォルダ数", count_label)

        if initial_path:
            self.path_edit.setText(initial_path)
        if initial_name:
            self.name_edit.setText(initial_name)
        initial_levels = max(2, len(self.initial_categories) or 0)
        for _ in range(initial_levels):
            self.add_category_row()
        for idx, value in enumerate(self.initial_categories):
            if idx < len(self.category_edits):
                self.category_edits[idx].setCurrentText(value)

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

    def add_category_row(self):
        level = len(self.category_edits) + 1
        combo = QComboBox()
        combo.setEditable(True)
        placeholder = f"例：カテゴリ{level}　記入 or 選択"
        if combo.lineEdit():
            combo.lineEdit().setPlaceholderText(placeholder)
        if level - 1 < len(self.category_options):
            combo.addItems(self.category_options[level - 1])
        combo.setCurrentIndex(-1)
        self.category_edits.append(combo)
        self.category_form.addRow(f"カテゴリ階層{level}", combo)

    def get_values(self) -> Tuple[str, str, List[str]]:
        categories = []
        for edit in self.category_edits:
            value = edit.currentText().strip()
            if not value:
                break
            categories.append(value)
        return (
            self.name_edit.text().strip(),
            self.path_edit.text().strip(),
            categories,
        )


class CategoryEditDialog(QDialog):
    def __init__(
        self,
        parent: QWidget | None = None,
        initial_name: str = "",
        initial_path: str = "",
        subfolder_count: Optional[int] = None,
    ):
        super().__init__(parent)
        self.setWindowTitle("カテゴリ編集")
        self.setMinimumWidth(480)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.name_edit = QLineEdit()
        self.name_edit.setText(initial_name)

        self.path_edit = QLineEdit()
        self.path_edit.setText(initial_path)
        self.path_btn = QPushButton("参照...")
        self.path_btn.clicked.connect(self.pick_folder)

        path_row = QHBoxLayout()
        path_row.addWidget(self.path_edit, 1)
        path_row.addWidget(self.path_btn)

        form.addRow("カテゴリ名", self.name_edit)
        form.addRow("登録フォルダパス", path_row)
        if subfolder_count is not None:
            count_label = QLabel(str(subfolder_count))
            form.addRow("配下フォルダ数", count_label)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("更新")
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def pick_folder(self):
        path = QFileDialog.getExistingDirectory(self, "フォルダを選択")
        if path:
            self.path_edit.setText(path)

    def get_values(self) -> Tuple[str, str]:
        return self.name_edit.text().strip(), self.path_edit.text().strip()


class BatchRegisterDialog(QDialog):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("一括登録")
        self.setMinimumWidth(520)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.path_edit = QLineEdit()
        self.path_btn = QPushButton("参照...")
        self.path_btn.clicked.connect(self.pick_folder)

        path_row = QHBoxLayout()
        path_row.addWidget(self.path_edit, 1)
        path_row.addWidget(self.path_btn)

        self.depth_spin = QSpinBox()
        self.depth_spin.setRange(0, 10)
        self.depth_spin.setValue(2)

        form.addRow("対象ディレクトリ", path_row)
        form.addRow("取得階層", self.depth_spin)
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("OK")
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def pick_folder(self):
        path = QFileDialog.getExistingDirectory(self, "登録するディレクトリを選択")
        if path:
            self.path_edit.setText(path)

    def get_path(self) -> str:
        return self.path_edit.text().strip()

    def get_depth(self) -> int:
        return int(self.depth_spin.value())


class BatchPreviewDialog(QDialog):
    def __init__(self, items: List[Dict[str, Any]], parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("一括登録プレビュー")
        self.setMinimumWidth(720)
        self.setMinimumHeight(460)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("以下の内容で取り込みます。よろしいですか？"))
        self.select_all_checkbox = QCheckBox("全選択")
        self.select_all_checkbox.setChecked(True)
        self.select_all_checkbox.toggled.connect(self.on_select_all_toggled)
        layout.addWidget(self.select_all_checkbox)

        self.tree = QTreeWidget()
        self.tree.setColumnCount(3)
        self.tree.setHeaderLabels(["取込", "登録名", "フォルダパス"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tree.header().setSectionResizeMode(2, QHeaderView.Stretch)
        self.tree.setSelectionMode(QAbstractItemView.SingleSelection)
        layout.addWidget(self.tree, 1)

        nodes: Dict[Tuple[str, ...], QTreeWidgetItem] = {}
        for entry in items:
            parts = entry.get("rel_parts", [])
            if not isinstance(parts, list):
                continue
            parent_item: Optional[QTreeWidgetItem] = None
            for depth, part in enumerate(parts):
                key = tuple(parts[: depth + 1])
                item = nodes.get(key)
                if not item:
                    item = QTreeWidgetItem([ "", part, "" ])
                    item.setFlags(item.flags() | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                    item.setCheckState(0, Qt.Checked)
                    if parent_item is None:
                        self.tree.addTopLevelItem(item)
                    else:
                        parent_item.addChild(item)
                    nodes[key] = item
                parent_item = item

            if parent_item is None:
                parent_item = QTreeWidgetItem([ "", entry.get("name", ""), entry.get("path", "") ])
                self.tree.addTopLevelItem(parent_item)
            parent_item.setText(1, entry.get("name", ""))
            parent_item.setText(2, entry.get("path", ""))
            parent_item.setData(0, Qt.UserRole, entry)
            parent_item.setFlags(parent_item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEditable)
            parent_item.setCheckState(0, Qt.Checked)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("登録")
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def on_select_all_toggled(self, checked: bool) -> None:
        target_state = Qt.Checked if checked else Qt.Unchecked
        def walk(item: QTreeWidgetItem):
            item.setCheckState(0, target_state)
            for i in range(item.childCount()):
                walk(item.child(i))

        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i))

    def selected_items(self) -> List[Dict[str, Any]]:
        selected: List[Dict[str, Any]] = []
        def walk(item: QTreeWidgetItem, ancestor_checked: bool):
            current_checked = ancestor_checked and item.checkState(0) == Qt.Checked
            entry = item.data(0, Qt.UserRole)
            if entry and isinstance(entry, dict):
                if current_checked:
                    name = item.text(1).strip()
                    path = entry.get("path")
                    categories = entry.get("categories")
                    if name and path and categories:
                        selected.append({
                            "name": name,
                            "path": path,
                            "categories": categories,
                        })
            for i in range(item.childCount()):
                walk(item.child(i), current_checked)

        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i), True)
        return selected

    def unchecked_items(self) -> List[Dict[str, Any]]:
        skipped: List[Dict[str, Any]] = []

        def walk(item: QTreeWidgetItem, ancestor_checked: bool):
            current_checked = ancestor_checked and item.checkState(0) == Qt.Checked
            entry = item.data(0, Qt.UserRole)
            if entry and isinstance(entry, dict) and not current_checked:
                path = entry.get("path")
                if path:
                    skipped.append({"path": path})
            for i in range(item.childCount()):
                walk(item.child(i), current_checked)

        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i), True)
        return skipped


class ArchiveDialog(QDialog):
    def __init__(self, entries: List[Dict[str, Any]], parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("アーカイブ")
        self.setMinimumWidth(520)
        self.setMinimumHeight(360)
        self.selected_action: Optional[str] = None

        layout = QVBoxLayout(self)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["カテゴリ階層", "件数"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        layout.addWidget(self.table, 1)

        for entry in entries:
            row = self.table.rowCount()
            self.table.insertRow(row)
            path_text = entry.get("label", "")
            count = entry.get("count", 0)
            it_path = QTableWidgetItem(path_text)
            it_path.setData(Qt.UserRole, entry.get("path", []))
            self.table.setItem(row, 0, it_path)
            self.table.setItem(row, 1, QTableWidgetItem(str(count)))

        btn_row = QHBoxLayout()
        self.btn_restore = QPushButton("復帰")
        self.btn_delete = QPushButton("削除")
        self.btn_cancel = QPushButton("キャンセル")
        btn_row.addWidget(self.btn_restore)
        btn_row.addWidget(self.btn_delete)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_cancel)
        layout.addLayout(btn_row)

        self.btn_restore.clicked.connect(self.on_restore)
        self.btn_delete.clicked.connect(self.on_delete)
        self.btn_cancel.clicked.connect(self.reject)

    def selected_path(self) -> List[str]:
        row = self.table.currentRow()
        if row < 0:
            return []
        item = self.table.item(row, 0)
        if not item:
            return []
        return item.data(Qt.UserRole) or []

    def ensure_selection(self) -> bool:
        if self.table.currentRow() < 0:
            QMessageBox.warning(self, "注意", "対象を選択してください。")
            return False
        return True

    def on_restore(self):
        if not self.ensure_selection():
            return
        self.selected_action = "restore"
        self.accept()

    def on_delete(self):
        if not self.ensure_selection():
            return
        self.selected_action = "delete"
        self.accept()


class VersionSelectDialog(QDialog):
    def __init__(
        self,
        current_rev: str,
        version_rules: Optional[Dict[str, Any]] = None,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.setWindowTitle("更新")
        self.setMinimumWidth(420)

        self.current_rev = current_rev
        self.current_version = parse_rev_numbers(current_rev)
        self.version_rules = normalize_version_rules(version_rules)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        current_text = "未設定"
        if self.current_version:
            current_text = format_version_numbers(*self.current_version)
        self.current_label = QLabel(current_text)

        self.next_label = QLabel("")

        form.addRow("ファイルのバージョン", self.current_label)
        form.addRow("作られるファイルのバージョン", self.next_label)
        layout.addLayout(form)

        self.bump_group = QGroupBox("バージョンの選択")
        bump_layout = QVBoxLayout(self.bump_group)
        self.radio_major = QRadioButton(f"メジャーバージョン：{self.version_rules['major']}")
        self.radio_minor = QRadioButton(f"マイナーバージョン：{self.version_rules['minor']}")
        self.radio_patch = QRadioButton(f"パッチバージョン：{self.version_rules['patch']}")
        self.radio_patch.setChecked(True)
        bump_layout.addWidget(self.radio_major)
        bump_layout.addWidget(self.radio_minor)
        bump_layout.addWidget(self.radio_patch)
        layout.addWidget(self.bump_group)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.button(QDialogButtonBox.Ok).setText("OK")
        self.buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

        self.radio_major.toggled.connect(self.update_next_version)
        self.radio_minor.toggled.connect(self.update_next_version)
        self.radio_patch.toggled.connect(self.update_next_version)
        self.update_next_version()

    def selected_bump(self) -> str:
        if self.radio_major.isChecked():
            return "major"
        if self.radio_minor.isChecked():
            return "minor"
        return "patch"

    def update_next_version(self):
        bump = self.selected_bump()
        next_rev = next_rev_with_bump(self.current_rev, bump)
        self.next_label.setText(format_version_numbers(*next_rev))

    def selected_next_version(self) -> Tuple[int, int, int]:
        return next_rev_with_bump(self.current_rev, self.selected_bump())


class HistoryClearDialog(QDialog):
    def __init__(self, history_items: List[Dict[str, Any]], current_rev: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("History Clear")
        self.setMinimumWidth(560)
        self.history_items = history_items
        self.current_rev = current_rev

        layout = QVBoxLayout(self)

        current_text = "未設定"
        current_version = parse_rev_numbers(current_rev)
        if current_version:
            current_text = format_version_numbers(*current_version)
        layout.addWidget(QLabel(f"現在のバージョン: {current_text}"))

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["選択", "rev", "ファイル"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        layout.addWidget(self.table)

        btn_row = QHBoxLayout()
        self.btn_patch = QPushButton("パッチ削除")
        self.btn_minor = QPushButton("マイナー削除")
        self.btn_delete = QPushButton("削除")
        self.btn_close = QPushButton("閉じる")
        btn_row.addWidget(self.btn_patch)
        btn_row.addWidget(self.btn_minor)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_delete)
        btn_row.addWidget(self.btn_close)
        layout.addLayout(btn_row)

        self.btn_patch.clicked.connect(self.select_patch_versions)
        self.btn_minor.clicked.connect(self.select_minor_versions)
        self.btn_delete.clicked.connect(self.accept)
        self.btn_close.clicked.connect(self.reject)

        self.populate_table()

    def populate_table(self):
        self.table.setRowCount(0)
        seen = set()
        for item in self.history_items:
            key = (item.get("file", ""), item.get("rev", ""))
            if key in seen:
                continue
            seen.add(key)
            r = self.table.rowCount()
            self.table.insertRow(r)
            check_item = QTableWidgetItem("")
            check_item.setCheckState(Qt.Unchecked)
            check_item.setData(Qt.UserRole, item)
            rev_item = QTableWidgetItem(item.get("rev", ""))
            file_item = QTableWidgetItem(item.get("file", ""))
            self.table.setItem(r, 0, check_item)
            self.table.setItem(r, 1, rev_item)
            self.table.setItem(r, 2, file_item)

    def selected_items(self) -> List[Dict[str, Any]]:
        selected = []
        for r in range(self.table.rowCount()):
            check_item = self.table.item(r, 0)
            if check_item and check_item.checkState() == Qt.Checked:
                data = check_item.data(Qt.UserRole)
                if isinstance(data, dict):
                    selected.append(data)
        return selected

    def latest_version_tuple(self) -> Optional[Tuple[int, int, int]]:
        return parse_rev_numbers(self.current_rev)

    def select_patch_versions(self):
        latest = self.latest_version_tuple()
        if not latest:
            return
        for r in range(self.table.rowCount()):
            check_item = self.table.item(r, 0)
            if not check_item:
                continue
            data = check_item.data(Qt.UserRole)
            if not isinstance(data, dict):
                continue
            rev = data.get("rev", "")
            candidate = parse_rev_from_rev_string(rev)
            if not candidate:
                check_item.setCheckState(Qt.Unchecked)
                continue
            if should_select_patch(latest, candidate):
                check_item.setCheckState(Qt.Checked)
            else:
                check_item.setCheckState(Qt.Unchecked)

    def select_minor_versions(self):
        latest = self.latest_version_tuple()
        if not latest:
            return
        for r in range(self.table.rowCount()):
            check_item = self.table.item(r, 0)
            if not check_item:
                continue
            data = check_item.data(Qt.UserRole)
            if not isinstance(data, dict):
                continue
            rev = data.get("rev", "")
            candidate = parse_rev_from_rev_string(rev)
            if not candidate:
                check_item.setCheckState(Qt.Unchecked)
                continue
            if should_select_minor(latest, candidate):
                check_item.setCheckState(Qt.Checked)
            else:
                check_item.setCheckState(Qt.Unchecked)


class HistorySelectDialog(QDialog):
    def __init__(self, history_items: List[Dict[str, Any]], parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("差し戻し")
        self.setMinimumWidth(520)
        self.history_items = history_items

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("差し戻す履歴を選択してください。"))

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["rev", "ファイル"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        layout.addWidget(self.table)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("差し戻し")
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.populate_table()

    def populate_table(self):
        self.table.setRowCount(0)
        seen = set()
        for item in self.history_items:
            key = (item.get("file", ""), item.get("rev", ""))
            if key in seen:
                continue
            seen.add(key)
            r = self.table.rowCount()
            self.table.insertRow(r)
            rev_item = QTableWidgetItem(item.get("rev", ""))
            rev_item.setData(Qt.UserRole, item)
            file_item = QTableWidgetItem(item.get("file", ""))
            self.table.setItem(r, 0, rev_item)
            self.table.setItem(r, 1, file_item)

    def selected_item(self) -> Optional[Dict[str, Any]]:
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            return None
        row = sel[0].row()
        item = self.table.item(row, 0)
        if not item:
            return None
        data = item.data(Qt.UserRole)
        if isinstance(data, dict):
            return data
        return None


class HistoryDetailDialog(QDialog):
    def __init__(self, entry: Dict[str, Any], parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("履歴詳細")
        self.setMinimumWidth(560)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        kind_text = entry.get("kind", "") or ""
        rev_text = display_rev(entry.get("rev", "") or "")
        updated_text = (entry.get("updated_at", "") or "").replace("T", " ")
        updated_by = entry.get("updated_by", "") or ""
        file_name = entry.get("file", "") or ""

        form.addRow("区分", QLabel(kind_text))
        form.addRow("rev", QLabel(rev_text))
        form.addRow("更新日時", QLabel(updated_text))
        form.addRow("更新者", QLabel(updated_by))
        form.addRow("ファイル", QLabel(file_name))
        layout.addLayout(form)

        memo_label = QLabel("メモ")
        memo_view = QPlainTextEdit()
        memo_view.setReadOnly(True)
        memo_view.setPlainText(entry.get("memo", "") or "")
        layout.addWidget(memo_label)
        layout.addWidget(memo_view)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.button(QDialogButtonBox.Ok).setText("閉じる")
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


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


class DropLineEdit(QLineEdit):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):  # noqa: N802
        urls = event.mimeData().urls()
        if not urls:
            event.ignore()
            return
        path = urls[0].toLocalFile()
        if not path:
            event.ignore()
            return
        self.setText(path)
        event.acceptProposedAction()


class ReplaceDialog(QDialog):
    def __init__(self, target_filename: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("差し替え")
        self.setMinimumWidth(640)
        self.setAcceptDrops(True)

        self.target_filename = target_filename

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"差し替え対象（現行）: {target_filename}"))

        self.file_edit = DropLineEdit()
        self.file_btn = QPushButton("ファイル選択...")
        self.file_btn.clicked.connect(self.pick_file)

        row = QHBoxLayout()
        row.addWidget(self.file_edit, 1)
        row.addWidget(self.file_btn)
        layout.addLayout(row)

        self.memo = QPlainTextEdit()
        self.memo.setAcceptDrops(False)
        self.memo.setPlaceholderText("作業メモ（空欄可） 例：外注成果品差し替え、図面反映 等")
        layout.addWidget(self.memo, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("差し替え実行")
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def dragEnterEvent(self, event):  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):  # noqa: N802
        urls = event.mimeData().urls()
        if not urls:
            event.ignore()
            return
        path = urls[0].toLocalFile()
        if not path:
            event.ignore()
            return
        self.file_edit.setText(path)
        event.acceptProposedAction()

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "差し替えるファイルを選択")
        if path:
            self.file_edit.setText(path)

    def get_values(self) -> Tuple[str, str]:
        return self.file_edit.text().strip(), self.memo.toPlainText().strip()


class OptionsDialog(QDialog):
    def __init__(
        self,
        timeout_min: int,
        ignore_types: Optional[Dict[str, Any]] = None,
        version_rules: Optional[Dict[str, Any]] = None,
        parent: QWidget | None = None,
    ):
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

        rules = normalize_version_rules(version_rules)
        rules_group = QGroupBox("バージョンルール")
        rules_form = QFormLayout(rules_group)
        self.rule_major = QLineEdit(rules["major"])
        self.rule_minor = QLineEdit(rules["minor"])
        self.rule_patch = QLineEdit(rules["patch"])
        rules_form.addRow("メジャーバージョン", self.rule_major)
        rules_form.addRow("マイナーバージョン", self.rule_minor)
        rules_form.addRow("パッチバージョン", self.rule_patch)
        layout.addWidget(rules_group)

        ignore_flags = normalize_ignore_types(ignore_types)
        ignore_group = QGroupBox("無視するファイル種類")
        ignore_layout = QVBoxLayout(ignore_group)
        self.ignore_shortcut = QCheckBox("ショートカット（.lnk）")
        self.ignore_shortcut.setChecked(ignore_flags.get("shortcut", False))
        ignore_layout.addWidget(self.ignore_shortcut)
        self.ignore_bak = QCheckBox(".bak")
        self.ignore_bak.setChecked(ignore_flags.get("bak", False))
        ignore_layout.addWidget(self.ignore_bak)
        self.ignore_log = QCheckBox(".log")
        self.ignore_log.setChecked(ignore_flags.get("log", False))
        ignore_layout.addWidget(self.ignore_log)
        self.ignore_dwl = QCheckBox(".dwl")
        self.ignore_dwl.setChecked(ignore_flags.get("dwl", False))
        ignore_layout.addWidget(self.ignore_dwl)
        self.ignore_dwl2 = QCheckBox(".dwl2")
        self.ignore_dwl2.setChecked(ignore_flags.get("dwl2", False))
        ignore_layout.addWidget(self.ignore_dwl2)
        self.ignore_ini = QCheckBox(".ini")
        self.ignore_ini.setChecked(ignore_flags.get("ini", False))
        ignore_layout.addWidget(self.ignore_ini)
        layout.addWidget(ignore_group)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("保存")
        buttons.button(QDialogButtonBox.Cancel).setText("キャンセル")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_timeout_min(self) -> int:
        return int(self.timeout_spin.value())

    def get_ignore_types(self) -> Dict[str, bool]:
        return {
            "shortcut": self.ignore_shortcut.isChecked(),
            "bak": self.ignore_bak.isChecked(),
            "log": self.ignore_log.isChecked(),
            "dwl": self.ignore_dwl.isChecked(),
            "dwl2": self.ignore_dwl2.isChecked(),
            "ini": self.ignore_ini.isChecked(),
        }

    def get_version_rules(self) -> Dict[str, str]:
        return normalize_version_rules({
            "major": self.rule_major.text(),
            "minor": self.rule_minor.text(),
            "patch": self.rule_patch.text(),
        })


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Libra")
        self.resize(1300, 760)
        icon = load_app_icon()
        if icon:
            self.setWindowIcon(icon)

        self.registry = load_registry()
        self.settings = load_settings()
        self.user_checks = load_user_checks()
        self.memo_timeout_min = int(self.settings.get("memo_timeout_min", DEFAULT_MEMO_TIMEOUT_MIN))
        self.ignore_types = normalize_ignore_types(self.settings.get("ignore_types"))
        self.version_rules = normalize_version_rules(self.settings.get("version_rules"))

        root = QWidget()
        self.setCentralWidget(root)
        root_layout = QVBoxLayout(root)

        toolbar = QToolBar("ツール")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        act_options = QAction("オプション", self)
        act_options.triggered.connect(self.on_options)
        toolbar.addAction(act_options)
        act_archive = QAction("アーカイブ", self)
        act_archive.triggered.connect(self.on_archive)
        toolbar.addAction(act_archive)
        act_cache_clear = QAction("キャッシュクリア", self)
        act_cache_clear.triggered.connect(self.on_cache_clear)
        toolbar.addAction(act_cache_clear)

        # Top controls
        top = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("検索（登録名でフィルタ）")
        self.search.textChanged.connect(self.refresh_folder_table)

        btn_rescan = QPushButton("再スキャン")
        btn_rescan.clicked.connect(self.on_rescan)

        top.addWidget(self.search, 1)
        top.addWidget(btn_rescan)

        root_layout.addLayout(top)

        splitter = QSplitter(Qt.Horizontal)
        root_layout.addWidget(splitter, 1)

        # Left: category tree
        tree_box = QWidget()
        tree_layout = QVBoxLayout(tree_box)
        tree_layout.setContentsMargins(0, 0, 0, 0)

        tree_title = QLabel("カテゴリツリー")

        btn_register = QPushButton("個別登録")
        btn_register.clicked.connect(self.on_register)

        btn_batch_register = QPushButton("一括登録")
        btn_batch_register.clicked.connect(self.on_batch_register)

        tree_button_row = QHBoxLayout()
        tree_button_row.addWidget(btn_register)
        tree_button_row.addWidget(btn_batch_register)
        tree_button_row.addStretch(1)

        self.category_tree = CategoryTreeWidget()
        self.category_tree.setHeaderHidden(True)
        self.category_tree.setDragEnabled(True)
        self.category_tree.setAcceptDrops(True)
        self.category_tree.setDropIndicatorShown(True)
        self.category_tree.setDragDropMode(QAbstractItemView.InternalMove)
        self.category_tree.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.category_tree.order_changed.connect(self.on_category_tree_order_changed)
        self.category_tree.itemSelectionChanged.connect(self.on_category_tree_selected)
        self.category_tree.itemDoubleClicked.connect(self.on_category_tree_double_clicked)
        self.category_tree.itemChanged.connect(self.on_category_tree_item_changed)
        self.category_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.category_tree.customContextMenuRequested.connect(self.on_category_tree_context_menu)
        tree_layout.addWidget(tree_title)
        tree_layout.addLayout(tree_button_row)
        tree_layout.addWidget(self.category_tree)
        splitter.addWidget(tree_box)

        # Left-middle: registered folders table
        left_box = QWidget()
        left_layout = QVBoxLayout(left_box)
        left_layout.setContentsMargins(0, 0, 0, 0)

        folders_title = QLabel("フォルダリスト")

        btn_edit_folder = QPushButton("編集")
        btn_edit_folder.clicked.connect(self.on_edit_selected_folder)

        folders_button_row = QHBoxLayout()
        folders_button_row.addWidget(btn_edit_folder)
        folders_button_row.addStretch(1)

        self.folders_table = QTableWidget(0, 2)
        self.folders_table.setHorizontalHeaderLabels(["登録名", "最終更新日"])
        self.folders_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.folders_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.folders_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.folders_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.folders_table.itemSelectionChanged.connect(self.on_folder_selected)
        self.folders_table.itemDoubleClicked.connect(self.on_folder_double_clicked)
        self.folders_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.folders_table.customContextMenuRequested.connect(self.on_folders_table_context_menu)

        left_layout.addWidget(folders_title)
        left_layout.addLayout(folders_button_row)
        left_layout.addWidget(self.folders_table)
        splitter.addWidget(left_box)

        # Middle: files table + history
        mid_box = QWidget()
        mid_layout = QVBoxLayout(mid_box)
        mid_layout.setContentsMargins(0, 0, 0, 0)

        files_title = QLabel("ファイルリスト")

        btn_update = QPushButton("更新")
        btn_update.clicked.connect(self.on_update)

        btn_replace = QPushButton("差し替え")
        btn_replace.clicked.connect(self.on_replace)

        btn_view = QPushButton("閲覧")
        btn_view.clicked.connect(self.on_view)

        files_button_row = QHBoxLayout()
        files_button_row.addWidget(btn_update)
        files_button_row.addWidget(btn_replace)
        files_button_row.addWidget(btn_view)
        files_button_row.addStretch(1)

        self.files_table = QTableWidget(0, 6)
        self.files_table.setHorizontalHeaderLabels(["", "ファイル", "rev", "更新日", "更新者", "DocKey"])
        files_header_item = QTableWidgetItem("")
        files_header_item.setFlags(files_header_item.flags() | Qt.ItemIsUserCheckable)
        files_header_item.setCheckState(Qt.Unchecked)
        self.files_table.setHorizontalHeaderItem(0, files_header_item)
        self.files_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.files_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().sectionClicked.connect(self.on_files_header_clicked)
        self.files_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.files_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.files_table.itemSelectionChanged.connect(self.on_file_selected)
        self.files_table.itemDoubleClicked.connect(self.on_file_double_clicked)
        self.files_table.itemChanged.connect(self.on_file_check_changed)
        self.files_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.files_table.customContextMenuRequested.connect(self.on_files_table_context_menu)

        hist_group = QWidget()
        hist_layout = QVBoxLayout(hist_group)
        hist_layout.setContentsMargins(0, 0, 0, 0)
        hist_title = QLabel("履歴")
        self.hist_table = QTableWidget(0, 5)
        self.hist_table.setHorizontalHeaderLabels(["区分", "rev", "更新日時", "更新者", "メモ"])
        self.hist_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.hist_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.hist_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.hist_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.hist_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.hist_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.hist_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.hist_table.itemDoubleClicked.connect(self.on_history_item_double_clicked)
        hist_layout.addWidget(hist_title)
        hist_layout.addWidget(self.hist_table)

        mid_layout.addWidget(files_title)
        mid_layout.addLayout(files_button_row)
        mid_layout.addWidget(self.files_table, 2)
        mid_layout.addWidget(hist_group, 1)
        splitter.addWidget(mid_box)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 2)
        splitter.setSizes([300, 300, 600])

        # State
        self.current_folder: Optional[Dict[str, str]] = None
        self.current_meta: Optional[Dict[str, Any]] = None
        self.current_file_rows: List[FileRow] = []
        self.selected_category_path: List[str] = []
        self.watch_timer = None
        self.watch_target_path = ""
        self.watch_folder_path = ""
        self.watch_doc_key = ""
        self.watch_started_at: Optional[dt.datetime] = None
        self.current_user = user_name()
        self._category_tree_refreshing = False
        self._category_tree_refresh_pending = False
        self.new_folder_highlights: set[str] = set()
        self.new_category_highlights: set[str] = set()

        self.startup_rescan()
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

    def strip_icon_prefix(self, text: str) -> str:
        for prefix in ("🔖 ", "📁 "):
            if text.startswith(prefix):
                return text[len(prefix):]
        return text

    def category_path_key(self, path: List[str]) -> str:
        return CATEGORY_PATH_SEP.join(path)

    def category_path_from_key(self, key: str) -> List[str]:
        if not key:
            return []
        return [part for part in key.split(CATEGORY_PATH_SEP) if part]

    def category_path_for_item(self, item: Dict[str, Any]) -> List[str]:
        categories = categories_from_item(item)
        if categories:
            return categories
        return ["未分類"]

    def archived_categories(self) -> List[List[str]]:
        data = self.settings.get("archived_categories", [])
        archived: List[List[str]] = []
        if isinstance(data, list):
            for entry in data:
                if isinstance(entry, list):
                    path = normalize_category_path([str(x) for x in entry])
                elif isinstance(entry, str):
                    path = normalize_category_path(entry.split(CATEGORY_PATH_SEP))
                else:
                    continue
                if path:
                    archived.append(path)
        return archived

    def save_archived_categories(self, archived: List[List[str]]) -> None:
        self.settings["archived_categories"] = archived
        save_settings(self.settings)

    def is_archived_path(self, categories: List[str]) -> bool:
        if not categories:
            return False
        for archived in self.archived_categories():
            if categories[:len(archived)] == archived:
                return True
        return False

    def archive_category_path(self, path: List[str]) -> bool:
        if not path:
            return False
        archived = self.archived_categories()
        for entry in archived:
            if path[:len(entry)] == entry:
                return False
        archived = [entry for entry in archived if entry[:len(path)] != path]
        archived.append(path)
        self.save_archived_categories(archived)
        return True

    def unarchive_category_path(self, path: List[str]) -> None:
        archived = [entry for entry in self.archived_categories() if entry != path]
        self.save_archived_categories(archived)

    def remove_archived_under_path(self, path: List[str]) -> None:
        archived = [
            entry for entry in self.archived_categories()
            if entry[:len(path)] != path
        ]
        self.save_archived_categories(archived)

    def root_category_name(self, root_path: str) -> str:
        base = os.path.basename(os.path.normpath(root_path))
        if base:
            return base
        drive, _ = os.path.splitdrive(root_path)
        return drive or root_path

    def batch_register_items(
        self,
        root_path: str,
        max_depth: int,
        base_categories: Optional[List[str]] = None,
    ) -> List[Dict[str, Any]]:
        base_categories = base_categories or []
        root_category = self.root_category_name(root_path)
        include_root_category = not (base_categories and base_categories[-1] == root_category)
        registered_paths = {os.path.normcase(item["path"]) for item in self.registry}
        category_folder_paths = {
            os.path.normcase(folder_path)
            for folder_path in self.category_folder_paths().values()
            if folder_path
        }
        reserved_paths = registered_paths | category_folder_paths
        items: List[Dict[str, Any]] = []
        for dirpath, dirnames, _filenames in os.walk(root_path):
            dirnames[:] = [d for d in dirnames if d.lower() not in {"_history", LEGACY_META_DIR.lower()}]
            dirnames.sort(key=str.casefold)
            rel = os.path.relpath(dirpath, root_path)
            depth = 0 if rel == "." else len(rel.split(os.sep))
            if depth > max_depth:
                dirnames[:] = []
                continue
            if depth >= max_depth:
                dirnames[:] = []
            if dirnames and depth < max_depth:
                continue
            if os.path.normcase(dirpath) in reserved_paths:
                continue
            rel_parts = [] if rel == "." else rel.split(os.sep)
            base_parts = base_categories + [root_category] if include_root_category else list(base_categories)
            if rel_parts:
                folder_name = rel_parts[-1]
                category_parts = base_parts + rel_parts[:-1]
            else:
                folder_name = root_category
                category_parts = base_parts
            categories = normalize_category_path(category_parts)
            items.append({
                "name": folder_name,
                "path": dirpath,
                "categories": categories,
                "depth": depth,
                "rel_parts": rel_parts,
            })
        items.sort(key=lambda item: [part.casefold() for part in item.get("rel_parts", [])])
        return items

    def category_options(self) -> List[List[str]]:
        levels: Dict[int, set[str]] = {}
        for item in self.registry:
            if self.is_archived_path(self.category_path_for_item(item)):
                continue
            categories = self.category_path_for_item(item)
            for idx, name in enumerate(categories):
                levels.setdefault(idx, set()).add(name)
        if not levels:
            return []
        max_level = max(levels.keys())
        return [sorted(levels.get(i, set())) for i in range(max_level + 1)]

    def category_order(self) -> Dict[str, Any]:
        order = self.settings.get("category_order")
        if not isinstance(order, dict):
            order = {"categories": {}, "folder": {}, "tree": {}}
            self.settings["category_order"] = order
        if "main" in order or "sub" in order:
            order = self.migrate_category_order(order)
            self.settings["category_order"] = order
        order.setdefault("categories", {})
        order.setdefault("folder", {})
        order.setdefault("tree", {})
        return order

    def category_folder_paths(self) -> Dict[str, str]:
        paths = self.settings.get("category_folder_paths")
        if not isinstance(paths, dict):
            paths = {}
            self.settings["category_folder_paths"] = paths
        return {str(k): str(v) for k, v in paths.items() if isinstance(v, str)}

    def category_folder_path_for_path(self, path: List[str]) -> str:
        key = self.category_path_key(path)
        return self.category_folder_paths().get(key, "")

    def set_category_folder_path(self, path: List[str], folder_path: str) -> None:
        paths = self.category_folder_paths()
        key = self.category_path_key(path)
        if folder_path:
            paths[key] = folder_path
            self.record_folder_subfolder_count(folder_path)
        else:
            paths.pop(key, None)
        self.settings["category_folder_paths"] = paths
        save_settings(self.settings)

    def replace_category_prefix(self, path: List[str], old_prefix: List[str], new_prefix: List[str]) -> List[str]:
        if path[:len(old_prefix)] != old_prefix:
            return path
        return new_prefix + path[len(old_prefix):]

    def rename_category_in_settings(self, old_path: List[str], new_path: List[str]) -> None:
        old_key = self.category_path_key(old_path)
        new_key = self.category_path_key(new_path)
        old_name = old_path[-1] if old_path else ""
        new_name = new_path[-1] if new_path else ""

        def remap_key(key: str) -> str:
            if key == old_key:
                return new_key
            prefix = f"{old_key}{CATEGORY_PATH_SEP}"
            if key.startswith(prefix):
                return new_key + key[len(old_key):]
            return key

        order = self.category_order()
        for section in ("categories", "folder", "tree"):
            section_map = order.get(section, {})
            if not isinstance(section_map, dict):
                continue
            remapped = {}
            for key, value in section_map.items():
                remapped[remap_key(str(key))] = value
            order[section] = remapped
        parent_key = self.category_path_key(old_path[:-1])
        siblings = order.get("categories", {}).get(parent_key)
        if isinstance(siblings, list):
            order["categories"][parent_key] = [new_name if name == old_name else name for name in siblings]
        tree_siblings = order.get("tree", {}).get(parent_key)
        if isinstance(tree_siblings, list):
            for entry in tree_siblings:
                if isinstance(entry, dict) and entry.get("type") == "category" and entry.get("name") == old_name:
                    entry["name"] = new_name
        self.settings["category_order"] = order

        checks = self.category_check_states()
        remapped_checks = {}
        for key, value in checks.items():
            remapped_checks[remap_key(str(key))] = value
        self.settings["category_check_states"] = remapped_checks

        paths = self.category_folder_paths()
        remapped_paths = {}
        for key, value in paths.items():
            remapped_paths[remap_key(str(key))] = value
        self.settings["category_folder_paths"] = remapped_paths

        archived = [
            self.replace_category_prefix(path, old_path, new_path)
            for path in self.archived_categories()
        ]
        self.save_archived_categories(archived)
        save_settings(self.settings)

    def category_check_states(self) -> Dict[str, bool]:
        checks = self.settings.get("category_check_states")
        if not isinstance(checks, dict):
            checks = {}
            self.settings["category_check_states"] = checks
        return checks

    def folder_tree_check_states(self) -> Dict[str, bool]:
        checks = self.settings.get("folder_tree_check_states")
        if not isinstance(checks, dict):
            checks = {}
            self.settings["folder_tree_check_states"] = checks
        return checks

    def is_category_checked(self, path: List[str]) -> bool:
        key = self.category_path_key(path)
        return bool(self.category_check_states().get(key, True))

    def is_folder_tree_checked(self, folder_path: str) -> bool:
        key = self.folder_key(folder_path)
        return bool(self.folder_tree_check_states().get(key, True))

    def set_category_checked(self, path: List[str], checked: bool) -> None:
        checks = self.category_check_states()
        checks[self.category_path_key(path)] = checked
        self.settings["category_check_states"] = checks
        save_settings(self.settings)

    def set_folder_tree_checked(self, folder_path: str, checked: bool) -> None:
        checks = self.folder_tree_check_states()
        checks[self.folder_key(folder_path)] = checked
        self.settings["folder_tree_check_states"] = checks
        save_settings(self.settings)

    def is_category_highlight_enabled_for_path(self, path: List[str]) -> bool:
        for depth in range(1, len(path) + 1):
            if not self.is_category_checked(path[:depth]):
                return False
        return True

    def migrate_category_order(self, legacy: Dict[str, Any]) -> Dict[str, Any]:
        new_order: Dict[str, Any] = {"categories": {}, "folder": {}}
        main_list = legacy.get("main", [])
        if isinstance(main_list, list):
            new_order["categories"][self.category_path_key([])] = list(main_list)
        sub_map = legacy.get("sub", {})
        if isinstance(sub_map, dict):
            for main_name, sub_list in sub_map.items():
                if isinstance(sub_list, list):
                    new_order["categories"][self.category_path_key([str(main_name)])] = list(sub_list)
        folder_map = legacy.get("folder", {})
        if isinstance(folder_map, dict):
            for main_name, sub_entries in folder_map.items():
                if not isinstance(sub_entries, dict):
                    continue
                for sub_name, paths in sub_entries.items():
                    if isinstance(paths, list):
                        key = self.category_path_key([str(main_name), str(sub_name)])
                        new_order["folder"][key] = list(paths)
        return new_order

    def ordered_list(self, items: List[str], preferred: List[str]) -> List[str]:
        ordered = [x for x in preferred if x in items]
        for x in items:
            if x not in ordered:
                ordered.append(x)
        return ordered

    def folders_sorted_for_category(self, folders: List[Dict[str, str]], order_paths: List[str]) -> List[Dict[str, str]]:
        mapping = {f["path"]: f for f in folders}
        ordered = []
        for path in order_paths:
            item = mapping.pop(path, None)
            if item:
                ordered.append(item)
        ordered.extend(sorted(mapping.values(), key=lambda x: x["name"].lower()))
        return ordered

    def folder_key(self, folder_path: str) -> str:
        return os.path.normcase(os.path.abspath(folder_path))

    def count_immediate_subfolders(self, folder_path: str) -> int:
        if not os.path.isdir(folder_path):
            return 0
        try:
            count = 0
            for name in os.listdir(folder_path):
                if name.lower() in {"_history", LEGACY_META_DIR.lower()}:
                    continue
                full = os.path.join(folder_path, name)
                if os.path.isdir(full):
                    count += 1
            return count
        except Exception:
            return 0

    def folder_subfolder_counts(self) -> Dict[str, int]:
        data = self.settings.get("folder_subfolder_counts", {})
        if not isinstance(data, dict):
            data = {}
        return {
            self.folder_key(str(key)): int(value)
            for key, value in data.items()
            if isinstance(key, str)
        }

    def save_folder_subfolder_counts(self, counts: Dict[str, int]) -> None:
        self.settings["folder_subfolder_counts"] = counts
        save_settings(self.settings)

    def folder_subfolder_count_for_path(self, folder_path: str) -> Optional[int]:
        if not folder_path:
            return None
        counts = self.folder_subfolder_counts()
        return counts.get(self.folder_key(folder_path))

    def set_folder_subfolder_count(self, folder_path: str, count: int) -> None:
        if not folder_path:
            return
        counts = self.folder_subfolder_counts()
        counts[self.folder_key(folder_path)] = int(count)
        self.save_folder_subfolder_counts(counts)

    def record_folder_subfolder_count(self, folder_path: str) -> None:
        if not folder_path:
            return
        counts = self.folder_subfolder_counts()
        counts[self.folder_key(folder_path)] = self.count_immediate_subfolders(folder_path)
        self.save_folder_subfolder_counts(counts)

    def remove_folder_settings_for_paths(self, paths: List[str]) -> None:
        if not paths:
            return
        keys = {self.folder_key(path) for path in paths if path}
        if not keys:
            return
        folder_checks = self.folder_tree_check_states()
        folder_checks = {k: v for k, v in folder_checks.items() if k not in keys}
        self.settings["folder_tree_check_states"] = folder_checks
        counts = self.folder_subfolder_counts()
        counts = {k: v for k, v in counts.items() if k not in keys}
        self.settings["folder_subfolder_counts"] = counts
        save_settings(self.settings)

    def remove_category_settings_under_path(self, path: List[str]) -> List[str]:
        target_key = self.category_path_key(path)
        prefix = f"{target_key}{CATEGORY_PATH_SEP}" if target_key else ""
        removed_paths: List[str] = []
        paths = self.category_folder_paths()
        remaining_paths = {}
        for key, value in paths.items():
            if key == target_key or (prefix and key.startswith(prefix)):
                removed_paths.append(value)
            else:
                remaining_paths[key] = value
        self.settings["category_folder_paths"] = remaining_paths
        checks = self.category_check_states()
        remaining_checks = {
            key: value
            for key, value in checks.items()
            if key != target_key and not (prefix and key.startswith(prefix))
        }
        self.settings["category_check_states"] = remaining_checks
        save_settings(self.settings)
        return removed_paths

    def clear_unused_cache(self) -> Tuple[int, int]:
        registry_paths = {
            self.folder_key(item["path"])
            for item in self.registry
            if isinstance(item.get("path"), str)
        }
        category_paths: set[str] = set()
        for item in self.registry:
            categories = self.category_path_for_item(item)
            for i in range(1, len(categories) + 1):
                category_paths.add(self.category_path_key(categories[:i]))
        order = self.category_order()
        for key in order.get("categories", {}):
            if isinstance(key, str) and key:
                category_paths.add(key)
        folder_paths = set(registry_paths)
        category_folder_paths = self.category_folder_paths()
        filtered_category_folder_paths = {
            key: value
            for key, value in category_folder_paths.items()
            if key in category_paths
        }
        for folder_path in filtered_category_folder_paths.values():
            if isinstance(folder_path, str):
                folder_paths.add(self.folder_key(folder_path))

        settings_removed = 0
        checks = self.category_check_states()
        filtered_checks = {k: v for k, v in checks.items() if k in category_paths}
        settings_removed += len(checks) - len(filtered_checks)
        self.settings["category_check_states"] = filtered_checks

        folder_checks = self.folder_tree_check_states()
        filtered_folder_checks = {k: v for k, v in folder_checks.items() if k in folder_paths}
        settings_removed += len(folder_checks) - len(filtered_folder_checks)
        self.settings["folder_tree_check_states"] = filtered_folder_checks

        counts = self.folder_subfolder_counts()
        filtered_counts = {k: v for k, v in counts.items() if k in folder_paths}
        settings_removed += len(counts) - len(filtered_counts)
        self.settings["folder_subfolder_counts"] = filtered_counts

        settings_removed += len(category_folder_paths) - len(filtered_category_folder_paths)
        self.settings["category_folder_paths"] = filtered_category_folder_paths

        save_settings(self.settings)

        user_checks_removed = 0
        folder_checks = self.user_checks
        filtered_user_checks = {k: v for k, v in folder_checks.items() if k in folder_paths}
        user_checks_removed = len(folder_checks) - len(filtered_user_checks)
        self.user_checks = filtered_user_checks
        save_user_checks(self.user_checks)

        return settings_removed, user_checks_removed

    def preview_subfolder_counts(self, root_path: str, items: List[Dict[str, Any]]) -> Dict[str, int]:
        paths = {
            entry.get("path")
            for entry in items
            if isinstance(entry.get("path"), str)
        }
        paths.add(root_path)
        counts: Dict[str, int] = {}
        for path in paths:
            if isinstance(path, str) and path:
                counts[path] = self.count_immediate_subfolders(path)
        return counts

    def detect_new_subfolders(self) -> set[str]:
        counts = self.folder_subfolder_counts()
        highlights: set[str] = set()
        updated_counts = False
        targets: set[str] = {
            item["path"]
            for item in self.registry
            if isinstance(item.get("path"), str)
        }
        for folder_path in self.category_folder_paths().values():
            if isinstance(folder_path, str) and folder_path:
                targets.add(folder_path)
        for root_path in targets:
            if not root_path or not os.path.isdir(root_path):
                continue
            root_key = self.folder_key(root_path)
            current_count = self.count_immediate_subfolders(root_path)
            if root_key not in counts:
                counts[root_key] = current_count
                updated_counts = True
                continue
            previous_count = counts.get(root_key, current_count)
            if current_count > previous_count:
                highlights.add(root_key)
        if updated_counts:
            self.save_folder_subfolder_counts(counts)
        return highlights

    def update_new_folder_highlights(self) -> None:
        self.new_folder_highlights = self.detect_new_subfolders()
        self.new_category_highlights = set()
        category_paths_by_folder: Dict[str, set[str]] = {}
        for path, folder_path in self.category_folder_paths().items():
            if isinstance(folder_path, str) and folder_path:
                category_paths_by_folder.setdefault(self.folder_key(folder_path), set()).add(path)
        for item in self.registry:
            categories = self.category_path_for_item(item)
            folder_path = item.get("path")
            if isinstance(folder_path, str) and folder_path:
                category_paths_by_folder.setdefault(self.folder_key(folder_path), set()).update(
                    self.category_path_key(categories[:i]) for i in range(1, len(categories) + 1)
                )
        for folder_key, paths in category_paths_by_folder.items():
            if folder_key in self.new_folder_highlights:
                self.new_category_highlights.update(paths)

    def doc_is_checked(self, folder_path: str, doc_key: str, doc_info: Optional[Dict[str, Any]] = None) -> bool:
        folder_key = self.folder_key(folder_path)
        folder_checks = self.user_checks.get(folder_key, {})
        if doc_key in folder_checks:
            return bool(folder_checks.get(doc_key))
        if doc_info and isinstance(doc_info.get("user_checks"), dict):
            legacy_checked = bool(doc_info["user_checks"].get(self.current_user, False))
            if legacy_checked:
                self.set_doc_checked(folder_path, doc_key, True)
            return legacy_checked
        return False

    def set_doc_checked(self, folder_path: str, doc_key: str, checked: bool) -> None:
        folder_key = self.folder_key(folder_path)
        folder_checks = self.user_checks.get(folder_key, {})
        if not isinstance(folder_checks, dict):
            folder_checks = {}
        folder_checks[doc_key] = checked
        self.user_checks[folder_key] = folder_checks
        save_user_checks(self.user_checks)

    def mark_doc_checked(self, folder_path: str, doc_key: str) -> None:
        self.set_doc_checked(folder_path, doc_key, True)

    def mark_all_docs_checked(self, folder_path: str) -> None:
        if not os.path.isdir(folder_path):
            return
        meta, _rows = scan_folder(folder_path, self.ignore_types)
        docs = meta.get("documents", {})
        if not isinstance(docs, dict) or not docs:
            return
        folder_key = self.folder_key(folder_path)
        folder_checks = self.user_checks.get(folder_key, {})
        if not isinstance(folder_checks, dict):
            folder_checks = {}
        updated = False
        for doc_key in docs.keys():
            if folder_checks.get(doc_key) is not True:
                folder_checks[doc_key] = True
                updated = True
        if updated:
            self.user_checks[folder_key] = folder_checks
            save_user_checks(self.user_checks)

    def folder_has_unchecked(self, folder_path: str) -> bool:
        if not os.path.isdir(folder_path):
            return False
        meta, _rows = scan_folder(folder_path, self.ignore_types)
        docs = meta.get("documents", {})
        if not isinstance(docs, dict):
            return False
        for doc_key, info in docs.items():
            if not isinstance(info, dict):
                continue
            if not info.get("current_file"):
                continue
            if not self.doc_is_checked(folder_path, doc_key, info):
                return True
        return False

    def set_item_unchecked_style(self, item: QTableWidgetItem, unchecked: bool) -> None:
        if unchecked:
            item.setForeground(QBrush(UNCHECKED_COLOR))
        else:
            item.setForeground(QBrush())

    def set_item_new_folder_style(self, item: QTableWidgetItem, has_new: bool) -> None:
        if has_new:
            item.setBackground(QBrush(self.new_folder_bg_color()))
        else:
            item.setBackground(QBrush())

    def new_folder_bg_color(self) -> QColor:
        base_color = self.palette().color(QPalette.Base)
        if base_color.lightness() < 128:
            return NEW_FOLDER_BG_COLOR_DARK
        return NEW_FOLDER_BG_COLOR_LIGHT

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
        category_options = self.category_options()
        subfolder_count = self.folder_subfolder_count_for_path(item.get("path", "")) or 0

        dlg = RegisterDialog(
            self,
            initial_name=item.get("name", ""),
            initial_path=item.get("path", ""),
            initial_categories=self.category_path_for_item(item),
            category_options=category_options,
            ok_label="更新",
            subfolder_count=subfolder_count,
        )
        if dlg.exec() != QDialog.Accepted:
            return

        name, new_path, categories = dlg.get_values()
        if not name or not new_path or not categories:
            self.warn("登録名・フォルダパス・カテゴリ階層1を入力してください。")
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
            new_meta = meta_path_for_folder(new_path)
            new_legacy_meta = legacy_meta_path_for_folder(new_path)
            if os.path.exists(new_history) or os.path.exists(new_meta) or os.path.exists(new_legacy_meta):
                self.warn("移動先に _History またはメタデータが既に存在します。更新をキャンセルしました。")
                return
            old_history = os.path.join(old_path, "_History")
            old_meta = meta_path_for_folder(old_path)
            old_legacy_meta = legacy_meta_path_for_folder(old_path)
            try:
                if os.path.exists(old_history):
                    shutil.move(old_history, new_history)
                if os.path.exists(old_meta):
                    shutil.move(old_meta, new_meta)
                elif os.path.exists(old_legacy_meta):
                    os.makedirs(os.path.dirname(new_legacy_meta), exist_ok=True)
                    shutil.move(old_legacy_meta, new_legacy_meta)
            except Exception as e:
                self.warn(f"_History/メタデータの移動に失敗しました: {e}")
                return

        item["name"] = name
        item["path"] = new_path
        item["categories"] = categories
        self.registry[idx] = item
        save_registry(self.registry)
        self.record_folder_subfolder_count(new_path)
        self.refresh_folder_table()
        self.refresh_category_tree()
        if self.current_folder and os.path.normcase(self.current_folder["path"]) == os.path.normcase(old_path):
            self.current_folder = {"name": name, "path": new_path}
            self.refresh_files_table()
        self.info("更新しました。")

    def edit_category_path(self, path: List[str]):
        if not path:
            self.warn("カテゴリが見つかりません。")
            return
        initial_name = path[-1]
        initial_folder_path = self.category_folder_path_for_path(path)
        subfolder_count = self.folder_subfolder_count_for_path(initial_folder_path) or 0
        dlg = CategoryEditDialog(
            self,
            initial_name=initial_name,
            initial_path=initial_folder_path,
            subfolder_count=subfolder_count,
        )
        if dlg.exec() != QDialog.Accepted:
            return
        new_name, folder_path = dlg.get_values()
        if not new_name:
            self.warn("カテゴリ名を入力してください。")
            return
        if folder_path and not os.path.isdir(folder_path):
            self.warn("フォルダが存在しません。")
            return
        parent_path = path[:-1]
        root = self.build_category_tree_root()
        node = root
        for part in parent_path:
            node = node["children"].get(part, {"children": {}})
        sibling_names = set(node.get("children", {}).keys())
        if new_name != initial_name and new_name in sibling_names:
            self.warn("同じ階層に同名のカテゴリがあります。")
            return
        new_path = parent_path + [new_name]
        registry_changed = False
        if new_path != path:
            for item in self.registry:
                categories = categories_from_item(item)
                if categories[:len(path)] == path:
                    item["categories"] = self.replace_category_prefix(categories, path, new_path)
                    registry_changed = True
            if registry_changed:
                save_registry(self.registry)
            self.rename_category_in_settings(path, new_path)
            if self.selected_category_path[:len(path)] == path:
                self.selected_category_path = self.replace_category_prefix(
                    self.selected_category_path,
                    path,
                    new_path,
                )
        self.set_category_folder_path(new_path, folder_path)
        self.refresh_folder_table()
        self.refresh_category_tree()
        self.refresh_files_table()
        self.info("更新しました。")

    def delete_registered_folder(
        self,
        path: str,
        confirm: bool = True,
        show_info: bool = True,
    ) -> bool:
        idx = self.registry_index_by_path(path)
        if idx < 0:
            self.warn("登録情報が見つかりません。")
            return False
        if confirm and not self.ask("この登録を削除しますか？（_History/メタデータ は削除しません）"):
            return False
        self.registry.pop(idx)
        save_registry(self.registry)
        self.remove_folder_settings_for_paths([path])
        if self.current_folder and os.path.normcase(self.current_folder["path"]) == os.path.normcase(path):
            self.current_folder = None
            self.current_meta = None
            self.current_file_rows = []
        self.refresh_folder_table()
        self.refresh_category_tree()
        self.refresh_files_table()
        if show_info:
            self.info("削除しました。")
        return True

    def delete_category_hierarchy(self, path: List[str], show_info: bool = True) -> bool:
        if not path:
            return False
        targets = [
            item for item in self.registry
            if self.category_path_for_item(item)[:len(path)] == path
        ]
        if targets:
            if not self.ask(f"カテゴリ配下の登録 {len(targets)} 件を削除しますか？"):
                return False
        else:
            if not self.ask("このカテゴリを削除しますか？"):
                return False
        target_paths = {item["path"] for item in targets}
        self.registry = [item for item in self.registry if item["path"] not in target_paths]
        if target_paths:
            self.remove_folder_settings_for_paths(list(target_paths))
            self.remove_user_checks_for_paths(target_paths)
        self.remove_archived_under_path(path)
        removed_category_paths = self.remove_category_settings_under_path(path)
        self.remove_folder_settings_for_paths(removed_category_paths)
        order = self.category_order()
        target_key = self.category_path_key(path)
        parent_key = self.category_path_key(path[:-1])
        if path and parent_key in order.get("categories", {}):
            order["categories"][parent_key] = [
                name for name in order["categories"][parent_key]
                if name != path[-1]
            ]
            if not order["categories"][parent_key]:
                order["categories"].pop(parent_key, None)
        if path and parent_key in order.get("tree", {}):
            order["tree"][parent_key] = [
                entry for entry in order["tree"][parent_key]
                if isinstance(entry, dict)
                and not (entry.get("type") == "category" and entry.get("name") == path[-1])
            ]
            if not order["tree"][parent_key]:
                order["tree"].pop(parent_key, None)
        prefix = f"{target_key}{CATEGORY_PATH_SEP}" if target_key else ""
        for key in list(order.get("categories", {}).keys()):
            if key == target_key or (prefix and key.startswith(prefix)):
                order["categories"].pop(key, None)
        for key in list(order.get("tree", {}).keys()):
            if key == target_key or (prefix and key.startswith(prefix)):
                order["tree"].pop(key, None)
        for key in list(order.get("folder", {}).keys()):
            if key == target_key or (prefix and key.startswith(prefix)):
                order["folder"].pop(key, None)
        self.settings["category_order"] = order
        save_registry(self.registry)
        save_settings(self.settings)
        if self.current_folder and self.current_folder["path"] in target_paths:
            self.current_folder = None
            self.current_meta = None
            self.current_file_rows = []
        self.refresh_folder_table()
        self.refresh_category_tree()
        self.refresh_files_table()
        if show_info:
            self.info("削除しました。")
        return True

    def remove_user_checks_for_paths(self, paths: set[str]) -> None:
        for path in paths:
            key = self.folder_key(path)
            if key in self.user_checks:
                self.user_checks.pop(key, None)
        save_user_checks(self.user_checks)

    def remove_user_checks_for_docs(self, folder_path: str, doc_keys: set[str]) -> None:
        folder_key = self.folder_key(folder_path)
        folder_checks = self.user_checks.get(folder_key, {})
        if not folder_checks:
            return
        for doc_key in doc_keys:
            folder_checks.pop(doc_key, None)
        if folder_checks:
            self.user_checks[folder_key] = folder_checks
        else:
            self.user_checks.pop(folder_key, None)
        save_user_checks(self.user_checks)

    def collapse_category_paths(self, paths: List[List[str]]) -> List[List[str]]:
        unique: List[List[str]] = []
        for path in paths:
            if path not in unique:
                unique.append(path)
        unique.sort(key=len)
        collapsed: List[List[str]] = []
        for path in unique:
            if any(path[:len(existing)] == existing for existing in collapsed):
                continue
            collapsed.append(path)
        return collapsed

    # ---------- refresh ----------
    def build_category_tree_root(self) -> Dict[str, Any]:
        root = {"children": {}, "folders": []}
        for item in self.registry:
            categories = self.category_path_for_item(item)
            if self.is_archived_path(categories):
                continue
            node = root
            for category in categories:
                node = node["children"].setdefault(category, {"children": {}, "folders": []})
            node["folders"].append(item)

        order = self.category_order()
        for parent_key, children in order.get("categories", {}).items():
            if not isinstance(children, list):
                continue
            parent_path = self.category_path_from_key(str(parent_key))
            node = root
            for category in parent_path:
                node = node["children"].setdefault(category, {"children": {}, "folders": []})
            for name in children:
                if isinstance(name, str) and name:
                    node["children"].setdefault(name, {"children": {}, "folders": []})
        return root

    def capture_expanded_category_paths(self) -> Optional[set[str]]:
        if self.category_tree.topLevelItemCount() == 0:
            return None

        expanded: set[str] = set()

        def walk(item: QTreeWidgetItem):
            data = item.data(0, Qt.UserRole) or {}
            if data.get("type") == "category":
                path = data.get("path", [])
                if isinstance(path, list) and item.isExpanded():
                    expanded.add(self.category_path_key(path))
            for i in range(item.childCount()):
                walk(item.child(i))

        for i in range(self.category_tree.topLevelItemCount()):
            walk(self.category_tree.topLevelItem(i))
        return expanded

    def find_category_tree_item(self, path: List[str]) -> Optional[QTreeWidgetItem]:
        def walk(item: QTreeWidgetItem) -> Optional[QTreeWidgetItem]:
            data = item.data(0, Qt.UserRole) or {}
            if data.get("type") == "category" and data.get("path") == path:
                return item
            for i in range(item.childCount()):
                found = walk(item.child(i))
                if found:
                    return found
            return None

        for i in range(self.category_tree.topLevelItemCount()):
            found = walk(self.category_tree.topLevelItem(i))
            if found:
                return found
        return None

    def refresh_folder_table(self):
        preserve_key: Optional[Tuple[str, str]] = None
        if self.current_folder:
            preserve_key = ("folder", self.current_folder["path"])
        self.folders_table.blockSignals(True)
        query = self.search.text().strip().lower()
        items: List[Dict[str, Any]] = []
        root = self.build_category_tree_root()
        node = root
        for category in self.selected_category_path:
            if category not in node["children"]:
                node = None
                break
            node = node["children"][category]

        if node is not None:
            order = self.category_order()
            key = self.category_path_key(self.selected_category_path)
            child_names = self.ordered_list(
                list(node["children"].keys()),
                order.get("categories", {}).get(key, []),
            )
            folder_order = order.get("folder", {}).get(key, [])
            folder_mapping = {folder["path"]: folder for folder in node["folders"]}
            remaining_folders = self.folders_sorted_for_category(node["folders"], folder_order)
            tree_order = order.get("tree", {}).get(key, [])
            ordered_entries: List[Dict[str, Any]] = []
            added_categories: set[str] = set()
            added_folders: set[str] = set()
            if isinstance(tree_order, list):
                for entry in tree_order:
                    if not isinstance(entry, dict):
                        continue
                    entry_type = entry.get("type")
                    if entry_type == "category":
                        name = entry.get("name")
                        if isinstance(name, str) and name in node["children"] and name not in added_categories:
                            ordered_entries.append({"type": "category", "name": name})
                            added_categories.add(name)
                    elif entry_type == "folder":
                        folder_path = entry.get("path")
                        if isinstance(folder_path, str) and folder_path in folder_mapping and folder_path not in added_folders:
                            folder = folder_mapping[folder_path]
                            ordered_entries.append({"type": "folder", "name": folder["name"], "path": folder["path"]})
                            added_folders.add(folder_path)

            for name in child_names:
                if name not in added_categories:
                    ordered_entries.append({"type": "category", "name": name})
                    added_categories.add(name)

            for folder in remaining_folders:
                if folder["path"] in added_folders:
                    continue
                ordered_entries.append({"type": "folder", "name": folder["name"], "path": folder["path"]})
                added_folders.add(folder["path"])

            for entry in ordered_entries:
                entry_type = entry["type"]
                name = entry["name"]
                if query and query not in name.lower():
                    continue
                if entry_type == "category":
                    child_path = self.selected_category_path + [name]
                    items.append({
                        "type": "category",
                        "name": name,
                        "path": child_path,
                    })
                elif entry_type == "folder":
                    folder_path = entry["path"]
                    folder_item = folder_mapping.get(folder_path)
                    items.append({
                        "type": "folder",
                        "name": name,
                        "path": folder_path,
                        "categories": self.category_path_for_item(folder_item) if folder_item else [],
                    })

        self.folders_table.setRowCount(0)

        for item in items:
            item_type = item.get("type")
            name = item["name"]
            path = item["path"]
            last_date = ""
            has_unchecked = False
            has_new_subfolder = False
            if item_type == "folder":
                categories = item.get("categories") or []
                highlight_enabled = (
                    self.is_category_highlight_enabled_for_path(categories)
                    and self.is_folder_tree_checked(path)
                )
                has_unchecked = highlight_enabled and self.folder_has_unchecked(path)
                has_new_subfolder = (
                    highlight_enabled
                    and self.folder_key(path) in self.new_folder_highlights
                )

                if os.path.isdir(path):
                    try:
                        files = safe_list_files(path, self.ignore_types)
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
                            last_date = dt.datetime.fromtimestamp(latest_mtime).strftime("%Y-%m-%d")
                    except Exception:
                        pass

            r = self.folders_table.rowCount()
            self.folders_table.insertRow(r)

            if item_type == "category":
                category_folder_path = self.category_folder_path_for_path(path)
                icon_prefix = "📁 " if category_folder_path else "🔖 "
                it_name = QTableWidgetItem(f"{icon_prefix}{name}")
                if category_folder_path:
                    it_name.setToolTip(category_folder_path)
                    highlight_enabled = self.is_category_highlight_enabled_for_path(path)
                    if highlight_enabled and self.category_path_key(path) in self.new_category_highlights:
                        has_new_subfolder = True
                it_name.setData(Qt.UserRole, {"type": "category", "path": path})
            else:
                icon_prefix = "📁 "
                it_name = QTableWidgetItem(f"{icon_prefix}{name}")
                it_name.setToolTip(path)
                it_name.setData(Qt.UserRole, path)
            self.set_item_unchecked_style(it_name, has_unchecked)
            self.set_item_new_folder_style(it_name, has_new_subfolder)

            self.folders_table.setItem(r, 0, it_name)
            it_date = QTableWidgetItem(last_date)
            self.set_item_new_folder_style(it_date, has_new_subfolder)
            self.folders_table.setItem(r, 1, it_date)

        self.folders_table.repaint()
        self.folders_table.blockSignals(False)
        if preserve_key:
            if preserve_key[0] == "folder":
                self.select_folder_in_table(preserve_key[1])

    def refresh_category_tree(self):
        self._category_tree_refreshing = True
        expanded_paths = self.capture_expanded_category_paths()
        self.category_tree.blockSignals(True)
        self.category_tree.clear()
        order = self.category_order()
        root = self.build_category_tree_root()

        def add_nodes(
            parent_item: Optional[QTreeWidgetItem],
            node: Dict[str, Any],
            path: List[str],
            highlight_enabled: bool,
        ) -> bool:
            key = self.category_path_key(path)
            has_unchecked = False
            child_names = self.ordered_list(
                list(node["children"].keys()),
                order.get("categories", {}).get(key, []),
            )
            folder_order = order.get("folder", {}).get(key, [])
            folder_mapping = {folder["path"]: folder for folder in node["folders"]}
            remaining_folders = self.folders_sorted_for_category(node["folders"], folder_order)
            remaining_categories = list(child_names)
            added_categories: set[str] = set()
            added_folders: set[str] = set()

            def add_category(name: str) -> None:
                nonlocal has_unchecked
                child_node = node["children"][name]
                child_path = path + [name]
                folder_path = self.category_folder_path_for_path(child_path)
                icon_prefix = "📁 " if folder_path else "🔖 "
                child_item = QTreeWidgetItem([f"{icon_prefix}{name}"])
                child_item.setData(0, Qt.UserRole, {"type": "category", "path": child_path})
                if folder_path:
                    child_item.setToolTip(0, folder_path)
                child_item.setFlags(child_item.flags() | Qt.ItemIsUserCheckable)
                child_checked = self.is_category_checked(child_path)
                child_item.setCheckState(0, Qt.Checked if child_checked else Qt.Unchecked)
                if parent_item is None:
                    self.category_tree.addTopLevelItem(child_item)
                else:
                    parent_item.addChild(child_item)
                if expanded_paths is not None and self.category_path_key(child_path) in expanded_paths:
                    child_item.setExpanded(True)
                child_highlight_enabled = highlight_enabled and child_checked
                child_unchecked = add_nodes(child_item, child_node, child_path, child_highlight_enabled)
                if child_highlight_enabled and self.category_path_key(child_path) in self.new_category_highlights:
                    child_item.setBackground(0, QBrush(self.new_folder_bg_color()))
                if child_unchecked and child_highlight_enabled:
                    child_item.setForeground(0, QBrush(UNCHECKED_COLOR))
                    has_unchecked = True

            def add_folder(folder: Dict[str, str]) -> None:
                nonlocal has_unchecked
                folder_item = QTreeWidgetItem([f"📁 {folder['name']}"])
                folder_item.setToolTip(0, folder["path"])
                folder_item.setFlags(folder_item.flags() | Qt.ItemIsUserCheckable)
                folder_checked = self.is_folder_tree_checked(folder["path"])
                folder_item.setCheckState(0, Qt.Checked if folder_checked else Qt.Unchecked)
                folder_item.setData(0, Qt.UserRole, {
                    "type": "folder",
                    "path": folder["path"],
                    "category_path": path,
                })
                folder_highlight_enabled = highlight_enabled and folder_checked
                if folder_highlight_enabled and self.folder_key(folder["path"]) in self.new_folder_highlights:
                    folder_item.setBackground(0, QBrush(self.new_folder_bg_color()))
                if folder_highlight_enabled:
                    folder_unchecked = self.folder_has_unchecked(folder["path"])
                    if folder_unchecked:
                        folder_item.setForeground(0, QBrush(UNCHECKED_COLOR))
                        has_unchecked = True
                if parent_item is None:
                    self.category_tree.addTopLevelItem(folder_item)
                else:
                    parent_item.addChild(folder_item)

            tree_order = order.get("tree", {}).get(key, [])
            if isinstance(tree_order, list):
                for entry in tree_order:
                    if not isinstance(entry, dict):
                        continue
                    entry_type = entry.get("type")
                    if entry_type == "category":
                        name = entry.get("name")
                        if isinstance(name, str) and name in node["children"] and name not in added_categories:
                            add_category(name)
                            added_categories.add(name)
                    elif entry_type == "folder":
                        folder_path = entry.get("path")
                        if isinstance(folder_path, str) and folder_path in folder_mapping and folder_path not in added_folders:
                            add_folder(folder_mapping[folder_path])
                            added_folders.add(folder_path)

            for name in remaining_categories:
                if name in added_categories:
                    continue
                add_category(name)
                added_categories.add(name)

            for folder in remaining_folders:
                if folder["path"] in added_folders:
                    continue
                add_folder(folder)
                added_folders.add(folder["path"])
            return has_unchecked if highlight_enabled else False

        add_nodes(None, root, [], True)

        self.category_tree.blockSignals(False)
        self._category_tree_refreshing = False

    def schedule_category_tree_refresh(self):
        if self._category_tree_refresh_pending:
            return
        self._category_tree_refresh_pending = True
        QTimer.singleShot(0, self.apply_category_tree_refresh)

    def apply_category_tree_refresh(self):
        self._category_tree_refresh_pending = False
        self.refresh_category_tree()

    def update_category_order_from_tree(self):
        order = {"categories": {}, "folder": {}, "tree": {}}

        def walk(item: QTreeWidgetItem, path: List[str]):
            category_order = []
            folder_order = []
            tree_order = []
            for i in range(item.childCount()):
                child = item.child(i)
                data = child.data(0, Qt.UserRole) or {}
                if data.get("type") == "category":
                    child_path = data.get("path")
                    if isinstance(child_path, list) and child_path:
                        name = child_path[-1]
                    else:
                        name = self.strip_icon_prefix(child.text(0))
                    category_order.append(name)
                    tree_order.append({"type": "category", "name": name})
                elif data.get("type") == "folder":
                    folder_path = data.get("path")
                    if folder_path:
                        folder_order.append(folder_path)
                        tree_order.append({"type": "folder", "path": folder_path})
            if category_order:
                order["categories"][self.category_path_key(path)] = category_order
            if folder_order:
                order["folder"][self.category_path_key(path)] = folder_order
            if tree_order:
                order["tree"][self.category_path_key(path)] = tree_order
            for i in range(item.childCount()):
                child = item.child(i)
                data = child.data(0, Qt.UserRole) or {}
                if data.get("type") == "category":
                    child_path = data.get("path")
                    if isinstance(child_path, list) and child_path:
                        name = child_path[-1]
                    else:
                        name = self.strip_icon_prefix(child.text(0))
                    walk(child, path + [name])

        root_categories = []
        root_folders = []
        root_tree = []
        for i in range(self.category_tree.topLevelItemCount()):
            top = self.category_tree.topLevelItem(i)
            data = top.data(0, Qt.UserRole) or {}
            if data.get("type") == "category":
                top_path = data.get("path")
                if isinstance(top_path, list) and top_path:
                    name = top_path[-1]
                else:
                    name = self.strip_icon_prefix(top.text(0))
                root_categories.append(name)
                root_tree.append({"type": "category", "name": name})
                walk(top, [name])
            elif data.get("type") == "folder":
                path = data.get("path")
                if path:
                    root_folders.append(path)
                    root_tree.append({"type": "folder", "path": path})
        if root_categories:
            order["categories"][self.category_path_key([])] = root_categories
        if root_folders:
            order["folder"][self.category_path_key([])] = root_folders
        if root_tree:
            order["tree"][self.category_path_key([])] = root_tree
        self.settings["category_order"] = order
        save_settings(self.settings)

    def select_folder_in_table(self, path: str):
        for row in range(self.folders_table.rowCount()):
            item = self.folders_table.item(row, 0)
            if item and item.data(Qt.UserRole) == path:
                self.folders_table.selectRow(row)
                break

    def select_doc_key(self, doc_key: str):
        for row in range(self.files_table.rowCount()):
            item = self.files_table.item(row, 5)
            if item and item.text() == doc_key:
                self.files_table.selectRow(row)
                self.refresh_right_pane_for_doc(doc_key)
                break

    def refresh_files_table(self):
        self.files_table.blockSignals(True)
        self.files_table.setUpdatesEnabled(False)
        selected_doc_key = None
        selected_index = self.selected_file_index()
        if 0 <= selected_index < len(self.current_file_rows):
            selected_doc_key = self.current_file_rows[selected_index].doc_key
        self.hist_table.setRowCount(0)

        if not self.current_folder:
            self.update_files_header_check_state()
            self.files_table.blockSignals(False)
            self.files_table.setUpdatesEnabled(True)
            return

        folder_path = self.current_folder["path"]
        if not os.path.isdir(folder_path):
            self.warn("登録フォルダが見つかりません。パスを確認してください。")
            self.update_files_header_check_state()
            self.files_table.blockSignals(False)
            self.files_table.setUpdatesEnabled(True)
            return

        meta, rows = scan_folder(folder_path, self.ignore_types)
        self.current_meta = meta
        self.current_file_rows = rows

        self.files_table.setRowCount(len(rows))
        for r, row in enumerate(rows):

            doc_info = self.current_meta.get("documents", {}).get(row.doc_key, {})
            is_checked = self.doc_is_checked(folder_path, row.doc_key, doc_info if isinstance(doc_info, dict) else None)

            it_check = QTableWidgetItem("")
            it_check.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            it_check.setCheckState(Qt.Checked if is_checked else Qt.Unchecked)
            it_check.setData(Qt.UserRole, row.doc_key)
            self.files_table.setItem(r, 0, it_check)

            it_fn = QTableWidgetItem(row.filename)
            it_fn.setToolTip(os.path.join(folder_path, row.filename))
            self.set_item_unchecked_style(it_fn, not is_checked)
            self.files_table.setItem(r, 1, it_fn)
            self.files_table.setItem(r, 2, QTableWidgetItem(display_rev(row.rev)))
            self.files_table.setItem(r, 3, QTableWidgetItem((row.updated_at or "").replace("T", " ")))
            self.files_table.setItem(r, 4, QTableWidgetItem(row.updated_by or ""))
            self.files_table.setItem(r, 5, QTableWidgetItem(row.doc_key))

        # hide DocKey column by default (can be useful for debugging)
        self.files_table.setColumnHidden(5, True)
        self.update_files_header_check_state()
        if selected_doc_key:
            self.select_doc_key(selected_doc_key)
        self.files_table.blockSignals(False)
        self.files_table.setUpdatesEnabled(True)

    def refresh_right_pane_for_doc(self, doc_key: str):
        self.hist_table.setRowCount(0)

        if not self.current_meta:
            return
        docs = self.current_meta.get("documents", {})
        info = docs.get(doc_key)
        if not isinstance(info, dict):
            return

        history = info.get("history", [])
        if not isinstance(history, list):
            history = []

        latest_entry = {
            "kind": "最新",
            "rev": info.get("current_rev", "") or "",
            "updated_at": info.get("updated_at", "") or "",
            "updated_by": info.get("updated_by", "") or "",
            "memo": info.get("last_memo", "") or "",
            "file": info.get("current_file", "") or "",
        }
        self.add_history_row(latest_entry)

        # show newest first
        items = list(history)
        items.reverse()

        for h in items:
            entry = {
                "kind": "履歴",
                "rev": h.get("rev", "") or "",
                "updated_at": h.get("updated_at", "") or "",
                "updated_by": h.get("updated_by", "") or "",
                "memo": h.get("memo", "") or "",
                "file": h.get("file", "") or "",
            }
            self.add_history_row(entry)

    def add_history_row(self, entry: Dict[str, Any]) -> None:
        r = self.hist_table.rowCount()
        self.hist_table.insertRow(r)
        kind_item = QTableWidgetItem(entry.get("kind", "") or "")
        kind_item.setData(Qt.UserRole, entry)
        rev_text = display_rev(entry.get("rev", "") or "")
        updated_text = (entry.get("updated_at", "") or "").replace("T", " ")
        self.hist_table.setItem(r, 0, kind_item)
        self.hist_table.setItem(r, 1, QTableWidgetItem(rev_text))
        self.hist_table.setItem(r, 2, QTableWidgetItem(updated_text))
        self.hist_table.setItem(r, 3, QTableWidgetItem(entry.get("updated_by", "") or ""))
        self.hist_table.setItem(r, 4, QTableWidgetItem(entry.get("memo", "") or ""))

    # ---------- events ----------
    def on_category_tree_order_changed(self):
        self.update_category_order_from_tree()
        self.refresh_category_tree()

    def on_category_tree_item_changed(self, item: QTreeWidgetItem, column: int):
        if self._category_tree_refreshing:
            return
        if column != 0:
            return
        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")
        if item_type == "category":
            path = data.get("path", [])
            if not isinstance(path, list):
                return
            self.set_category_checked(path, item.checkState(0) == Qt.Checked)
        elif item_type == "folder":
            folder_path = data.get("path")
            if not isinstance(folder_path, str):
                return
            self.set_folder_tree_checked(folder_path, item.checkState(0) == Qt.Checked)
        else:
            return
        self.refresh_folder_table()
        self.schedule_category_tree_refresh()

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

    def selected_category_tree_paths(self) -> Tuple[List[List[str]], List[str]]:
        category_paths: List[List[str]] = []
        folder_paths: List[str] = []
        for item in self.category_tree.selectedItems():
            data = item.data(0, Qt.UserRole) or {}
            item_type = data.get("type")
            if item_type == "category":
                path = data.get("path", [])
                if isinstance(path, list) and path:
                    category_paths.append(path)
            elif item_type == "folder":
                path = data.get("path")
                if isinstance(path, str) and path:
                    folder_paths.append(path)
        return category_paths, folder_paths

    def archive_selected_category_paths(self, paths: List[List[str]]) -> None:
        collapsed = self.collapse_category_paths(paths)
        changed = False
        for path in collapsed:
            if self.archive_category_path(path):
                changed = True
        if changed:
            self.refresh_folder_table()
            self.refresh_category_tree()

    def delete_selected_category_tree_items(
        self,
        category_paths: List[List[str]],
        folder_paths: List[str],
    ) -> None:
        collapsed_categories = self.collapse_category_paths(category_paths)
        deleted_any = False
        for path in collapsed_categories:
            if self.delete_category_hierarchy(path, show_info=False):
                deleted_any = True
        if folder_paths:
            if not self.ask(f"{len(folder_paths)} 件の登録を削除しますか？（_History/メタデータ は削除しません）"):
                return
            for path in folder_paths:
                if self.delete_registered_folder(path, confirm=False, show_info=False):
                    deleted_any = True
        if deleted_any:
            self.info("削除しました。")

    def on_category_tree_context_menu(self, pos):
        item = self.category_tree.itemAt(pos)
        if not item:
            return
        if not item.isSelected():
            self.category_tree.clearSelection()
            item.setSelected(True)
            self.category_tree.setCurrentItem(item)
        selected_items = self.category_tree.selectedItems()
        if len(selected_items) > 1:
            category_paths, folder_paths = self.selected_category_tree_paths()
            if not category_paths and not folder_paths:
                return
            menu = QMenu(self)
            act_delete = menu.addAction("削除")
            act_archive = None
            if category_paths:
                act_archive = menu.addAction("アーカイブ")
            action = menu.exec(self.category_tree.viewport().mapToGlobal(pos))
            if action == act_delete:
                self.delete_selected_category_tree_items(category_paths, folder_paths)
            elif action == act_archive:
                self.archive_selected_category_paths(category_paths)
            return
        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")

        menu = QMenu(self)
        act_register = None
        act_batch_register = None
        act_edit = None
        act_delete = None
        act_archive = None
        act_register_as_category = None
        act_register_children = None

        if item_type == "category":
            act_register = menu.addAction("登録")
            act_batch_register = menu.addAction("一括登録")
            category_path = data.get("path", [])
            if isinstance(category_path, list) and self.category_folder_path_for_path(category_path):
                act_register_children = menu.addAction("下位フォルダを登録")
            act_edit = menu.addAction("編集")
            act_delete = menu.addAction("削除")
            act_archive = menu.addAction("アーカイブ")
        elif item_type == "folder":
            act_edit = menu.addAction("編集")
            act_register_as_category = menu.addAction("下位フォルダを登録")
            act_delete = menu.addAction("削除")
        else:
            return

        action = menu.exec(self.category_tree.viewport().mapToGlobal(pos))
        if not action:
            return
        if action == act_register:
            initial_categories = data.get("path", [])
            if not isinstance(initial_categories, list):
                initial_categories = []
            self.open_register_dialog(initial_categories=initial_categories)
        elif action == act_batch_register:
            initial_categories = data.get("path", [])
            if not isinstance(initial_categories, list):
                initial_categories = []
            self.on_batch_register_for_category(initial_categories)
        elif action == act_register_children:
            category_path = data.get("path", [])
            if not isinstance(category_path, list):
                return
            root_path = self.category_folder_path_for_path(category_path)
            if not root_path:
                return
            self.run_batch_register(root_path, 1, base_categories=category_path)
        elif action == act_edit:
            if item_type == "category":
                category_path = data.get("path", [])
                if isinstance(category_path, list):
                    self.edit_category_path(category_path)
            else:
                path = data.get("path")
                if not isinstance(path, str):
                    path = item.toolTip(0)
                if path:
                    self.edit_registered_folder(path)
        elif action == act_register_as_category:
            root_path = data.get("path")
            category_path = data.get("category_path", [])
            folder_name = ""
            if not isinstance(root_path, str) or not root_path:
                return
            if not isinstance(category_path, list):
                category_path = []
            parent_item = item.parent()
            current_category_order = []
            current_tree_order = []
            if parent_item is None:
                for i in range(self.category_tree.topLevelItemCount()):
                    top = self.category_tree.topLevelItem(i)
                    top_data = top.data(0, Qt.UserRole) or {}
                    if top_data.get("type") == "category":
                        top_path = top_data.get("path")
                        if isinstance(top_path, list) and top_path:
                            name = top_path[-1]
                        else:
                            name = self.strip_icon_prefix(top.text(0))
                        current_category_order.append(name)
                        current_tree_order.append({"type": "category", "name": name})
                    elif top_data.get("type") == "folder":
                        top_path = top_data.get("path")
                        if top_path:
                            current_tree_order.append({"type": "folder", "path": top_path})
            else:
                for i in range(parent_item.childCount()):
                    child = parent_item.child(i)
                    child_data = child.data(0, Qt.UserRole) or {}
                    if child_data.get("type") == "category":
                        child_path = child_data.get("path")
                        if isinstance(child_path, list) and child_path:
                            name = child_path[-1]
                        else:
                            name = self.strip_icon_prefix(child.text(0))
                        current_category_order.append(name)
                        current_tree_order.append({"type": "category", "name": name})
                    elif child_data.get("type") == "folder":
                        child_path = child_data.get("path")
                        if child_path:
                            current_tree_order.append({"type": "folder", "path": child_path})
            base_categories = normalize_category_path(category_path)
            if not self.run_batch_register(root_path, 1, base_categories=base_categories):
                return
            removed_item_name = ""
            idx = self.registry_index_by_path(root_path)
            if idx >= 0:
                removed_item_name = str(self.registry[idx].get("name", "") or "")
                self.registry.pop(idx)
                self.remove_user_checks_for_paths({root_path})
                save_registry(self.registry)
                if self.current_folder and os.path.normcase(self.current_folder["path"]) == os.path.normcase(root_path):
                    self.current_folder = None
                    self.current_meta = None
                    self.current_file_rows = []
            if removed_item_name:
                folder_name = removed_item_name
            if not folder_name:
                folder_name = self.strip_icon_prefix(item.text(0).strip())
            if not folder_name:
                folder_name = self.root_category_name(root_path)
            category_folder_path = root_path
            order = self.category_order()
            key = self.category_path_key(category_path)
            folder_order = order.get("folder", {}).get(key, [])
            if isinstance(folder_order, list):
                order["folder"][key] = [
                    path for path in folder_order
                    if os.path.normcase(path) != os.path.normcase(root_path)
                ]
            if folder_name and folder_name not in current_category_order:
                current_category_order.append(folder_name)
            if current_category_order:
                order["categories"][key] = current_category_order
            if current_tree_order:
                updated_tree = []
                replaced = False
                for entry in current_tree_order:
                    if not isinstance(entry, dict):
                        continue
                    if entry.get("type") == "folder" and os.path.normcase(str(entry.get("path", ""))) == os.path.normcase(root_path):
                        updated_tree.append({"type": "category", "name": folder_name})
                        replaced = True
                    else:
                        updated_tree.append(entry)
                if not replaced:
                    updated_tree.append({"type": "category", "name": folder_name})
                order.setdefault("tree", {})
                order["tree"][key] = updated_tree
            self.settings["category_order"] = order
            save_settings(self.settings)
            new_category_path = normalize_category_path(category_path + [folder_name])
            self.set_category_folder_path(new_category_path, category_folder_path)
            self.refresh_folder_table()
            self.refresh_category_tree()
            self.refresh_files_table()
        elif action == act_delete:
            path = data.get("path")
            if item_type == "folder" and not isinstance(path, str):
                path = item.toolTip(0)
            if not path:
                return
            if item_type == "folder":
                self.delete_registered_folder(path)
            else:
                if not isinstance(path, list):
                    return
                self.delete_category_hierarchy(path)
        elif action == act_archive:
            path = data.get("path")
            if not isinstance(path, list):
                return
            if path and self.archive_category_path(path):
                self.refresh_folder_table()
                self.refresh_category_tree()

    def selected_file_doc_keys(self) -> List[str]:
        if not self.current_file_rows:
            return []
        rows = [idx.row() for idx in self.files_table.selectionModel().selectedRows()]
        doc_keys: List[str] = []
        for row in sorted(set(rows)):
            if 0 <= row < len(self.current_file_rows):
                doc_keys.append(self.current_file_rows[row].doc_key)
        return doc_keys

    def remove_selected_files(self, archive: bool) -> None:
        if not self.current_folder:
            self.warn("フォルダを選択してください。")
            return
        doc_keys = self.selected_file_doc_keys()
        if not doc_keys:
            self.warn("対象ファイル（最新）を選択してください。")
            return
        action_label = "アーカイブ" if archive else "削除"
        if not self.ask(f"選択したファイル {len(doc_keys)} 件を{action_label}しますか？"):
            return

        folder_path = self.current_folder["path"]
        meta = load_meta(folder_path)
        docs = meta.get("documents", {})
        if not isinstance(docs, dict):
            self.warn("メタデータが見つかりません。再スキャンしてください。")
            return

        removed_keys: set[str] = set()
        errors = []

        for doc_key in doc_keys:
            info = docs.get(doc_key)
            if not isinstance(info, dict):
                continue
            cur_fn = info.get("current_file", "")
            if cur_fn:
                cur_path = os.path.join(folder_path, cur_fn)
                try:
                    if os.path.exists(cur_path):
                        if archive:
                            history_dir = ensure_history_dir(folder_path)
                            dest_name = cur_fn
                            dest_path = os.path.join(history_dir, dest_name)
                            if os.path.exists(dest_path):
                                ts = dt.datetime.now().strftime("%Y%m%d%H%M%S")
                                base, ext = split_name_ext(cur_fn)
                                dest_name = f"{base}_{ts}{ext}"
                                dest_path = os.path.join(history_dir, dest_name)
                            shutil.move(cur_path, dest_path)
                        else:
                            os.remove(cur_path)
                except Exception as e:
                    errors.append(f"{doc_key}: {e}")
                    continue
            if not archive:
                history_items = info.get("history", [])
                if isinstance(history_items, list):
                    history_dir = ensure_history_dir(folder_path)
                    for entry in history_items:
                        if not isinstance(entry, dict):
                            continue
                        fname = entry.get("file", "")
                        if not fname:
                            continue
                        try:
                            fpath = os.path.join(history_dir, fname)
                            if os.path.exists(fpath):
                                os.remove(fpath)
                        except Exception as e:
                            errors.append(f"{doc_key}: {e}")
            docs.pop(doc_key, None)
            removed_keys.add(doc_key)

        if removed_keys:
            meta["documents"] = docs
            save_meta(folder_path, meta)
            self.remove_user_checks_for_docs(folder_path, removed_keys)
            if self.current_folder and os.path.normcase(self.current_folder["path"]) == os.path.normcase(folder_path):
                self.current_meta = meta
            self.refresh_files_table()
            self.refresh_folder_table()
            self.refresh_category_tree()
            self.info(f"{action_label}しました。")

        if errors:
            self.warn("処理時に一部エラーが発生しました:\n" + "\n".join(errors))

    def on_files_table_context_menu(self, pos):
        item = self.files_table.itemAt(pos)
        if not item:
            return
        row = item.row()
        if row < 0 or row >= len(self.current_file_rows):
            return
        if not self.files_table.selectionModel().isRowSelected(row, self.files_table.rootIndex()):
            self.files_table.clearSelection()
            self.files_table.selectRow(row)

        menu = QMenu(self)
        selected_count = len(self.files_table.selectionModel().selectedRows())
        act_update = None
        act_replace = None
        act_rollback = None
        act_history_clear = None
        if selected_count == 1:
            act_update = menu.addAction("更新")
            act_replace = menu.addAction("差し替え")
            act_rollback = menu.addAction("差し戻し")
            act_history_clear = menu.addAction("History Clear")
            menu.addSeparator()
        act_delete = menu.addAction("削除")
        act_archive = menu.addAction("アーカイブ")
        action = menu.exec(self.files_table.viewport().mapToGlobal(pos))
        if action == act_update:
            self.on_update()
        elif action == act_replace:
            self.on_replace()
        elif action == act_rollback:
            self.on_rollback()
        elif action == act_history_clear:
            self.on_history_clear()
        elif action == act_delete:
            self.remove_selected_files(archive=False)
        elif action == act_archive:
            self.remove_selected_files(archive=True)

    def open_current_file(self, folder_path: str, filename: str, doc_key: str) -> None:
        file_path = os.path.join(folder_path, filename)
        if not os.path.exists(file_path):
            self.warn("現行ファイルが見つかりません。")
            return
        try:
            os.startfile(file_path)  # type: ignore[attr-defined]
        except Exception as e:
            self.warn(f"ファイルを開けませんでした: {e}")
            return
        self.mark_doc_checked(folder_path, doc_key)
        self.refresh_files_table()
        self.refresh_folder_table()
        self.refresh_category_tree()

    def on_file_double_clicked(self, item: QTableWidgetItem):
        row = item.row()
        if row < 0 or row >= len(self.current_file_rows):
            return
        if not self.current_folder:
            return
        row_info = self.current_file_rows[row]
        self.open_current_file(self.current_folder["path"], row_info.filename, row_info.doc_key)

    def on_file_check_changed(self, item: QTableWidgetItem):
        if item.column() != 0:
            return
        if not self.current_folder:
            return
        doc_key = item.data(Qt.UserRole)
        if not doc_key:
            return
        checked = item.checkState() == Qt.Checked
        self.set_doc_checked(self.current_folder["path"], doc_key, checked)
        name_item = self.files_table.item(item.row(), 1)
        if name_item:
            self.set_item_unchecked_style(name_item, not checked)
        self.update_files_header_check_state()
        self.refresh_folder_table()
        self.refresh_category_tree()

    def on_files_header_clicked(self, logical_index: int):
        if logical_index != 0:
            return
        header_item = self.files_table.horizontalHeaderItem(0)
        if not header_item:
            return
        if self.files_table.rowCount() == 0:
            header_item.setCheckState(Qt.Unchecked)
            return
        next_state = Qt.Unchecked if header_item.checkState() == Qt.Checked else Qt.Checked
        header_item.setCheckState(next_state)
        if not self.current_folder:
            return
        folder_path = self.current_folder["path"]
        self.files_table.blockSignals(True)
        for row in range(self.files_table.rowCount()):
            item = self.files_table.item(row, 0)
            if not item:
                continue
            doc_key = item.data(Qt.UserRole)
            if not doc_key:
                continue
            checked = next_state == Qt.Checked
            item.setCheckState(next_state)
            self.set_doc_checked(folder_path, doc_key, checked)
            name_item = self.files_table.item(row, 1)
            if name_item:
                self.set_item_unchecked_style(name_item, not checked)
        self.files_table.blockSignals(False)
        self.refresh_folder_table()
        self.refresh_category_tree()

    def update_files_header_check_state(self):
        header_item = self.files_table.horizontalHeaderItem(0)
        if not header_item:
            return
        total_rows = self.files_table.rowCount()
        if total_rows == 0:
            header_item.setCheckState(Qt.Unchecked)
            return
        checked_rows = 0
        for row in range(total_rows):
            item = self.files_table.item(row, 0)
            if item and item.checkState() == Qt.Checked:
                checked_rows += 1
        if checked_rows == 0:
            header_item.setCheckState(Qt.Unchecked)
        elif checked_rows == total_rows:
            header_item.setCheckState(Qt.Checked)
        else:
            header_item.setCheckState(Qt.PartiallyChecked)

    def open_register_dialog(self, initial_categories: Optional[List[str]] = None):
        category_options = self.category_options()
        dlg = RegisterDialog(
            self,
            initial_categories=initial_categories or [],
            category_options=category_options,
        )
        if dlg.exec() != QDialog.Accepted:
            return
        name, path, categories = dlg.get_values()
        if not name or not path or not categories:
            self.warn("登録名・フォルダパス・カテゴリ階層1を入力してください。")
            return
        if not os.path.isdir(path):
            self.warn("フォルダが存在しません。")
            return

        # Prevent duplicates by path
        for x in self.registry:
            if os.path.normcase(x["path"]) == os.path.normcase(path):
                self.warn("このフォルダは既に登録されています。")
                return

        self.registry.append({
            "name": name,
            "path": path,
            "categories": categories,
        })
        save_registry(self.registry)
        self.record_folder_subfolder_count(path)
        self.mark_all_docs_checked(path)
        self.refresh_folder_table()
        self.refresh_category_tree()
        self.info("登録しました。")

    def on_register(self):
        self.open_register_dialog()

    def on_batch_register(self):
        dlg = BatchRegisterDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return
        root_path = dlg.get_path()
        max_depth = dlg.get_depth()
        if not self.run_batch_register(root_path, max_depth):
            return
        root_category = self.root_category_name(root_path)
        if root_category:
            category_path = [root_category]
            if not self.category_folder_path_for_path(category_path):
                self.set_category_folder_path(category_path, root_path)
                self.refresh_category_tree()

    def run_batch_register(
        self,
        root_path: str,
        max_depth: int,
        base_categories: Optional[List[str]] = None,
    ) -> bool:
        if not root_path:
            self.warn("対象ディレクトリを選択してください。")
            return False
        if not os.path.isdir(root_path):
            self.warn("ディレクトリが存在しません。")
            return False

        items = self.batch_register_items(root_path, max_depth, base_categories=base_categories)
        if not items:
            self.warn("登録できるフォルダがありません。")
            return False

        preview_dialog = BatchPreviewDialog(items, self)
        if preview_dialog.exec() != QDialog.Accepted:
            return False

        selected_items = preview_dialog.selected_items()
        preview_counts = self.preview_subfolder_counts(root_path, items)
        for folder_path, count in preview_counts.items():
            self.set_folder_subfolder_count(folder_path, count)

        if not selected_items:
            self.update_new_folder_highlights()
            self.refresh_folder_table()
            self.refresh_category_tree()
            self.info("登録を行いませんでした。")
            return True

        self.registry.extend(selected_items)
        save_registry(self.registry)
        for item in selected_items:
            folder_path = item.get("path")
            if isinstance(folder_path, str):
                count = preview_counts.get(folder_path)
                if count is not None:
                    self.set_folder_subfolder_count(folder_path, count)
                else:
                    self.record_folder_subfolder_count(folder_path)
                self.mark_all_docs_checked(folder_path)
        self.update_new_folder_highlights()
        self.refresh_folder_table()
        self.refresh_category_tree()
        self.info(f"{len(selected_items)} 件を一括登録しました。")
        return True

    def on_batch_register_for_category(self, category_path: List[str]):
        dlg = BatchRegisterDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return
        root_path = dlg.get_path()
        max_depth = dlg.get_depth()
        if self.run_batch_register(root_path, max_depth, base_categories=category_path):
            self.set_category_folder_path(category_path, root_path)
            self.refresh_category_tree()

    def on_edit_selected_folder(self):
        idx = self.selected_folder_index()
        if idx < 0:
            self.warn("フォルダを選択してください。")
            return
        it = self.folders_table.item(idx, 0)
        if not it:
            self.warn("登録情報が見つかりません。")
            return
        path = it.data(Qt.UserRole)
        if not path:
            self.warn("登録情報が見つかりません。")
            return
        self.edit_registered_folder(path)

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
        data = it.data(Qt.UserRole)
        if isinstance(data, dict) and data.get("type") == "category":
            path = data.get("path", [])
            if not isinstance(path, list):
                return
            self.current_folder = None
            self.current_meta = None
            self.current_file_rows = []
            self.files_table.setRowCount(0)
            tree_item = self.find_category_tree_item(path)
            if tree_item:
                self.category_tree.setCurrentItem(tree_item)
                tree_item.setExpanded(True)
            return

        path = data
        name = it.text()
        self.current_folder = {"name": name, "path": path}
        self.refresh_files_table()

    def on_folder_double_clicked(self, item: QTableWidgetItem):
        if item.column() != 0:
            return
        data = item.data(Qt.UserRole)
        if isinstance(data, dict) and data.get("type") == "category":
            path = data.get("path", [])
            if not isinstance(path, list):
                return
            tree_item = self.find_category_tree_item(path)
            if tree_item:
                self.category_tree.setCurrentItem(tree_item)
                tree_item.setExpanded(True)
            return
        path = data
        if path and os.path.isdir(path):
            self.open_in_explorer(path)

    def on_category_tree_selected(self):
        items = self.category_tree.selectedItems()
        if not items:
            return
        data = items[0].data(0, Qt.UserRole) or {}
        item_type = data.get("type")
        if item_type == "category":
            self.selected_category_path = data.get("path", [])
            self.refresh_folder_table()
        elif item_type == "folder":
            self.selected_category_path = data.get("category_path", [])
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

    def on_history_item_double_clicked(self, item: QTableWidgetItem):
        row = item.row()
        data_item = self.hist_table.item(row, 0)
        if not data_item:
            return
        entry = data_item.data(Qt.UserRole)
        if not isinstance(entry, dict):
            return
        dlg = HistoryDetailDialog(entry, self)
        dlg.exec()

    def on_rescan(self):
        self.update_new_folder_highlights()
        self.refresh_folder_table()
        self.refresh_files_table()
        self.refresh_category_tree()

    def startup_rescan(self):
        for item in self.registry:
            path = item.get("path", "")
            if path and os.path.isdir(path):
                scan_folder(path, self.ignore_types)
        self.update_new_folder_highlights()
        self.refresh_folder_table()
        self.refresh_category_tree()

    def on_options(self):
        dlg = OptionsDialog(self.memo_timeout_min, self.ignore_types, self.version_rules, self)
        if dlg.exec() != QDialog.Accepted:
            return
        self.memo_timeout_min = dlg.get_timeout_min()
        self.ignore_types = normalize_ignore_types(dlg.get_ignore_types())
        self.version_rules = normalize_version_rules(dlg.get_version_rules())
        self.settings["memo_timeout_min"] = self.memo_timeout_min
        self.settings["ignore_types"] = self.ignore_types
        self.settings["version_rules"] = self.version_rules
        save_settings(self.settings)
        self.refresh_folder_table()
        self.refresh_files_table()
        self.info("設定を保存しました。")

    def on_archive(self):
        archived = self.archived_categories()
        if not archived:
            self.info("アーカイブはありません。")
            return
        entries = []
        for path in archived:
            count = sum(
                1 for item in self.registry
                if self.category_path_for_item(item)[:len(path)] == path
            )
            entries.append({
                "path": path,
                "label": " / ".join(path),
                "count": count,
            })
        dlg = ArchiveDialog(entries, self)
        if dlg.exec() != QDialog.Accepted:
            return

    def on_cache_clear(self):
        if not self.ask("未使用のキャッシュデータを削除しますか？"):
            return
        settings_removed, user_checks_removed = self.clear_unused_cache()
        self.refresh_folder_table()
        self.refresh_category_tree()
        self.refresh_files_table()
        self.info(f"キャッシュクリアが完了しました。settings: {settings_removed}件, user_checks: {user_checks_removed}件")

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

        version_dialog = VersionSelectDialog(cur_rev, self.version_rules, self)
        if version_dialog.exec() != QDialog.Accepted:
            return

        _base_name, ext = split_name_ext(cur_fn)
        doc_base, _ver_tuple, _rev_str = parse_rev_from_filename(cur_fn)
        # doc_base includes ext; we want base without ext for naming
        doc_base_no_ext, _ = split_name_ext(doc_base)

        A, B, C = version_dialog.selected_next_version()
        new_rev = format_rev(A, B, C, today_yyyymmdd())
        new_fn = f"{doc_base_no_ext}_{new_rev}{ext}"

        history_dir = ensure_history_dir(folder_path)

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
            self.mark_doc_checked(folder_path, doc_key)

            # Refresh
            self.refresh_files_table()
            self.refresh_folder_table()
            self.refresh_category_tree()

            # Open the new file for convenience
            try:
                os.startfile(new_path)  # type: ignore[attr-defined]
            except Exception:
                pass
            self.start_file_lock_watch(new_path, folder_path, doc_key)

        except Exception as e:
            self.warn(f"更新に失敗しました: {e}")

    def on_view(self):
        sel = self._get_selected_doc()
        if not sel:
            return
        folder_path, doc_key, info = sel
        cur_fn = info.get("current_file", "")
        if not cur_fn:
            self.warn("現行ファイルが不明です。")
            return
        self.open_current_file(folder_path, cur_fn, doc_key)

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

        history_dir = ensure_history_dir(folder_path)

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
            self.mark_doc_checked(folder_path, doc_key)

            # Refresh
            self.refresh_files_table()
            self.refresh_folder_table()
            self.refresh_category_tree()

        except Exception as e:
            self.warn(f"差し替えに失敗しました: {e}")

    def on_rollback(self):
        sel = self._get_selected_doc()
        if not sel:
            return
        folder_path, doc_key, info = sel

        history_items = info.get("history", [])
        if not isinstance(history_items, list) or not history_items:
            self.warn("差し戻し対象の履歴がありません。")
            return

        dlg = HistorySelectDialog(history_items, self)
        if dlg.exec() != QDialog.Accepted:
            return
        history_entry = dlg.selected_item()
        if not history_entry:
            self.warn("差し戻し対象の履歴を選択してください。")
            return

        cur_fn = info.get("current_file", "")
        cur_rev = info.get("current_rev", "")
        if not cur_fn:
            self.warn("現行ファイルが不明です。")
            return

        history_file = history_entry.get("file", "")
        if not history_file:
            self.warn("履歴ファイル名が取得できません。")
            return

        history_dir = ensure_history_dir(folder_path)
        history_path = os.path.join(history_dir, history_file)
        if not os.path.exists(history_path):
            self.warn("履歴ファイルが見つかりません。")
            return

        doc_base, _t, _r = parse_rev_from_filename(cur_fn)
        doc_base_no_ext, ext = split_name_ext(doc_base)

        A, B, C = next_rev_from_current(cur_rev)
        new_rev = format_rev(A, B, C, today_yyyymmdd())
        new_fn = f"{doc_base_no_ext}_{new_rev}{ext}"
        new_path = os.path.join(folder_path, new_fn)
        cur_path = os.path.join(folder_path, cur_fn)

        if os.path.exists(new_path):
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

            # 2) copy selected history -> new current
            shutil.copy2(history_path, new_path)

            # 3) update meta
            meta = load_meta(folder_path)
            docs = meta.get("documents", {})
            d = docs.get(doc_key, {
                "title": doc_key,
                "current_file": new_fn,
                "current_rev": new_rev,
                "updated_at": now_iso(),
                "updated_by": user_name(),
                "last_memo": "",
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

            d["current_file"] = new_fn
            d["current_rev"] = new_rev
            d["updated_at"] = now_iso()
            d["updated_by"] = user_name()
            rollback_rev = history_entry.get("rev", "")
            d["last_memo"] = f"差し戻し: {display_rev(rollback_rev)}"

            docs[doc_key] = d
            meta["documents"] = docs
            save_meta(folder_path, meta)
            self.mark_doc_checked(folder_path, doc_key)

            self.refresh_files_table()
            self.refresh_folder_table()
            self.refresh_category_tree()
            self.refresh_right_pane_for_doc(doc_key)

        except Exception as e:
            self.warn(f"差し戻しに失敗しました: {e}")

    def on_history_clear(self):
        sel = self._get_selected_doc()
        if not sel:
            return
        folder_path, doc_key, info = sel
        history_items = info.get("history", [])
        if not isinstance(history_items, list) or not history_items:
            self.warn("削除できる履歴がありません。")
            return
        history_dir = ensure_history_dir(folder_path)
        history_items = [
            item for item in history_items
            if isinstance(item, dict) and item.get("file")
            and os.path.exists(os.path.join(history_dir, item.get("file", "")))
        ]
        if not history_items:
            self.warn("削除できる履歴がありません。")
            return

        dlg = HistoryClearDialog(history_items, info.get("current_rev", ""), self)
        if dlg.exec() != QDialog.Accepted:
            return
        selected_items = dlg.selected_items()
        if not selected_items:
            self.warn("削除対象が選択されていません。")
            return

        deleted_files = set()
        errors = []
        for item in selected_items:
            fname = item.get("file", "")
            if not fname:
                continue
            fpath = os.path.join(history_dir, fname)
            try:
                if os.path.exists(fpath):
                    os.remove(fpath)
                deleted_files.add(fname)
            except Exception as e:
                errors.append(f"{fname}: {e}")

        if deleted_files:
            self.refresh_right_pane_for_doc(doc_key)

        if errors:
            self.warn("削除時に一部エラーが発生しました:\n" + "\n".join(errors))

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
    icon = load_app_icon()
    if icon:
        app.setWindowIcon(icon)
    w = MainWindow()
    if icon:
        w.setWindowIcon(icon)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
