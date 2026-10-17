"""Directory Indexer — a PySide6 desktop app.

Features:
    - Pick any directory (internal or external drive)
    - Recursively scan all files and subdirectories
    - Two columns: Filename + Directory (full path, without the filename)
    - Three filter boxes (AND logic, case-insensitive)
    - "Find Files Not Matching MP4" window with its own filter and Excel export
    - "Flatten" — move every file from all subfolders into the top-level folder
    - "Group Files by Name" — for each file, move it into a subfolder named
      after its base name (files sharing a base name go into the same folder)
    - Export the currently visible rows to .xlsx

Dependencies:
    pip install PySide6 XlsxWriter
"""

import os
import shutil
import sys
from collections import defaultdict

import xlsxwriter
from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QSortFilterProxyModel,
    QThread,
    QTimer,
    Qt,
    Signal,
)
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTableView,
    QVBoxLayout,
    QWidget,
)


# ---------------------------------------------------------------------------
#  Scanner thread
# ---------------------------------------------------------------------------
class Scanner(QThread):
    progress = Signal(int)
    finished_scan = Signal(list)
    failed = Signal(str)

    def __init__(self, root: str):
        super().__init__()
        # Normalize once so every derived path uses the OS separator.
        self.root = os.path.normpath(root)
        self._abort = False

    def abort(self) -> None:
        self._abort = True

    def run(self) -> None:
        rows = []
        count = 0
        try:
            stack = [self.root]
            while stack and not self._abort:
                current = stack.pop()
                try:
                    with os.scandir(current) as it:
                        for entry in it:
                            if self._abort:
                                break
                            try:
                                if entry.is_dir(follow_symlinks=False):
                                    stack.append(entry.path)
                                elif entry.is_file(follow_symlinks=False):
                                    rows.append(
                                        (entry.name, os.path.dirname(entry.path))
                                    )
                                    count += 1
                                    if count % 500 == 0:
                                        self.progress.emit(count)
                            except OSError:
                                continue
                except OSError:
                    continue

            self.progress.emit(count)
            if not self._abort:
                self.finished_scan.emit(rows)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))


# ---------------------------------------------------------------------------
#  Flatten worker — moves every file from subfolders into the root
# ---------------------------------------------------------------------------
class FlattenWorker(QThread):
    progress = Signal(int, int)         # done, total
    finished_flatten = Signal(list)     # list of (src, dst, status)
    failed = Signal(str)

    def __init__(self, root: str, files: list[tuple[str, str]]):
        super().__init__()
        self.root = os.path.normpath(root)
        self.files = files
        self._abort = False

    def abort(self) -> None:
        self._abort = True

    def _unique_target(self, filename: str) -> str:
        base, ext = os.path.splitext(filename)
        candidate = os.path.join(self.root, filename)
        n = 1
        while os.path.exists(candidate):
            candidate = os.path.join(self.root, f"{base}_{n}{ext}")
            n += 1
        return candidate

    def run(self) -> None:
        try:
            to_move = [
                (name, d) for name, d in self.files
                if os.path.normpath(d) != self.root
            ]
            total = len(to_move)
            results: list[tuple[str, str, str]] = []
            for i, (name, d) in enumerate(to_move, 1):
                if self._abort:
                    break
                src = os.path.join(d, name)
                try:
                    target = self._unique_target(name)
                    shutil.move(src, target)
                    renamed = os.path.basename(target) != name
                    results.append((src, target, "renamed" if renamed else "moved"))
                except Exception as exc:  # noqa: BLE001
                    results.append((src, "", f"error: {exc}"))
                if i % 50 == 0 or i == total:
                    self.progress.emit(i, total)
            self.finished_flatten.emit(results)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))


# ---------------------------------------------------------------------------
#  Group-by-name worker
# ---------------------------------------------------------------------------
class GroupWorker(QThread):
    """Groups files sharing a base name (case-insensitive) into subfolders.

    Every file that is not already inside a folder matching its base name is
    moved into ``<parent>/<BaseName>/``. Multiple files with the same base
    name in the same folder all go into the same subfolder. Single-file
    groups get their own folder too.

    Files already living inside a folder named after their base name are
    skipped, so re-running this operation is idempotent.
    """

    progress = Signal(int, int)         # done, total
    finished_group = Signal(list)       # list of (src, dst, status)
    failed = Signal(str)

    def __init__(self, files: list[tuple[str, str]]):
        super().__init__()
        self.files = files
        self._abort = False

    def abort(self) -> None:
        self._abort = True

    @staticmethod
    def _unique_target(folder: str, filename: str) -> str:
        base, ext = os.path.splitext(filename)
        candidate = os.path.join(folder, filename)
        n = 1
        while os.path.exists(candidate):
            candidate = os.path.join(folder, f"{base}_{n}{ext}")
            n += 1
        return candidate

    def run(self) -> None:
        try:
            # Group files by (directory, lowercased base name).
            groups: dict[tuple[str, str], list[str]] = defaultdict(list)
            for name, d in self.files:
                stem = os.path.splitext(name)[0].lower()
                groups[(os.path.normpath(d), stem)].append(name)

            # Every group is processed — including single-file groups.
            # Only skip groups already inside a folder of the same name.
            to_process: list[tuple[str, str, list[str]]] = []
            for (d, stem_lower), names in groups.items():
                if os.path.basename(d).lower() == stem_lower:
                    continue
                to_process.append((d, stem_lower, names))

            total = sum(len(names) for _, _, names in to_process)
            results: list[tuple[str, str, str]] = []
            done = 0

            for d, _stem_lower, names in to_process:
                if self._abort:
                    break

                # Folder name uses the original casing of the first file.
                folder_name = os.path.splitext(names[0])[0]
                target_folder = os.path.join(d, folder_name)

                try:
                    os.makedirs(target_folder, exist_ok=True)
                except Exception as exc:  # noqa: BLE001
                    for f in names:
                        results.append(
                            (os.path.join(d, f), "", f"error: {exc}")
                        )
                        done += 1
                    self.progress.emit(done, total)
                    continue

                for f in names:
                    if self._abort:
                        break
                    src = os.path.join(d, f)
                    try:
                        target = self._unique_target(target_folder, f)
                        shutil.move(src, target)
                        status = "moved"
                        if os.path.basename(target) != f:
                            status = "renamed"
                        results.append((src, target, status))
                    except Exception as exc:  # noqa: BLE001
                        results.append((src, "", f"error: {exc}"))
                    done += 1
                    if done % 50 == 0 or done == total:
                        self.progress.emit(done, total)

            self.finished_group.emit(results)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))


# ---------------------------------------------------------------------------
#  Main table model
# ---------------------------------------------------------------------------
class ResultsModel(QAbstractTableModel):
    HEADERS = ("Filename", "Directory")

    def __init__(self) -> None:
        super().__init__()
        self.rows: list[tuple[str, str]] = []

    def set_rows(self, rows: list[tuple[str, str]]) -> None:
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else 2

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or role != Qt.DisplayRole:
            return None
        return self.rows[index.row()][index.column()]

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self.HEADERS[section]
        return None


# ---------------------------------------------------------------------------
#  Main table filter proxy
# ---------------------------------------------------------------------------
class FilterProxy(QSortFilterProxyModel):
    def __init__(self) -> None:
        super().__init__()
        self.keywords: list[str] = ["", "", ""]

    def set_keywords(self, keywords: list[str]) -> None:
        self.keywords = [k.strip().lower() for k in keywords]
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent) -> bool:
        model = self.sourceModel()
        name = model.data(model.index(source_row, 0, source_parent), Qt.DisplayRole) or ""
        directory = model.data(model.index(source_row, 1, source_parent), Qt.DisplayRole) or ""
        haystack = f"{name}\n{directory}".lower()
        return all(kw in haystack for kw in self.keywords if kw)


# ---------------------------------------------------------------------------
#  Excel helper (used by both windows)
# ---------------------------------------------------------------------------
def write_xlsx(path: str, headers: tuple[str, ...], rows: list[tuple]) -> None:
    workbook = xlsxwriter.Workbook(path, {"constant_memory": True})
    sheet = workbook.add_worksheet()
    for col, h in enumerate(headers):
        sheet.write(0, col, h)
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            sheet.write_string(r, c, str(val))
    workbook.close()


# ---------------------------------------------------------------------------
#  MP4 mismatch analysis
# ---------------------------------------------------------------------------
def find_mp4_mismatches(
    rows: list[tuple[str, str]]
) -> list[tuple[str, str, str]]:
    """For every folder that contains at least one .mp4 file, return every
    *other* file in that folder whose base name does not match any .mp4 base
    name in the same folder. Comparison is case-insensitive.
    """
    by_dir: dict[str, list[str]] = defaultdict(list)
    for name, directory in rows:
        by_dir[directory].append(name)

    results: list[tuple[str, str, str]] = []
    for directory, files in by_dir.items():
        mp4s = [f for f in files if f.lower().endswith(".mp4")]
        if not mp4s:
            continue
        mp4_stems = {os.path.splitext(f)[0].lower() for f in mp4s}
        mp4_label = " | ".join(mp4s)

        for f in files:
            if f.lower().endswith(".mp4"):
                continue
            stem = os.path.splitext(f)[0].lower()
            if stem not in mp4_stems:
                results.append((directory, f, mp4_label))

    return results


class MismatchModel(QAbstractTableModel):
    HEADERS = ("Folder", "Mismatched file", "MP4 file")

    def __init__(self, rows: list[tuple[str, str, str]]) -> None:
        super().__init__()
        self.rows = rows

    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else 3

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or role != Qt.DisplayRole:
            return None
        return self.rows[index.row()][index.column()]

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self.HEADERS[section]
        return None


class MismatchWindow(QMainWindow):
    def __init__(self, all_rows: list[tuple[str, str]]) -> None:
        super().__init__()
        self.setWindowTitle("Files Not Matching Their MP4")
        self.resize(1200, 700)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        info = QLabel(
            "Every folder that contains an .mp4 file, listing each other file "
            "in that folder whose name does not match the .mp4's name "
            "(case-insensitive). Subtitle, cover, and other companion files "
            "that share the movie's base name are not shown."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Filter mismatches…")
        self.filter_edit.textChanged.connect(self._on_filter)
        layout.addWidget(self.filter_edit)

        self.data = find_mp4_mismatches(all_rows)
        self.model = MismatchModel(self.data)
        self.proxy = QSortFilterProxyModel()
        self.proxy.setSourceModel(self.model)
        self.proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.proxy.setFilterKeyColumn(-1)  # search all columns

        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setAlternatingRowColors(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.setColumnWidth(0, 500)
        self.table.setColumnWidth(2, 260)
        layout.addWidget(self.table, 1)

        bottom = QHBoxLayout()
        self.lbl_count = QLabel(f"{len(self.data)} mismatches")
        btn_export = QPushButton("Export to Excel…")
        btn_export.clicked.connect(self._export)
        btn_export.setEnabled(bool(self.data))
        bottom.addWidget(self.lbl_count)
        bottom.addStretch(1)
        bottom.addWidget(btn_export)
        layout.addLayout(bottom)

    def _on_filter(self, text: str) -> None:
        self.proxy.setFilterFixedString(text)
        self.lbl_count.setText(
            f"{self.proxy.rowCount()} of {len(self.data)} mismatches"
        )

    def _export(self) -> None:
        if self.proxy.rowCount() == 0:
            QMessageBox.information(self, "Nothing to export", "There are no rows.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel file", "mp4_mismatches.xlsx", "Excel Workbook (*.xlsx)"
        )
        if not path:
            return
        try:
            visible: list[tuple[str, str, str]] = []
            for r in range(self.proxy.rowCount()):
                src_row = self.proxy.mapToSource(self.proxy.index(r, 0)).row()
                visible.append(self.data[src_row])
            write_xlsx(path, MismatchModel.HEADERS, visible)
            QMessageBox.information(
                self, "Export complete", f"Saved {len(visible)} rows to:\n{path}"
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Export error", str(exc))


# ---------------------------------------------------------------------------
#  Main window
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Directory Indexer")
        self.resize(1200, 700)
        self.scanner: Scanner | None = None
        self.flatten_worker: FlattenWorker | None = None
        self.group_worker: GroupWorker | None = None
        self.mismatch_window: MismatchWindow | None = None

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        top = QHBoxLayout()
        self.btn_select = QPushButton("Select Directory…")
        self.btn_select.clicked.connect(self.select_directory)
        self.lbl_dir = QLabel("No directory selected")
        top.addWidget(self.btn_select)
        top.addWidget(self.lbl_dir, 1)
        root.addLayout(top)

        filters = QHBoxLayout()
        self.f1, self.f2, self.f3 = QLineEdit(), QLineEdit(), QLineEdit()
        for i, field in enumerate((self.f1, self.f2, self.f3), start=1):
            field.setPlaceholderText(f"keyword {i}")
            field.textChanged.connect(self._schedule_filter)
            filters.addWidget(QLabel(f"Filter {i}:"))
            filters.addWidget(field)
        root.addLayout(filters)

        self._filter_timer = QTimer(self)
        self._filter_timer.setSingleShot(True)
        self._filter_timer.setInterval(250)
        self._filter_timer.timeout.connect(self.apply_filters)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        self.model = ResultsModel()
        self.proxy = FilterProxy()
        self.proxy.setSourceModel(self.model)
        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 320)
        root.addWidget(self.table, 1)

        bottom = QHBoxLayout()
        self.lbl_count = QLabel("0 files")
        self.btn_cancel = QPushButton("Cancel Scan")
        self.btn_cancel.setVisible(False)
        self.btn_cancel.clicked.connect(self.cancel_scan)
        self.btn_flatten = QPushButton("Flatten…")
        self.btn_flatten.setToolTip("Move every file from all subfolders into the top-level folder")
        self.btn_flatten.setEnabled(False)
        self.btn_flatten.clicked.connect(self.flatten_files)
        self.btn_group = QPushButton("Group Files by Name…")
        self.btn_group.setToolTip(
            "Move each file into a subfolder named after its base name "
            "(files sharing a base name go into the same folder)"
        )
        self.btn_group.setEnabled(False)
        self.btn_group.clicked.connect(self.group_files)
        self.btn_mismatch = QPushButton("MP4 Mismatches…")
        self.btn_mismatch.setToolTip("Find files in mp4 folders that don't match the mp4 name")
        self.btn_mismatch.setEnabled(False)
        self.btn_mismatch.clicked.connect(self.show_mismatches)
        self.btn_export = QPushButton("Export to Excel…")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self.export)
        bottom.addWidget(self.lbl_count)
        bottom.addStretch(1)
        bottom.addWidget(self.btn_cancel)
        bottom.addWidget(self.btn_flatten)
        bottom.addWidget(self.btn_group)
        bottom.addWidget(self.btn_mismatch)
        bottom.addWidget(self.btn_export)
        root.addLayout(bottom)

    # -- Filters ----------------------------------------------------------
    def _schedule_filter(self) -> None:
        self._filter_timer.start()

    def apply_filters(self) -> None:
        self.proxy.set_keywords([self.f1.text(), self.f2.text(), self.f3.text()])
        self.refresh_count()

    def refresh_count(self) -> None:
        total = self.model.rowCount()
        shown = self.proxy.rowCount()
        if total == shown:
            self.lbl_count.setText(f"{shown} files")
        else:
            self.lbl_count.setText(f"{shown} of {total} files")

    # -- Scan lifecycle ---------------------------------------------------
    def select_directory(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            self.start_scan(directory)

    def start_scan(self, directory: str) -> None:
        if self.mismatch_window is not None:
            self.mismatch_window.close()
            self.mismatch_window = None

        self.lbl_dir.setText(directory)
        self._set_actions_enabled(False)
        self.btn_cancel.setVisible(True)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.progress.setFormat("Scanning…")
        self.model.set_rows([])
        self.refresh_count()

        self.scanner = Scanner(directory)
        self.scanner.progress.connect(self.on_progress)
        self.scanner.finished_scan.connect(self.on_done)
        self.scanner.failed.connect(self.on_failed)
        self.scanner.start()

    def _set_actions_enabled(self, enabled: bool) -> None:
        self.btn_select.setEnabled(enabled)
        has_rows = enabled and bool(self.model.rows)
        self.btn_export.setEnabled(has_rows)
        self.btn_mismatch.setEnabled(has_rows)
        self.btn_flatten.setEnabled(has_rows)
        self.btn_group.setEnabled(has_rows)

    def on_progress(self, count: int) -> None:
        self.progress.setFormat(f"Scanning… {count} files found")

    def on_done(self, rows: list[tuple[str, str]]) -> None:
        self.model.set_rows(rows)
        self.refresh_count()
        self.progress.setVisible(False)
        self.btn_select.setEnabled(True)
        self.btn_cancel.setVisible(False)
        self._set_actions_enabled(True)

    def on_failed(self, message: str) -> None:
        self.progress.setVisible(False)
        self.btn_select.setEnabled(True)
        self.btn_cancel.setVisible(False)
        self._set_actions_enabled(True)
        QMessageBox.critical(self, "Scan error", message)

    def cancel_scan(self) -> None:
        if self.scanner:
            self.scanner.abort()

    # -- Flatten ----------------------------------------------------------
    def flatten_files(self) -> None:
        if not self.model.rows:
            return
        root = self.lbl_dir.text()
        root_norm = os.path.normpath(root)
        count = sum(
            1 for _, d in self.model.rows
            if os.path.normpath(d) != root_norm
        )
        if count == 0:
            QMessageBox.information(
                self, "Nothing to move",
                "All files are already in the top-level folder."
            )
            return

        reply = QMessageBox.question(
            self,
            "Move all files to top-level folder?",
            f"This will MOVE {count} file(s) from every subfolder into:\n\n"
            f"{root}\n\n"
            "• Empty subfolders are left in place (not deleted).\n"
            "• Files with duplicate names will be auto-renamed (e.g. Movie_1.mp4).\n"
            "• This cannot be undone from within the app.\n\n"
            "Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        self._set_actions_enabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, count)
        self.progress.setValue(0)
        self.progress.setFormat(f"Moving files… 0/{count}")

        self.flatten_worker = FlattenWorker(root, list(self.model.rows))
        self.flatten_worker.progress.connect(self.on_flatten_progress)
        self.flatten_worker.finished_flatten.connect(self.on_flatten_done)
        self.flatten_worker.failed.connect(self.on_flatten_failed)
        self.flatten_worker.start()

    def on_flatten_progress(self, done: int, total: int) -> None:
        self.progress.setRange(0, total)
        self.progress.setValue(done)
        self.progress.setFormat(f"Moving files… {done}/{total}")

    def on_flatten_done(self, results: list[tuple[str, str, str]]) -> None:
        self.progress.setVisible(False)
        self.btn_select.setEnabled(True)

        moved = sum(1 for _, _, s in results if s == "moved")
        renamed = sum(1 for _, _, s in results if s == "renamed")
        errors = [(src, status) for src, _, status in results if status.startswith("error")]

        summary = (
            f"Moved: {moved}\n"
            f"Renamed (name conflict): {renamed}\n"
            f"Errors: {len(errors)}"
        )
        if errors:
            shown = "\n".join(
                f"  • {os.path.basename(src)}: {err}" for src, err in errors[:15]
            )
            if len(errors) > 15:
                shown += f"\n  …and {len(errors) - 15} more"
            summary += "\n\n" + shown

        QMessageBox.information(self, "Flatten complete", summary)
        self.start_scan(self.lbl_dir.text())

    def on_flatten_failed(self, message: str) -> None:
        self.progress.setVisible(False)
        self._set_actions_enabled(True)
        QMessageBox.critical(self, "Flatten error", message)

    # -- Group files by base name ----------------------------------------
    def _preview_groups(self) -> tuple[int, int, list[str]]:
        """Return (num_folders, num_files, sample_lines) for the current scan."""
        groups: dict[tuple[str, str], list[str]] = defaultdict(list)
        for name, d in self.model.rows:
            stem = os.path.splitext(name)[0].lower()
            groups[(os.path.normpath(d), stem)].append(name)

        count_folders = 0
        count_files = 0
        samples: list[str] = []
        for (d, stem_lower), names in groups.items():
            if os.path.basename(d).lower() == stem_lower:
                continue
            count_folders += 1
            count_files += len(names)
            if len(samples) < 5:
                samples.append(
                    f"  • {os.path.splitext(names[0])[0]}  ({len(names)} file"
                    f"{'s' if len(names) != 1 else ''})"
                )
        return count_folders, count_files, samples

    def group_files(self) -> None:
        if not self.model.rows:
            return

        num_folders, num_files, samples = self._preview_groups()
        if num_files == 0:
            QMessageBox.information(
                self, "Nothing to group",
                "Every file is already inside a folder named after its base name."
            )
            return

        sample_text = "\n".join(samples)
        if num_folders > len(samples):
            sample_text += f"\n  …and {num_folders - len(samples)} more folders"

        reply = QMessageBox.question(
            self,
            "Group files into subfolders by name?",
            f"This will create {num_folders} subfolder(s) and MOVE "
            f"{num_files} file(s) into them.\n\n"
            f"Examples:\n{sample_text}\n\n"
            "• Files sharing a base name (case-insensitive) go into the same folder.\n"
            "• A file on its own still gets its own folder.\n"
            "• Files already inside a matching folder are skipped.\n"
            "• Name conflicts inside a new folder get auto-suffixed "
            "(e.g. Movie_1.mp4).\n"
            "• This cannot be undone from within the app.\n\n"
            "Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        self._set_actions_enabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, num_files)
        self.progress.setValue(0)
        self.progress.setFormat(f"Grouping files… 0/{num_files}")

        self.group_worker = GroupWorker(list(self.model.rows))
        self.group_worker.progress.connect(self.on_group_progress)
        self.group_worker.finished_group.connect(self.on_group_done)
        self.group_worker.failed.connect(self.on_group_failed)
        self.group_worker.start()

    def on_group_progress(self, done: int, total: int) -> None:
        self.progress.setRange(0, total)
        self.progress.setValue(done)
        self.progress.setFormat(f"Grouping files… {done}/{total}")

    def on_group_done(self, results: list[tuple[str, str, str]]) -> None:
        self.progress.setVisible(False)
        self.btn_select.setEnabled(True)

        moved = sum(1 for _, _, s in results if s == "moved")
        renamed = sum(1 for _, _, s in results if s == "renamed")
        errors = [(src, status) for src, _, status in results if status.startswith("error")]

        summary = (
            f"Moved into subfolders: {moved}\n"
            f"Renamed (name conflict): {renamed}\n"
            f"Errors: {len(errors)}"
        )
        if errors:
            shown = "\n".join(
                f"  • {os.path.basename(src)}: {err}" for src, err in errors[:15]
            )
            if len(errors) > 15:
                shown += f"\n  …and {len(errors) - 15} more"
            summary += "\n\n" + shown

        QMessageBox.information(self, "Group complete", summary)
        self.start_scan(self.lbl_dir.text())

    def on_group_failed(self, message: str) -> None:
        self.progress.setVisible(False)
        self._set_actions_enabled(True)
        QMessageBox.critical(self, "Group error", message)

    # -- MP4 mismatch window ---------------------------------------------
    def show_mismatches(self) -> None:
        if not self.model.rows:
            return
        if self.mismatch_window is not None:
            self.mismatch_window.close()
        self.mismatch_window = MismatchWindow(self.model.rows)
        self.mismatch_window.show()

    # -- Main export ------------------------------------------------------
    def export(self) -> None:
        if self.proxy.rowCount() == 0:
            QMessageBox.information(self, "Nothing to export", "There are no rows.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel file", "directory_index.xlsx", "Excel Workbook (*.xlsx)"
        )
        if not path:
            return
        try:
            rows = self.filtered_rows()
            write_xlsx(path, ResultsModel.HEADERS, rows)
            QMessageBox.information(
                self, "Export complete", f"Saved {len(rows)} rows to:\n{path}"
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Export error", str(exc))

    def filtered_rows(self) -> list[tuple[str, str]]:
        if not any(kw.strip() for kw in self.proxy.keywords):
            return list(self.model.rows)
        rows = []
        for r in range(self.proxy.rowCount()):
            src = self.proxy.mapToSource(self.proxy.index(r, 0)).row()
            rows.append(self.model.rows[src])
        return rows

    # -- Window close -----------------------------------------------------
    def closeEvent(self, event) -> None:
        if self.scanner and self.scanner.isRunning():
            self.scanner.abort()
            self.scanner.wait(3000)
        if self.flatten_worker and self.flatten_worker.isRunning():
            self.flatten_worker.abort()
            self.flatten_worker.wait(3000)
        if self.group_worker and self.group_worker.isRunning():
            self.group_worker.abort()
            self.group_worker.wait(3000)
        if self.mismatch_window is not None:
            self.mismatch_window.close()
        super().closeEvent(event)


# ---------------------------------------------------------------------------
#  Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()