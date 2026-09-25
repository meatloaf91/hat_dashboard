from __future__ import annotations

import csv
from datetime import datetime
import re
import sys
import tempfile
import time
from pathlib import Path
import pandas as pd
from collections import defaultdict
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure

# ── Python 3.14 compatibility ─────────────────────────────────────────────────
# matplotlib's Path.__deepcopy__ calls copy.deepcopy(super(), memo) which
# triggers infinite recursion in Python ≥3.14 due to changed super() proxy
# deepcopy behaviour.  Replace with a direct dict-based deepcopy that avoids
# the super() proxy entirely.
def _patch_matplotlib_path_deepcopy() -> None:
    import sys as _sys
    if _sys.version_info < (3, 14):
        return
    import copy as _copy
    import matplotlib.path as _mp

    def _safe_deepcopy(self, memo):  # type: ignore[override]
        cls = type(self)
        result = cls.__new__(cls)
        memo[id(self)] = result
        for k, v in self.__dict__.items():
            object.__setattr__(result, k, _copy.deepcopy(v, memo))
        return result

    _mp.Path.__deepcopy__ = _safe_deepcopy  # type: ignore[method-assign]

_patch_matplotlib_path_deepcopy()
# ─────────────────────────────────────────────────────────────────────────────

_BASE_DIR = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))

from PySide6.QtCore import Qt, Property, QPropertyAnimation, QEasingCurve, QEvent, QRect, Signal, QTimer
from PySide6.QtGui import QColor, QCursor, QImage, QKeySequence, QMovie, QPainter, QPainterPath, QPixmap, QLinearGradient, QBrush, QPen, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QTextEdit,
    QDoubleSpinBox,
    QDialog,
    QColorDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSlider,
    QTabWidget,
    QToolButton,
    QMenu,
    QStackedWidget,
    QSpacerItem,
    QSizePolicy,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from status_collector import StatusCollector
from thumbnail_generator import ThumbnailGenerator
from packshot_naming_generator import PackshotNamingGenerator
from sap_reformat import SapTableReformatError, SapTableReformatter
from general_functions import clear_other_panel_inputs
from hat_config import HatConfig
from body_mapper import CompareParams, run_comparison, run_set_comparison
from reference_collector import RefCollectorParams, run_reference_collector
from artwork_processing import ArtworkInspection, ArtworkProcessor, ProcessOptions, inspect_artwork
from artwork_processing import ArtworkInspection, ArtworkProcessor, Bounds, CutSelection, ProcessOptions, inspect_artwork
import review_project
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


class ToggleSwitch(QWidget):
    """Animated pill toggle that mimics a native iOS/Android switch."""

    toggled = Signal(bool)
    def __init__(self, parent=None, width=46, height=26):
        super().__init__(parent)
        self._checked = False
        self._thumb_x = float(height // 2)        # centre-x of thumb
        self._w = width
        self._h = height
        self._r = height // 2                     # pill corner radius
        self._pad = 3                             # gap between thumb edge and pill edge
        self._thumb_r = self._r - self._pad       # thumb radius
        self._x_off = float(self._r)              # thumb centre when off
        self._x_on  = float(self._w - self._r)   # thumb centre when on
        self._thumb_x = self._x_off
        self.setFixedSize(self._w, self._h)
        self.setCursor(Qt.PointingHandCursor)

        self._anim = QPropertyAnimation(self, b"thumb_x", self)
        self._anim.setDuration(160)
        self._anim.setEasingCurve(QEasingCurve.Type.InOutQuad)

    # ── Qt property so QPropertyAnimation can drive it ───────────────
    def _get_thumb_x(self) -> float:
        return self._thumb_x

    def _set_thumb_x(self, val: float):
        self._thumb_x = val
        self.update()

    thumb_x = Property(float, _get_thumb_x, _set_thumb_x)

    # ── Public API ───────────────────────────────────────────────────
    def isChecked(self) -> bool:
        return self._checked

    def setChecked(self, state: bool):
        if state == self._checked:
            return
        self._checked = state
        self._animate_to(self._x_on if state else self._x_off)
        self.update()
        self.toggled.emit(state)

    def toggle(self):
        self.setChecked(not self._checked)

    # ── Interaction ──────────────────────────────────────────────────
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.toggle()

    # ── Animation ────────────────────────────────────────────────────
    def _animate_to(self, target: float):
        self._anim.stop()
        self._anim.setStartValue(self._thumb_x)
        self._anim.setEndValue(target)
        self._anim.start()

    # ── Painting ─────────────────────────────────────────────────────
    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        # ---- pill background ----------------------------------------
        if self._checked:
            # teal → blue-purple gradient (left to right)
            grad = QLinearGradient(0, 0, self._w, 0)
            grad.setColorAt(0.0, QColor("#5de0c8"))   # teal
            grad.setColorAt(1.0, QColor("#6b88e6"))   # blue-purple
            p.setBrush(QBrush(grad))
        else:
            p.setBrush(QBrush(QColor("#88929e")))

        p.setPen(Qt.NoPen)
        p.drawRoundedRect(0, 0, self._w, self._h, self._r, self._r)

        # ---- white thumb --------------------------------------------
        cx = int(self._thumb_x)
        cy = self._h // 2
        p.setBrush(QBrush(QColor("#ffffff")))
        # subtle drop shadow
        shadow_pen = QPen(QColor(0, 0, 0, 40))
        shadow_pen.setWidth(0)
        p.setPen(shadow_pen)
        p.drawEllipse(
            cx - self._thumb_r + 1,
            cy - self._thumb_r + 1,
            self._thumb_r * 2,
            self._thumb_r * 2,
        )
        p.setPen(Qt.NoPen)
        p.drawEllipse(
            cx - self._thumb_r,
            cy - self._thumb_r,
            self._thumb_r * 2,
            self._thumb_r * 2,
        )

        p.end()


class _PlainTextLineAdapter:
    """Compatibility wrapper so QLineEdit can be used by legacy textEdit-based logic."""

    def __init__(self, line_edit: QLineEdit) -> None:
        self._line_edit = line_edit

    def setPlainText(self, text: str) -> None:
        self._line_edit.setText(text)

    def toPlainText(self) -> str:
        return self._line_edit.text()

    def clear(self) -> None:
        self._line_edit.clear()

    def setEnabled(self, enabled: bool) -> None:
        self._line_edit.setEnabled(enabled)


class _ClipboardTableWidget(QTableWidget):
    def __init__(self, rows: int, cols: int, parent: QWidget | None = None) -> None:
        super().__init__(rows, cols, parent)
        self.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
            | QAbstractItemView.EditTrigger.AnyKeyPressed
        )
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

    def keyPressEvent(self, event) -> None:
        if event.matches(QKeySequence.StandardKey.Copy):
            self._copy_selection()
            return
        if event.matches(QKeySequence.StandardKey.Paste):
            self._paste_clipboard()
            return
        else:
            super().keyPressEvent(event)

    def _copy_selection(self) -> None:
        selected_ranges = self.selectedRanges()
        if not selected_ranges:
            return

        # Copy first selected rectangular range as tab/newline delimited text.
        selected_range = selected_ranges[0]
        lines = []
        for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
            row_values = []
            for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                cell = self.item(row, col)
                row_values.append(cell.text() if cell else "")
            lines.append("\t".join(row_values))

        QApplication.clipboard().setText("\n".join(lines))

    def _paste_clipboard(self) -> None:
        if getattr(self, '_paste_disabled', False):
            return
        clip_text = QApplication.clipboard().text()
        if not clip_text:
            return

        parent = self.parent()
        begin_batch = getattr(parent, "_on_table_batch_edit_begin", None)
        end_batch = getattr(parent, "_on_table_batch_edit_end", None)

        normalized = clip_text.replace("\r\n", "\n").replace("\r", "\n")
        raw_lines = normalized.split("\n")
        if len(raw_lines) > 1 and raw_lines[-1] == "":
            raw_lines = raw_lines[:-1]

        rows = [line.split("\t") for line in raw_lines]
        if not rows:
            return

        if callable(begin_batch):
            begin_batch()

        # If multiple cells are selected, apply clipboard values across all selected cells.
        # Single value => fill all selected cells.
        # Multiple values => cycle through selected cells in row/column order.
        selected_indexes = self.selectedIndexes()
        if len(selected_indexes) > 1:
            flat_values = [value for row_values in rows for value in row_values]
            if not flat_values:
                if callable(end_batch):
                    end_batch()
                return

            ordered = sorted(selected_indexes, key=lambda idx: (idx.row(), idx.column()))
            for i, index in enumerate(ordered):
                r = index.row()
                c = index.column()
                item = self.item(r, c)
                if item is None:
                    item = QTableWidgetItem("")
                    self.setItem(r, c, item)
                item.setText(flat_values[i % len(flat_values)])
            if callable(end_batch):
                end_batch()
            return

        selected_ranges = self.selectedRanges()
        if selected_ranges:
            selected_range = sorted(selected_ranges, key=lambda r: (r.topRow(), r.leftColumn()))[0]
            start_row = selected_range.topRow()
            start_col = selected_range.leftColumn()
        elif self.currentRow() >= 0 and self.currentColumn() >= 0:
            start_row = self.currentRow()
            start_col = self.currentColumn()
        else:
            start_row = 0
            start_col = 0

        required_rows = start_row + len(rows)
        if required_rows > self.rowCount():
            self.setRowCount(required_rows)

        for r_offset, values in enumerate(rows):
            for c_offset, value in enumerate(values):
                col = start_col + c_offset
                if col >= self.columnCount():
                    break
                row = start_row + r_offset
                item = self.item(row, col)
                if item is None:
                    item = QTableWidgetItem("")
                    self.setItem(row, col, item)
                item.setText(value)

        if callable(end_batch):
            end_batch()


class _PackshotClipboardTableDialog(QDialog):
    _BLANK_FILTER = "__BLANK__"

    def __init__(self, parent: QWidget | None = None, start_dir: str = "") -> None:
        super().__init__(parent)
        self._browse_start_dir = start_dir
        self.setWindowTitle("Paste on Table")
        self.resize(1160, 620)
        self.setMinimumSize(1080, 520)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        title_row = QHBoxLayout()
        title_row.setSpacing(8)
        title = QLabel("Item Information")
        title.setObjectName("packshotTableTitle")
        title_row.addWidget(title, 0)
        title_row.addSpacing(80)
        self.label_visible_count = QLabel("Count: <b>0</b>")
        self.label_visible_count.setObjectName("packshotRowCountLabel")
        self.label_visible_count.setTextFormat(Qt.TextFormat.RichText)
        title_row.addWidget(self.label_visible_count, 0)
        title_row.addStretch(1)
        layout.addLayout(title_row)

        config_row = QHBoxLayout()
        config_row.setSpacing(8)
        row_count_label = QLabel("Row count")
        row_count_label.setObjectName("packshotRowCountLabel")
        config_row.addWidget(row_count_label, 0)

        self.row_count_input = QLineEdit("5")
        self.row_count_input.setObjectName("packshotRowCountInput")
        self.row_count_input.setFixedWidth(90)
        config_row.addWidget(self.row_count_input, 0)

        self.btn_update_rows = QPushButton("Update")
        self.btn_update_rows.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_update_rows, 0)

        self.btn_reset_table = QPushButton("Reset")
        self.btn_reset_table.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_reset_table, 0)

        self.btn_delete_row = QPushButton("Delete")
        self.btn_delete_row.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_delete_row, 0)

        self.btn_undo_table = QPushButton("Undo")
        self.btn_undo_table.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_undo_table, 0)
        config_row.addSpacing(6)

        self.btn_reset_filter = QPushButton("Reset Filter")
        self.btn_reset_filter.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_reset_filter, 0)

        starting_row_label = QLabel("Starting Row")
        starting_row_label.setObjectName("packshotRowCountLabel")
        config_row.addWidget(starting_row_label, 0)
        self.starting_row_input = QLineEdit("1")
        self.starting_row_input.setObjectName("packshotRowCountInput")
        self.starting_row_input.setFixedWidth(90)
        config_row.addWidget(self.starting_row_input, 0)

        self.btn_import_from_tracker = QPushButton("Import")
        self.btn_import_from_tracker.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_import_from_tracker, 0)

        self.btn_generate_packshot_naming = QPushButton("Generate")
        self.btn_generate_packshot_naming.setObjectName("packshotPrimaryActionBtn")
        self.btn_generate_packshot_naming.setMinimumWidth(150)
        config_row.addWidget(self.btn_generate_packshot_naming, 0)

        self.btn_export_table = QPushButton("Export")
        self.btn_export_table.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_export_table, 0)

        self.btn_read_me = QPushButton("Read Me")
        self.btn_read_me.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_read_me, 0)
        config_row.addStretch(1)
        layout.addLayout(config_row)

        self.table = _ClipboardTableWidget(5, 6, self)
        self.table.setObjectName("packshotClipboardTable")
        self._set_headers()
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.horizontalHeader().setDefaultSectionSize(170)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(5, 300)
        self.table.horizontalHeader().setSectionsClickable(True)
        self.table.horizontalHeader().sectionClicked.connect(self._on_filter_header_clicked)

        # ── Copy-all button overlaid on the "Packshot Naming" header cell ────────
        _hdr = self.table.horizontalHeader()
        self._copy_packshot_btn = QPushButton("\u29c9", _hdr)  # ⧉ copy symbol
        self._copy_packshot_btn.setObjectName("packshotCopyHeaderBtn")
        self._copy_packshot_btn.setFixedSize(28, 24)
        self._copy_packshot_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._copy_packshot_btn.setToolTip("Copy all Packshot Naming values to clipboard")
        self._copy_packshot_btn.setStyleSheet(
            "QPushButton#packshotCopyHeaderBtn {"
            "  background: rgba(255,255,255,0.18);"
            "  color: #FFFFFF;"
            "  border: 1px solid rgba(255,255,255,0.50);"
            "  border-radius: 4px;"
            "  font-size: 14px;"
            "  padding: 0;"
            "}"
            "QPushButton#packshotCopyHeaderBtn:hover {"
            "  background: rgba(255,255,255,0.35);"
            "}"
            "QPushButton#packshotCopyHeaderBtn:pressed {"
            "  background: rgba(255,255,255,0.60);"
            "}"
        )
        self._copy_packshot_btn.clicked.connect(self._copy_packshot_naming_column)
        _hdr.sectionResized.connect(lambda *_: self._position_copy_btn())
        QTimer.singleShot(0, self._position_copy_btn)

        for row in range(5):
            for col in range(6):
                self.table.setItem(row, col, QTableWidgetItem(""))

        self._undo_stack: list[list[list[str]]] = []
        self._is_restoring_undo = False
        self._is_batch_edit = False

        layout.addWidget(self.table)
        self.btn_update_rows.clicked.connect(self._on_update_rows_clicked)
        self.btn_reset_table.clicked.connect(self._on_reset_table_clicked)
        self.btn_delete_row.clicked.connect(self._on_delete_rows_clicked)
        self.btn_undo_table.clicked.connect(self._on_undo_table_clicked)
        self.btn_reset_filter.clicked.connect(self._on_reset_filter_clicked)
        self.btn_read_me.clicked.connect(self._on_read_me_clicked)
        self.table.itemChanged.connect(self._on_table_item_changed)
        self._push_undo_snapshot(force=True)
        self._apply_filters()

        self.setStyleSheet(
            """
            QDialog {
                background-color: #F4F4F4;
            }

            QTableWidget#packshotClipboardTable {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                gridline-color: #B8B8B8;
                selection-background-color: #DCE6F5;
                selection-color: #111111;
                font-family: "Segoe UI";
                font-size: 12px;
            }

            QTableWidget#packshotClipboardTable QHeaderView::section {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #7D8694;
                padding: 6px 8px;
                font-weight: 700;
            }

            QTableWidget#packshotClipboardTable QTableCornerButton::section {
                background-color: #111F35;
                border: 1px solid #7D8694;
            }

            QLabel#packshotTableTitle {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 18px;
                font-weight: 800;
            }

            QLabel#packshotRowCountLabel {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 600;
            }

            QLineEdit#packshotRowCountInput {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 8px;
                font-family: "Segoe UI";
                font-size: 12px;
            }

            QProgressBar#sapImportProgress {
                background-color: #FFFFFF;
                border: 1px solid #A9A9A9;
                border-radius: 7px;
                text-align: center;
                color: #111111;
                font-family: "Segoe UI";
                font-size: 11px;
                min-height: 24px;
            }

            QProgressBar#sapImportProgress::chunk {
                background-color: #8A244B;
                border-radius: 6px;
            }

            QPushButton#packshotUpdateRowsBtn {
                background-color: #9EA3AB;
                color: #000000;
                border: 1px solid #8B9098;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }

            QPushButton#packshotUpdateRowsBtn:pressed {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #111F35;
            }

            QPushButton#packshotPrimaryActionBtn {
                background-color: #8A244B;
                color: #FFFFFF;
                border: 1px solid #8A244B;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 700;
            }

            QPushButton#packshotPrimaryActionBtn:pressed {
                background-color: #F63049;
                border: 1px solid #F63049;
            }
            """
        )

    def _set_headers(self) -> None:
        self._header_labels = [
            "Product Name",
            "IDH",
            "Packaging Type",
            "Packaging Size",
            "View",
            "Packshot Naming",
        ]
        self._column_filters: dict[int, set[str] | None] = {i: None for i in range(len(self._header_labels))}
        self._refresh_header_labels()

        last_header_item = self.table.horizontalHeaderItem(5)
        if last_header_item is not None:
            last_header_item.setBackground(Qt.GlobalColor.transparent)
            last_header_item.setBackground(QColor("#8A244B"))
            last_header_item.setForeground(QColor("#FFFFFF"))

    def _refresh_header_labels(self) -> None:
        labels = []
        for idx, base_label in enumerate(self._header_labels):
            has_active_filter = self._column_filters.get(idx) is not None
            labels.append(f"{base_label} {'▾*' if has_active_filter else '▾'}")
        self.table.setHorizontalHeaderLabels(labels)

        last_header_item = self.table.horizontalHeaderItem(5)
        if last_header_item is not None:
            last_header_item.setBackground(Qt.GlobalColor.transparent)
            last_header_item.setBackground(QColor("#8A244B"))
            last_header_item.setForeground(QColor("#FFFFFF"))

    def _position_copy_btn(self) -> None:
        """Pin the copy button to the right edge of the Packshot Naming header section."""
        hdr = self.table.horizontalHeader()
        col_x = hdr.sectionViewportPosition(5)
        col_w = hdr.sectionSize(5)
        btn_w = self._copy_packshot_btn.width()
        btn_h = self._copy_packshot_btn.height()
        x = col_x + col_w - btn_w - 6
        y = max(0, (hdr.height() - btn_h) // 2)
        self._copy_packshot_btn.move(x, y)
        self._copy_packshot_btn.show()

    def _copy_packshot_naming_column(self) -> None:
        """Copy all visible Packshot Naming cell values to clipboard, one value per line."""
        values: list[str] = []
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            item = self.table.item(row, 5)
            values.append(item.text() if item is not None else "")
        QApplication.clipboard().setText("\n".join(values))
        # Brief tick feedback so the user knows it worked
        self._copy_packshot_btn.setText("\u2713")  # ✓
        QTimer.singleShot(1500, lambda: self._copy_packshot_btn.setText("\u29c9"))

    def _on_filter_header_clicked(self, column: int) -> None:
        values: list[str] = []
        seen: set[str] = set()
        has_blank = False
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            item = self.table.item(row, column)
            cell_text = item.text().strip() if item is not None else ""
            if cell_text == "":
                has_blank = True
                continue
            if cell_text in seen:
                continue
            seen.add(cell_text)
            values.append(cell_text)

        value_pairs: list[tuple[str, str]] = [(value, value) for value in values]
        if has_blank:
            value_pairs.append((self._BLANK_FILTER, "(Blanks)"))

        popup = _MapperReformattedTableDialog._FilterPopup(value_pairs, self._column_filters.get(column), self)
        popup.move(QCursor.pos())
        if popup.exec() != QDialog.DialogCode.Accepted:
            return

        selected = popup.get_selected_values()
        if len(selected) == len(value_pairs):
            self._column_filters[column] = None
        else:
            self._column_filters[column] = selected
        self._apply_filters()

    def _apply_filters(self) -> None:
        for row in range(self.table.rowCount()):
            row_matches = True
            for col, filter_value in self._column_filters.items():
                if filter_value is None:
                    continue
                item = self.table.item(row, col)
                cell_text = item.text().strip() if item is not None else ""
                normalized = self._BLANK_FILTER if cell_text == "" else cell_text
                if normalized not in filter_value:
                    row_matches = False
                    break
            self.table.setRowHidden(row, not row_matches)

        self._refresh_header_labels()
        self._update_visible_count_label()

    def _update_visible_count_label(self) -> None:
        count = 0
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            count += 1
        self.label_visible_count.setText(f"Count: <b>{count}</b>")

    def _on_update_rows_clicked(self) -> None:
        raw_value = self.row_count_input.text().strip()
        try:
            new_count = int(raw_value)
        except ValueError:
            self.row_count_input.setText(str(self.table.rowCount()))
            return

        new_count = max(1, min(new_count, 5000))
        self.table.setRowCount(new_count)
        for row in range(new_count):
            for col in range(self.table.columnCount()):
                if self.table.item(row, col) is None:
                    self.table.setItem(row, col, QTableWidgetItem(""))
        self.row_count_input.setText(str(new_count))
        self._apply_filters()
        self._push_undo_snapshot()

    def _on_reset_table_clicked(self) -> None:
        self.row_count_input.setText("5")
        if hasattr(self, "starting_row_input"):
            self.starting_row_input.setText("1")
        self.table.clearContents()
        self.table.setRowCount(5)
        self.table.clearSelection()
        for row in range(5):
            for col in range(self.table.columnCount()):
                self.table.setItem(row, col, QTableWidgetItem(""))
        for column in self._column_filters:
            self._column_filters[column] = None
        self._apply_filters()
        self._push_undo_snapshot()

    def _on_delete_rows_clicked(self) -> None:
        selected_rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        if not selected_rows:
            current_row = self.table.currentRow()
            if current_row >= 0:
                selected_rows = [current_row]
            else:
                return

        total_rows = self.table.rowCount()
        if total_rows <= 0:
            return

        if len(selected_rows) >= total_rows:
            self.table.setRowCount(1)
            for col in range(self.table.columnCount()):
                self.table.setItem(0, col, QTableWidgetItem(""))
        else:
            for row in selected_rows:
                if 0 <= row < self.table.rowCount():
                    self.table.removeRow(row)

        self.table.clearSelection()
        self.row_count_input.setText(str(self.table.rowCount()))
        self._apply_filters()
        self._push_undo_snapshot()

    def _on_reset_filter_clicked(self) -> None:
        for column in self._column_filters:
            self._column_filters[column] = None
        self._apply_filters()

    def _capture_table_state(self) -> list[list[str]]:
        state: list[list[str]] = []
        for row in range(self.table.rowCount()):
            row_values: list[str] = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_values.append(item.text() if item is not None else "")
            state.append(row_values)
        return state

    def _restore_table_state(self, state: list[list[str]]) -> None:
        self._is_restoring_undo = True
        self.table.blockSignals(True)
        try:
            row_count = max(1, len(state))
            col_count = self.table.columnCount()
            self.table.setRowCount(row_count)

            for row in range(row_count):
                row_values = state[row] if row < len(state) else [""] * col_count
                for col in range(col_count):
                    value = row_values[col] if col < len(row_values) else ""
                    item = self.table.item(row, col)
                    if item is None:
                        item = QTableWidgetItem("")
                        self.table.setItem(row, col, item)
                    item.setText(value)
        finally:
            self.table.blockSignals(False)
            self._is_restoring_undo = False

        self.row_count_input.setText(str(self.table.rowCount()))
        self._apply_filters()

    def _push_undo_snapshot(self, force: bool = False) -> None:
        if self._is_restoring_undo or self._is_batch_edit:
            return

        snapshot = self._capture_table_state()
        if not force and self._undo_stack and self._undo_stack[-1] == snapshot:
            return

        self._undo_stack.append(snapshot)
        if len(self._undo_stack) > 200:
            self._undo_stack.pop(0)

    def _on_table_item_changed(self, _item: QTableWidgetItem) -> None:
        if self._is_batch_edit:
            return
        if any(value is not None for value in self._column_filters.values()):
            self._apply_filters()
        else:
            self._update_visible_count_label()
        self._push_undo_snapshot()

    def _on_table_batch_edit_begin(self) -> None:
        if self._is_restoring_undo:
            return
        self._push_undo_snapshot(force=False)
        self._is_batch_edit = True

    def _on_table_batch_edit_end(self) -> None:
        if self._is_restoring_undo:
            return
        self._is_batch_edit = False
        self._apply_filters()
        self._push_undo_snapshot(force=False)

    def _on_undo_table_clicked(self) -> None:
        if self._is_restoring_undo:
            return

        current_snapshot = self._capture_table_state()
        if not self._undo_stack:
            self._undo_stack.append(current_snapshot)
            return

        if self._undo_stack[-1] != current_snapshot:
            self._undo_stack.append(current_snapshot)

        if len(self._undo_stack) <= 1:
            return

        self._undo_stack.pop()
        previous_snapshot = self._undo_stack[-1]
        self._restore_table_state(previous_snapshot)

    def _on_read_me_clicked(self) -> None:
        msg = QMessageBox(self)
        msg.setWindowTitle("Read Me")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setStyleSheet(
            """
            QMessageBox {
                background-color: #F4F4F4;
            }
            QMessageBox QLabel {
                color: #111111;
                font-family: "Segoe UI";
                font-size: 12px;
            }
            QMessageBox QPushButton {
                background-color: #9EA3AB;
                color: #000000;
                border: 1px solid #8B9098;
                border-radius: 8px;
                min-height: 28px;
                min-width: 80px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }
            """
        )
        msg.setText(
            "Logic Info:\n"
            "1. User can copy paste data on the table similar to excel.\n"
            "2. When importing data, the following rows from the excel tracker will be ignored:\n"
            "-if IDH column is blank.\n"
            "-if Product Name and IDH column have value \"admin\"."
        )
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()


class _SingleTrackerColumnPickerDialog(QDialog):
    """Ask the user to choose one column from a list of column identifiers."""

    def __init__(self, columns: list[str], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Select Column")
        self.setMinimumWidth(360)
        self._selected: str | None = None

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        lbl = QLabel("Multiple columns detected.\nSelect the column to import:")
        lbl.setWordWrap(True)
        layout.addWidget(lbl)

        self._combo = QComboBox()
        for col in columns:
            self._combo.addItem(col)
        layout.addWidget(self._combo)

        btn_row = QHBoxLayout()
        btn_ok = QPushButton("OK")
        btn_ok.setDefault(True)
        btn_cancel = QPushButton("Cancel")
        btn_row.addStretch(1)
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

        btn_ok.clicked.connect(self._on_ok)
        btn_cancel.clicked.connect(self.reject)

        self.setStyleSheet(
            """
            QDialog { background-color: #F4F4F4; }
            QLabel { color: #111F35; font-family: "Segoe UI"; font-size: 13px; }
            QComboBox {
                background-color: #FFFFFF; color: #111111;
                border: 1px solid #A9A9A9; border-radius: 6px;
                min-height: 28px; padding: 0 8px;
                font-family: "Segoe UI"; font-size: 12px;
            }
            QPushButton {
                background-color: #9EA3AB; color: #000000;
                border: 1px solid #8B9098; border-radius: 8px;
                min-height: 28px; min-width: 72px; padding: 0 12px;
                font-family: "Segoe UI"; font-size: 12px; font-weight: 600;
            }
            QPushButton:pressed { background-color: #111F35; color: #FFFFFF; }
            """
        )

    def _on_ok(self) -> None:
        self._selected = self._combo.currentText()
        self.accept()

    def get_selected(self) -> str | None:
        return self._selected


class _SingleTrackerWindow(QDialog):
    """Value Duplicate Check – Single Tracker window."""

    _COL_VALUE = 0
    _COL_DUP   = 1
    _COL_PARENT = 2

    def __init__(self, parent: QWidget | None = None, start_dir: str = "") -> None:
        super().__init__(parent)
        self._browse_dir = start_dir
        self.setWindowTitle("Value Duplicate Check – Single Tracker")
        self.resize(990, 560)
        self.setMinimumSize(736, 400)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # ── title row ────────────────────────────────────────────────────────
        title_row = QHBoxLayout()
        title_row.setSpacing(8)
        title_lbl = QLabel("Value Duplicate Check")
        title_lbl.setObjectName("stTitle")
        title_row.addWidget(title_lbl, 0)
        title_row.addSpacing(60)
        self.label_visible_count = QLabel("Count: <b>0</b>")
        self.label_visible_count.setObjectName("stCountLabel")
        self.label_visible_count.setTextFormat(Qt.TextFormat.RichText)
        title_row.addWidget(self.label_visible_count, 0)
        title_row.addStretch(1)
        layout.addLayout(title_row)

        # ── controls row ─────────────────────────────────────────────────────
        ctrl = QHBoxLayout()
        ctrl.setSpacing(8)

        row_count_label = QLabel("Row count")
        row_count_label.setObjectName("stCountLabel")
        ctrl.addWidget(row_count_label, 0)

        self.row_count_input = QLineEdit("5")
        self.row_count_input.setObjectName("stInput")
        self.row_count_input.setFixedWidth(70)
        ctrl.addWidget(self.row_count_input, 0)

        self.btn_update_rows = QPushButton("Update")
        self.btn_update_rows.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_update_rows, 0)

        self.btn_reset_table = QPushButton("Reset")
        self.btn_reset_table.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_reset_table, 0)

        self.btn_delete_row = QPushButton("Delete")
        self.btn_delete_row.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_delete_row, 0)

        self.btn_undo_table = QPushButton("Undo")
        self.btn_undo_table.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_undo_table, 0)

        ctrl.addSpacing(6)

        start_row_lbl = QLabel("Header Row")
        start_row_lbl.setObjectName("stCountLabel")
        ctrl.addWidget(start_row_lbl, 0)

        self.starting_row_input = QLineEdit("7")
        self.starting_row_input.setObjectName("stInput")
        self.starting_row_input.setFixedWidth(70)
        ctrl.addWidget(self.starting_row_input, 0)

        self.btn_import = QPushButton("Import")
        self.btn_import.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_import, 0)

        self.btn_evaluate = QPushButton("Evaluate")
        self.btn_evaluate.setObjectName("stPrimaryBtn")
        self.btn_evaluate.setMinimumWidth(120)
        ctrl.addWidget(self.btn_evaluate, 0)

        self.btn_export = QPushButton("Export")
        self.btn_export.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_export, 0)

        ctrl.addStretch(1)
        layout.addLayout(ctrl)

        # ── table ─────────────────────────────────────────────────────────────
        self.table = _ClipboardTableWidget(5, 3, self)
        self.table.setObjectName("stTable")
        self._header_labels = ["Value", "Duplicate or Not", "Parent Row"]
        self.table.setHorizontalHeaderLabels(self._header_labels)
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.horizontalHeader().setDefaultSectionSize(200)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 260)
        self.table.setColumnWidth(1, 200)

        for row in range(5):
            for col in range(3):
                self.table.setItem(row, col, QTableWidgetItem(""))
            # lock cols 1 & 2 from direct editing
            for col in (self._COL_DUP, self._COL_PARENT):
                item = self.table.item(row, col)
                if item:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

        self._undo_stack: list[list[list[str]]] = []
        self._is_restoring_undo = False
        self._is_batch_edit = False

        layout.addWidget(self.table)

        # ── connections ───────────────────────────────────────────────────────
        self.btn_update_rows.clicked.connect(self._on_update_rows)
        self.btn_reset_table.clicked.connect(self._on_reset_table)
        self.btn_delete_row.clicked.connect(self._on_delete_row)
        self.btn_undo_table.clicked.connect(self._on_undo)
        self.btn_import.clicked.connect(self._on_import)
        self.btn_evaluate.clicked.connect(self._on_evaluate)
        self.btn_export.clicked.connect(self._on_export)
        self.table.itemChanged.connect(self._on_item_changed)

        self._push_undo_snapshot(force=True)
        self._update_count_label()

        self.setStyleSheet(
            """
            QDialog { background-color: #F4F4F4; }

            QTableWidget#stTable {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                gridline-color: #B8B8B8;
                selection-background-color: #DCE6F5;
                selection-color: #111111;
                font-family: "Segoe UI";
                font-size: 12px;
            }
            QTableWidget#stTable QHeaderView::section {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #7D8694;
                padding: 6px 8px;
                font-weight: 700;
            }
            QTableWidget#stTable QTableCornerButton::section {
                background-color: #111F35;
                border: 1px solid #7D8694;
            }
            QLabel#stTitle {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 18px;
                font-weight: 800;
            }
            QLabel#stCountLabel {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 600;
            }
            QLineEdit#stInput {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 8px;
                font-family: "Segoe UI";
                font-size: 12px;
            }
            QPushButton#stGrayBtn {
                background-color: #9EA3AB;
                color: #000000;
                border: 1px solid #8B9098;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }
            QPushButton#stGrayBtn:pressed {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #111F35;
            }
            QPushButton#stPrimaryBtn {
                background-color: #8A244B;
                color: #FFFFFF;
                border: 1px solid #8A244B;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 700;
            }
            QPushButton#stPrimaryBtn:pressed {
                background-color: #F63049;
                border: 1px solid #F63049;
            }
            """
        )

    # ── helpers ───────────────────────────────────────────────────────────────

    def _cell_text(self, row: int, col: int) -> str:
        item = self.table.item(row, col)
        return item.text() if item is not None else ""

    def _set_cell(self, row: int, col: int, value: str, editable: bool = False) -> None:
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem("")
            self.table.setItem(row, col, item)
        flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
        if editable:
            flags |= Qt.ItemFlag.ItemIsEditable
        item.setFlags(flags)
        item.setText(value)

    def _update_count_label(self) -> None:
        count = sum(1 for r in range(self.table.rowCount()) if not self.table.isRowHidden(r))
        self.label_visible_count.setText(f"Count: <b>{count}</b>")

    def _capture_state(self) -> list[list[str]]:
        return [
            [self._cell_text(r, c) for c in range(self.table.columnCount())]
            for r in range(self.table.rowCount())
        ]

    def _restore_state(self, state: list[list[str]]) -> None:
        self._is_restoring_undo = True
        self.table.blockSignals(True)
        try:
            self.table.setRowCount(max(1, len(state)))
            for r, row_vals in enumerate(state):
                for c in range(self.table.columnCount()):
                    val = row_vals[c] if c < len(row_vals) else ""
                    editable = (c == self._COL_VALUE)
                    self._set_cell(r, c, val, editable=editable)
        finally:
            self.table.blockSignals(False)
            self._is_restoring_undo = False
        self.row_count_input.setText(str(self.table.rowCount()))
        self._update_count_label()

    def _push_undo_snapshot(self, force: bool = False) -> None:
        if self._is_restoring_undo or self._is_batch_edit:
            return
        snap = self._capture_state()
        if not force and self._undo_stack and self._undo_stack[-1] == snap:
            return
        self._undo_stack.append(snap)
        if len(self._undo_stack) > 200:
            self._undo_stack.pop(0)

    def _init_rows(self, count: int) -> None:
        self.table.setRowCount(count)
        for r in range(count):
            self._set_cell(r, self._COL_VALUE, self._cell_text(r, self._COL_VALUE) or "", editable=True)
            for c in (self._COL_DUP, self._COL_PARENT):
                self._set_cell(r, c, "", editable=False)

    # ── slot handlers ─────────────────────────────────────────────────────────

    def _on_item_changed(self, _item: QTableWidgetItem) -> None:
        if not self._is_batch_edit:
            self._update_count_label()
            self._push_undo_snapshot()

    def _on_update_rows(self) -> None:
        try:
            n = max(1, min(int(self.row_count_input.text().strip()), 5000))
        except ValueError:
            self.row_count_input.setText(str(self.table.rowCount()))
            return
        self._init_rows(n)
        self.row_count_input.setText(str(n))
        self._update_count_label()
        self._push_undo_snapshot()

    def _on_reset_table(self) -> None:
        self.row_count_input.setText("5")
        self.starting_row_input.setText("7")
        self.table.blockSignals(True)
        self.table.clearContents()
        self.table.setRowCount(5)
        for r in range(5):
            self._set_cell(r, self._COL_VALUE, "", editable=True)
            for c in (self._COL_DUP, self._COL_PARENT):
                self._set_cell(r, c, "", editable=False)
        self.table.blockSignals(False)
        self._update_count_label()
        self._push_undo_snapshot()

    def _on_delete_row(self) -> None:
        sel_rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not sel_rows:
            cur = self.table.currentRow()
            if cur >= 0:
                sel_rows = [cur]
            else:
                return
        total = self.table.rowCount()
        if len(sel_rows) >= total:
            self.table.setRowCount(1)
            self._set_cell(0, self._COL_VALUE, "", editable=True)
            for c in (self._COL_DUP, self._COL_PARENT):
                self._set_cell(0, c, "", editable=False)
        else:
            for r in sel_rows:
                if 0 <= r < self.table.rowCount():
                    self.table.removeRow(r)
        self.table.clearSelection()
        self.row_count_input.setText(str(self.table.rowCount()))
        self._update_count_label()
        self._push_undo_snapshot()

    def _on_undo(self) -> None:
        if self._is_restoring_undo:
            return
        cur = self._capture_state()
        if not self._undo_stack:
            self._undo_stack.append(cur)
            return
        if self._undo_stack[-1] != cur:
            self._undo_stack.append(cur)
        if len(self._undo_stack) <= 1:
            return
        self._undo_stack.pop()
        self._restore_state(self._undo_stack[-1])

    # ── evaluate ──────────────────────────────────────────────────────────────

    def _on_evaluate(self) -> None:
        n = self.table.rowCount()
        # gather visible (not hidden) values; map (visual_row -> value)
        values: list[str] = [self._cell_text(r, self._COL_VALUE).strip() for r in range(n)]

        # first_seen: value -> row index (0-based display "row N+1")
        first_seen: dict[str, int] = {}
        dup_results: list[str] = [""] * n
        parent_results: list[str] = [""] * n

        for r in range(n):
            v = values[r]
            if v == "":
                continue
            if v not in first_seen:
                first_seen[v] = r
            else:
                dup_results[r] = "duplicate"
                parent_results[r] = f"row {first_seen[v] + 1}"

        self.table.blockSignals(True)
        try:
            for r in range(n):
                self._set_cell(r, self._COL_DUP, dup_results[r], editable=False)
                self._set_cell(r, self._COL_PARENT, parent_results[r], editable=False)
        finally:
            self.table.blockSignals(False)
        self._push_undo_snapshot()

    # ── export ────────────────────────────────────────────────────────────────

    def _on_export(self) -> None:
        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        except ImportError:
            QMessageBox.warning(self, "Export Error", "openpyxl is required for export.\nInstall it with: pip install openpyxl")
            return

        import datetime as _dt
        _ts = _dt.datetime.now().strftime("%Y_%m_%d_%H_%M")
        _default_name = f"duplicate_check_{_ts}.xlsx"
        import os as _os
        _default_path = _os.path.join(self._browse_dir, _default_name) if self._browse_dir else _default_name
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Table", _default_path, "Excel Files (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Duplicate Check"

        header_fill = PatternFill("solid", fgColor="111F35")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_align = Alignment(horizontal="center", vertical="center")
        thin_side = Side(style="thin", color="7D8694")
        header_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        col_labels = ["Value", "Duplicate or Not", "Parent Row"]
        col_widths  = [32, 22, 18]
        for ci, (label, width) in enumerate(zip(col_labels, col_widths), start=1):
            cell = ws.cell(row=1, column=ci, value=label)
            cell.fill   = header_fill
            cell.font   = header_font
            cell.alignment = header_align
            cell.border = header_border
            ws.column_dimensions[cell.column_letter].width = width
        ws.row_dimensions[1].height = 22

        body_font   = Font(name="Segoe UI", size=10)
        body_align  = Alignment(vertical="center")
        body_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for row_i in range(self.table.rowCount()):
            if self.table.isRowHidden(row_i):
                continue
            for ci in range(3):
                val = self._cell_text(row_i, ci)
                cell = ws.cell(row=ws.max_row + 1 if ci == 0 else ws.max_row, column=ci + 1, value=val)
                cell.font      = body_font
                cell.alignment = body_align
                cell.border    = body_border

        try:
            wb.save(path)
            QMessageBox.information(self, "Export", f"Exported successfully:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export Error", str(exc))

    # ── import ────────────────────────────────────────────────────────────────

    def _on_import(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Import File", self._browse_dir, "Excel / CSV Files (*.xlsx *.xls *.csv)"
        )
        if not path:
            return

        # Resolve header row (1-based from user input)
        try:
            header_row = max(1, int(self.starting_row_input.text().strip())) - 1  # 0-based index
        except ValueError:
            header_row = 0

        try:
            import pandas as _pd_imp
            if path.lower().endswith(".csv"):
                df = _pd_imp.read_csv(path, header=None, dtype=str)
            else:
                df = _pd_imp.read_excel(path, header=None, dtype=str)
        except Exception as exc:
            QMessageBox.critical(self, "Import Error", str(exc))
            return

        if df.empty:
            QMessageBox.information(self, "Import", "The file appears to be empty.")
            return

        # Normalise a raw cell value — treat NaN/None/empty as ""
        def _norm(v) -> str:
            s = str(v).strip() if v is not None else ""
            return "" if s.lower() in ("nan", "none", "") else s

        # Use the user-specified header row for column labels
        if header_row < df.shape[0]:
            header_cells = [_norm(v) for v in df.iloc[header_row]]
        else:
            header_cells = [f"Column {i+1}" for i in range(df.shape[1])]

        # Build col_map from the header row (only columns with a non-empty label)
        col_map: list[tuple[int, str]] = [
            (i, header_cells[i]) for i in range(df.shape[1]) if header_cells[i] != ""
        ]
        if not col_map:
            # Fallback: no recognisable headers — label by number
            col_map = [(i, f"Column {i+1}") for i in range(min(df.shape[1], 8))]

        # Data rows start immediately after the header row
        data_rows = df.iloc[header_row + 1:].reset_index(drop=True)

        col_identifiers = [label for _, label in col_map]
        col_idx_map = {label: idx for idx, label in col_map}

        # If only one meaningful column, use it directly; otherwise ask user
        if len(col_map) == 1:
            col_idx = col_map[0][0]
        else:
            picker = _SingleTrackerColumnPickerDialog(col_identifiers, self)
            if picker.exec() != QDialog.DialogCode.Accepted or picker.get_selected() is None:
                return
            chosen = picker.get_selected()
            col_idx = col_idx_map.get(chosen, col_map[0][0])

        values = [_norm(v) for v in data_rows.iloc[:, col_idx]]

        # Always paste values starting at table row 0
        start_row = 0

        required = start_row + len(values)
        if required > self.table.rowCount():
            self.table.setRowCount(required)
            for r in range(self.table.rowCount()):
                for c in range(3):
                    if self.table.item(r, c) is None:
                        editable = (c == self._COL_VALUE)
                        self._set_cell(r, c, "", editable=editable)

        self.table.blockSignals(True)
        self._is_batch_edit = True
        try:
            for i, val in enumerate(values):
                r = start_row + i
                self._set_cell(r, self._COL_VALUE, val, editable=True)
                self._set_cell(r, self._COL_DUP,   "", editable=False)
                self._set_cell(r, self._COL_PARENT, "", editable=False)
        finally:
            self.table.blockSignals(False)
            self._is_batch_edit = False

        self.row_count_input.setText(str(self.table.rowCount()))
        self._update_count_label()
        self._push_undo_snapshot()


class _MultipleTrackerWindow(QDialog):
    """Value Duplicate Check – Multiple Trackers window."""

    _COL_VALUE = 0
    _COL_TRACKERS = 1

    def __init__(self, parent: QWidget | None = None, start_dir: str = "") -> None:
        super().__init__(parent)
        self._browse_dir = start_dir
        self.setWindowTitle("Duplicate Window – Multiple Trackers")
        self.resize(990, 560)
        self.setMinimumSize(736, 400)

        # Internal state: imported per-tracker data
        self._tracker_data: dict[str, list[str]] = {}  # tracker_name -> list of values
        self._selected_column: str | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # ── title row ───────────────────────────────────────────────────
        title_row = QHBoxLayout()
        title_row.setSpacing(8)
        title_lbl = QLabel("Value Duplicate Check – Multiple Trackers")
        title_lbl.setObjectName("stTitle")
        title_row.addWidget(title_lbl, 0)
        title_row.addSpacing(60)
        self.label_visible_count = QLabel("Count: <b>0</b>")
        self.label_visible_count.setObjectName("stCountLabel")
        self.label_visible_count.setTextFormat(Qt.TextFormat.RichText)
        title_row.addWidget(self.label_visible_count, 0)
        title_row.addStretch(1)
        layout.addLayout(title_row)

        # ── controls row ────────────────────────────────────────────────
        ctrl = QHBoxLayout()
        ctrl.setSpacing(8)

        row_count_label = QLabel("Row count")
        row_count_label.setObjectName("stCountLabel")
        ctrl.addWidget(row_count_label, 0)

        self.row_count_input = QLineEdit("5")
        self.row_count_input.setObjectName("stInput")
        self.row_count_input.setFixedWidth(70)
        ctrl.addWidget(self.row_count_input, 0)

        self.btn_update_rows = QPushButton("Update")
        self.btn_update_rows.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_update_rows, 0)

        self.btn_reset_table = QPushButton("Reset")
        self.btn_reset_table.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_reset_table, 0)

        self.btn_delete_row = QPushButton("Delete")
        self.btn_delete_row.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_delete_row, 0)

        self.btn_undo_table = QPushButton("Undo")
        self.btn_undo_table.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_undo_table, 0)

        ctrl.addSpacing(6)

        start_row_lbl = QLabel("Header Row")
        start_row_lbl.setObjectName("stCountLabel")
        ctrl.addWidget(start_row_lbl, 0)

        self.starting_row_input = QLineEdit("7")
        self.starting_row_input.setObjectName("stInput")
        self.starting_row_input.setFixedWidth(70)
        ctrl.addWidget(self.starting_row_input, 0)

        self.btn_import = QPushButton("Import")
        self.btn_import.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_import, 0)

        self.btn_evaluate = QPushButton("Evaluate")
        self.btn_evaluate.setObjectName("stPrimaryBtn")
        self.btn_evaluate.setMinimumWidth(120)
        ctrl.addWidget(self.btn_evaluate, 0)

        self.btn_export = QPushButton("Export")
        self.btn_export.setObjectName("stGrayBtn")
        ctrl.addWidget(self.btn_export, 0)

        ctrl.addStretch(1)
        layout.addLayout(ctrl)

        # ── table (2 columns) ──────────────────────────────────────────
        self.table = _ClipboardTableWidget(5, 2, self)
        self.table.setObjectName("stTable")
        self._header_labels = ["Value", "Tracker files sharing the value"]
        self.table.setHorizontalHeaderLabels(self._header_labels)
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.horizontalHeader().setDefaultSectionSize(200)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 260)

        for row in range(5):
            for col in range(2):
                self.table.setItem(row, col, QTableWidgetItem(""))
            item = self.table.item(row, self._COL_TRACKERS)
            if item:
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

        self._undo_stack: list[list[list[str]]] = []
        self._is_restoring_undo = False
        self._is_batch_edit = False

        layout.addWidget(self.table)

        # ── connections ─────────────────────────────────────────────────
        self.btn_update_rows.clicked.connect(self._on_update_rows)
        self.btn_reset_table.clicked.connect(self._on_reset_table)
        self.btn_delete_row.clicked.connect(self._on_delete_row)
        self.btn_undo_table.clicked.connect(self._on_undo)
        self.btn_import.clicked.connect(self._on_import)
        self.btn_evaluate.clicked.connect(self._on_evaluate)
        self.btn_export.clicked.connect(self._on_export)
        self.table.itemChanged.connect(self._on_item_changed)

        self._push_undo_snapshot(force=True)
        self._update_count_label()

        self.setStyleSheet(
            """
            QDialog { background-color: #F4F4F4; }

            QTableWidget#stTable {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                gridline-color: #B8B8B8;
                selection-background-color: #DCE6F5;
                selection-color: #111111;
                font-family: "Segoe UI";
                font-size: 12px;
            }
            QTableWidget#stTable QHeaderView::section {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #7D8694;
                padding: 6px 8px;
                font-weight: 700;
            }
            QTableWidget#stTable QTableCornerButton::section {
                background-color: #111F35;
                border: 1px solid #7D8694;
            }
            QLabel#stTitle {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 18px;
                font-weight: 800;
            }
            QLabel#stCountLabel {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 600;
            }
            QLineEdit#stInput {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 8px;
                font-family: "Segoe UI";
                font-size: 12px;
            }
            QPushButton#stGrayBtn {
                background-color: #9EA3AB;
                color: #000000;
                border: 1px solid #8B9098;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }
            QPushButton#stGrayBtn:pressed {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #111F35;
            }
            QPushButton#stPrimaryBtn {
                background-color: #8A244B;
                color: #FFFFFF;
                border: 1px solid #8A244B;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 700;
            }
            QPushButton#stPrimaryBtn:pressed {
                background-color: #F63049;
                border: 1px solid #F63049;
            }
            """
        )

    # ── helpers ──────────────────────────────────────────────────────────

    def _cell_text(self, row: int, col: int) -> str:
        item = self.table.item(row, col)
        return item.text() if item is not None else ""

    def _set_cell(self, row: int, col: int, value: str, editable: bool = False) -> None:
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem("")
            self.table.setItem(row, col, item)
        flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
        if editable:
            flags |= Qt.ItemFlag.ItemIsEditable
        item.setFlags(flags)
        item.setText(value)

    def _update_count_label(self) -> None:
        count = sum(1 for r in range(self.table.rowCount()) if not self.table.isRowHidden(r))
        self.label_visible_count.setText(f"Count: <b>{count}</b>")

    def _capture_state(self) -> list[list[str]]:
        return [
            [self._cell_text(r, c) for c in range(self.table.columnCount())]
            for r in range(self.table.rowCount())
        ]

    def _restore_state(self, state: list[list[str]]) -> None:
        self._is_restoring_undo = True
        self.table.blockSignals(True)
        try:
            self.table.setRowCount(max(1, len(state)))
            for r, row_vals in enumerate(state):
                for c in range(self.table.columnCount()):
                    val = row_vals[c] if c < len(row_vals) else ""
                    editable = (c == self._COL_VALUE)
                    self._set_cell(r, c, val, editable=editable)
        finally:
            self.table.blockSignals(False)
            self._is_restoring_undo = False
        self.row_count_input.setText(str(self.table.rowCount()))
        self._update_count_label()

    def _push_undo_snapshot(self, force: bool = False) -> None:
        if self._is_restoring_undo or self._is_batch_edit:
            return
        snap = self._capture_state()
        if not force and self._undo_stack and self._undo_stack[-1] == snap:
            return
        self._undo_stack.append(snap)
        if len(self._undo_stack) > 200:
            self._undo_stack.pop(0)

    def _init_rows(self, count: int) -> None:
        self.table.setRowCount(count)
        for r in range(count):
            self._set_cell(r, self._COL_VALUE, self._cell_text(r, self._COL_VALUE) or "", editable=True)
            self._set_cell(r, self._COL_TRACKERS, "", editable=False)

    def _resize_rows_to_content(self) -> None:
        """Resize each row height so multi-line cell text is fully visible."""
        fm = self.table.fontMetrics()
        line_h = fm.height() + 4  # per-line height with small padding
        pad = 10  # top+bottom cell padding
        min_h = 30
        for r in range(self.table.rowCount()):
            max_lines = 1
            for c in range(self.table.columnCount()):
                text = self._cell_text(r, c)
                lines = text.count("\n") + 1 if text else 1
                if lines > max_lines:
                    max_lines = lines
            self.table.setRowHeight(r, max(min_h, max_lines * line_h + pad))

    # ── slot handlers ───────────────────────────────────────────────────

    def _on_item_changed(self, _item: QTableWidgetItem) -> None:
        if not self._is_batch_edit:
            self._update_count_label()
            self._push_undo_snapshot()

    def _on_update_rows(self) -> None:
        try:
            n = max(1, min(int(self.row_count_input.text().strip()), 5000))
        except ValueError:
            self.row_count_input.setText(str(self.table.rowCount()))
            return
        self._init_rows(n)
        self.row_count_input.setText(str(n))
        self._update_count_label()
        self._push_undo_snapshot()

    def _on_reset_table(self) -> None:
        self.row_count_input.setText("5")
        self.starting_row_input.setText("7")
        self._tracker_data.clear()
        self._selected_column = None
        self.table.blockSignals(True)
        self.table.clearContents()
        self.table.setRowCount(5)
        for r in range(5):
            self._set_cell(r, self._COL_VALUE, "", editable=True)
            self._set_cell(r, self._COL_TRACKERS, "", editable=False)
        self.table.blockSignals(False)
        self._update_count_label()
        self._push_undo_snapshot()

    def _on_delete_row(self) -> None:
        sel_rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not sel_rows:
            cur = self.table.currentRow()
            if cur >= 0:
                sel_rows = [cur]
            else:
                return
        total = self.table.rowCount()
        if len(sel_rows) >= total:
            self.table.setRowCount(1)
            self._set_cell(0, self._COL_VALUE, "", editable=True)
            self._set_cell(0, self._COL_TRACKERS, "", editable=False)
        else:
            for r in sel_rows:
                if 0 <= r < self.table.rowCount():
                    self.table.removeRow(r)
        self.table.clearSelection()
        self.row_count_input.setText(str(self.table.rowCount()))
        self._update_count_label()
        self._push_undo_snapshot()

    def _on_undo(self) -> None:
        if self._is_restoring_undo:
            return
        cur = self._capture_state()
        if not self._undo_stack:
            self._undo_stack.append(cur)
            return
        if self._undo_stack[-1] != cur:
            self._undo_stack.append(cur)
        if len(self._undo_stack) <= 1:
            return
        self._undo_stack.pop()
        self._restore_state(self._undo_stack[-1])

    # ── import ──────────────────────────────────────────────────────────

    def _on_import(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Tracker Files", self._browse_dir,
            "Excel Files (*.xlsx *.xls *.csv)"
        )
        if not paths:
            return

        try:
            header_row = max(1, int(self.starting_row_input.text().strip())) - 1
        except ValueError:
            header_row = 0

        import pandas as _pd
        from pathlib import Path as _Path

        def _norm(v) -> str:
            s = str(v).strip() if v is not None else ""
            return "" if s.lower() in ("nan", "none", "") else s

        # Collect headers from all files to find common columns
        all_headers: dict[str, list[tuple[int, str]]] = {}  # path -> [(col_idx, label)]
        file_dfs: dict[str, _pd.DataFrame] = {}

        for fp in paths:
            try:
                if fp.lower().endswith(".csv"):
                    df = _pd.read_csv(fp, header=None, dtype=str)
                else:
                    df = _pd.read_excel(fp, header=None, dtype=str)
            except Exception:
                continue

            file_dfs[fp] = df

            if header_row < df.shape[0]:
                hdr_cells = [_norm(v) for v in df.iloc[header_row]]
            else:
                hdr_cells = [f"Column {i+1}" for i in range(df.shape[1])]

            col_map = [(i, hdr_cells[i]) for i in range(df.shape[1]) if hdr_cells[i] != ""]
            all_headers[fp] = col_map

        if not file_dfs:
            QMessageBox.information(self, "Import", "No valid files could be read.")
            return

        # Gather unique column labels across all files
        unique_labels: list[str] = []
        seen: set[str] = set()
        for col_map in all_headers.values():
            for _, label in col_map:
                key = label.strip().lower()
                if key not in seen:
                    seen.add(key)
                    unique_labels.append(label)

        if not unique_labels:
            QMessageBox.information(self, "Import", "No column headers detected.")
            return

        # Ask user to pick one column
        if len(unique_labels) == 1:
            chosen = unique_labels[0]
        else:
            picker = _SingleTrackerColumnPickerDialog(unique_labels, self)
            if picker.exec() != QDialog.DialogCode.Accepted or picker.get_selected() is None:
                return
            chosen = picker.get_selected()

        self._selected_column = chosen
        chosen_lower = chosen.strip().lower()

        # Read all values from the chosen column across all trackers, dedup
        tracker_data: dict[str, list[str]] = {}  # tracker_name -> values
        all_values_ordered: list[str] = []
        all_values_set: set[str] = set()

        for fp, df in file_dfs.items():
            tracker_name = _Path(fp).stem  # filename without extension
            col_map = all_headers.get(fp, [])

            # Find the column index for the chosen column in this file
            col_idx = None
            for idx, label in col_map:
                if label.strip().lower() == chosen_lower:
                    col_idx = idx
                    break

            if col_idx is None:
                tracker_data[tracker_name] = []
                continue

            data_rows = df.iloc[header_row + 1:].reset_index(drop=True)
            values = []
            for v in data_rows.iloc[:, col_idx]:
                val = _norm(v)
                if val:
                    values.append(val)
                    if val not in all_values_set:
                        all_values_set.add(val)
                        all_values_ordered.append(val)

            tracker_data[tracker_name] = values

        self._tracker_data = tracker_data

        # Populate the table with deduped values
        n = len(all_values_ordered)
        if n == 0:
            QMessageBox.information(self, "Import", "No values found in the selected column.")
            return

        self.table.blockSignals(True)
        self._is_batch_edit = True
        try:
            self.table.setRowCount(n)
            for r, val in enumerate(all_values_ordered):
                self._set_cell(r, self._COL_VALUE, val, editable=True)
                self._set_cell(r, self._COL_TRACKERS, "", editable=False)
        finally:
            self.table.blockSignals(False)
            self._is_batch_edit = False

        self.row_count_input.setText(str(n))
        self._update_count_label()
        self._resize_rows_to_content()
        self._push_undo_snapshot()

    # ── evaluate ────────────────────────────────────────────────────────

    def _on_evaluate(self) -> None:
        if not self._tracker_data:
            QMessageBox.information(self, "Evaluate", "No tracker data loaded. Import trackers first.")
            return

        n = self.table.rowCount()

        # Build a lookup: value -> set of tracker names that contain it
        value_to_trackers: dict[str, set[str]] = {}
        for tracker_name, values in self._tracker_data.items():
            for v in values:
                value_to_trackers.setdefault(v, set()).add(tracker_name)

        # Collect only rows whose value appears in 2+ trackers
        dup_rows: list[tuple[str, str]] = []
        for r in range(n):
            val = self._cell_text(r, self._COL_VALUE).strip()
            if not val:
                continue
            trackers = value_to_trackers.get(val, set())
            if len(trackers) >= 2:
                dup_rows.append((val, "\n".join(sorted(trackers))))

        # Rebuild the table with only duplicate rows
        self.table.blockSignals(True)
        self._is_batch_edit = True
        try:
            self.table.setRowCount(max(1, len(dup_rows)))
            if dup_rows:
                for r, (val, names_str) in enumerate(dup_rows):
                    self._set_cell(r, self._COL_VALUE, val, editable=True)
                    self._set_cell(r, self._COL_TRACKERS, names_str, editable=False)
            else:
                self._set_cell(0, self._COL_VALUE, "", editable=True)
                self._set_cell(0, self._COL_TRACKERS, "", editable=False)
        finally:
            self.table.blockSignals(False)
            self._is_batch_edit = False

        self.row_count_input.setText(str(self.table.rowCount()))
        self._update_count_label()
        self._resize_rows_to_content()
        self._push_undo_snapshot()

    # ── export ──────────────────────────────────────────────────────────

    def _on_export(self) -> None:
        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
        except ImportError:
            QMessageBox.warning(self, "Export Error",
                                "openpyxl is required.\nInstall with: pip install openpyxl")
            return

        import datetime as _dt
        import os as _os

        _ts = _dt.datetime.now().strftime("%Y_%m_%d_%H_%M")
        _default_name = f"duplicate_check_{_ts}.xlsx"
        _default_path = _os.path.join(self._browse_dir, _default_name) if self._browse_dir else _default_name
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Table", _default_path, "Excel Files (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        # ── Style constants ─────────────────────────────────────────────
        header_fill   = PatternFill("solid", fgColor="111F35")
        header_font   = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_align  = Alignment(horizontal="center", vertical="center")
        thin_side     = Side(style="thin", color="7D8694")
        hdr_border    = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        body_font     = Font(name="Segoe UI", size=10)
        body_align    = Alignment(vertical="center", wrap_text=True)
        body_border   = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        def _style_header(ws, headers, widths):
            for ci, (h, w) in enumerate(zip(headers, widths), start=1):
                cell = ws.cell(row=1, column=ci, value=h)
                cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, header_align, hdr_border
                ws.column_dimensions[get_column_letter(ci)].width = w
            ws.row_dimensions[1].height = 22

        def _write_row(ws, row_num, values):
            for ci, val in enumerate(values, start=1):
                cell = ws.cell(row=row_num, column=ci, value=val)
                cell.font, cell.alignment, cell.border = body_font, body_align, body_border

        wb = openpyxl.Workbook()

        # ── Sheet 1: dup_check_main (general summary) ──────────────────
        ws_main = wb.active
        ws_main.title = "dup_check_main"
        _style_header(ws_main, ["Value", "Tracker files sharing the value"], [32, 50])

        for r in range(self.table.rowCount()):
            if self.table.isRowHidden(r):
                continue
            val = self._cell_text(r, self._COL_VALUE)
            trackers_text = self._cell_text(r, self._COL_TRACKERS)
            _write_row(ws_main, ws_main.max_row + 1, [val, trackers_text])

        # ── Per-tracker sheets ──────────────────────────────────────────
        # Build value -> set(tracker_names) from _tracker_data
        value_to_trackers: dict[str, set[str]] = {}
        for tracker_name, values in self._tracker_data.items():
            for v in values:
                value_to_trackers.setdefault(v, set()).add(tracker_name)

        tracker_names = sorted(self._tracker_data.keys())
        for sheet_idx, tracker_name in enumerate(tracker_names, start=1):
            sheet_title = f"dup_check_{sheet_idx}"
            # Ensure sheet name is valid (max 31 chars)
            if len(sheet_title) > 31:
                sheet_title = sheet_title[:31]
            ws = wb.create_sheet(title=sheet_title)
            _style_header(ws, ["File Name", "Value", "Duplicate or Not",
                               "Tracker files sharing the value"],
                          [28, 32, 22, 50])

            tracker_values = self._tracker_data[tracker_name]
            seen: set[str] = set()
            for val in tracker_values:
                if val in seen or not val:
                    continue
                seen.add(val)
                trackers_with_val = value_to_trackers.get(val, set())
                # Only include values that are shared with OTHER trackers
                other_trackers = trackers_with_val - {tracker_name}
                if not other_trackers:
                    continue
                others_str = "\n".join(sorted(other_trackers))
                _write_row(ws, ws.max_row + 1, [tracker_name, val, "duplicate", others_str])

        try:
            wb.save(path)
            QMessageBox.information(self, "Export", f"Exported successfully:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export Error", str(exc))


class _SetMismatchDialog(QDialog):
    """Shown when target and master cu-level file sets do not fully match."""

    def __init__(
        self,
        target_levels: "list[int]",
        master_levels: "list[int]",
        common_levels: "list[int]",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Compare Set — File Mismatch")
        self.setModal(True)
        self.resize(420, 180)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        master_str = ", ".join(f"cu{l}" for l in sorted(master_levels))
        target_str = ", ".join(f"cu{l}" for l in sorted(target_levels))
        common_str = ", ".join(f"cu{l}" for l in sorted(common_levels))

        msg = QLabel(
            f"Compare files:\n"
            f"RSD: Master:  {master_str}\n"
            f"RSD: Target:  {target_str}\n\n"
            f"File count mismatch — tool will proceed to compare {common_str}"
        )
        msg.setWordWrap(True)
        layout.addWidget(msg)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_proceed = QPushButton("Proceed")
        btn_proceed.setObjectName("compareRunBtn")
        btn_cancel = QPushButton("Cancel")
        btn_cancel.setObjectName("compareGrayBtn")
        btn_row.addWidget(btn_proceed)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

        btn_proceed.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)


class _SdcSelectionDialog(QDialog):
    """
    Shown when multiple SDC output files were created.
    User selects which ones to run the Reference Collector against.
    """

    def __init__(self, sdc_paths: list[str], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Reference Collector – Select SDC File(s)")
        self.setMinimumWidth(480)

        self.selected_paths: list[str] = []
        self._checkboxes: list[tuple[QCheckBox, str]] = []

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        lbl = QLabel("Select SDC file(s) to use for Reference Collector:")
        lbl.setWordWrap(True)
        layout.addWidget(lbl)

        for path in sdc_paths:
            cb = QCheckBox(Path(path).name)
            cb.setChecked(True)
            self._checkboxes.append((cb, path))
            layout.addWidget(cb)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_select_all = QPushButton("Select All")
        btn_select_all.clicked.connect(self._select_all)
        btn_row.addWidget(btn_select_all)
        btn_ok = QPushButton("OK")
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self._on_ok)
        btn_row.addWidget(btn_ok)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

    def _select_all(self) -> None:
        for cb, _ in self._checkboxes:
            cb.setChecked(True)

    def _on_ok(self) -> None:
        self.selected_paths = [p for cb, p in self._checkboxes if cb.isChecked()]
        self.accept()


class _TscOption1TableDialog(QDialog):
    """Tracker Information window opened from TSC Option 1."""

    _BLANK_FILTER = "__BLANK__"

    TSC_HEADERS = [
        "SBU",
        "Product Name",
        "IDH Number",
        "Build Type",
        "Status",
        "Packaging Type",
        "Packaging Size",
        "Project Name",
        "Basic Number",
        "Label Size",
        "Actual Hand-In Date",
        "Total Cost",
        "Is Deployment",
        "Packshot Naming",
        "Working Files",
    ]

    class _FilterPopup(QDialog):
        def __init__(
            self,
            values: list[tuple[str, str]],
            selected_values: set[str] | None,
            parent: QWidget | None = None,
        ) -> None:
            super().__init__(parent)
            self.setWindowFlags(Qt.WindowType.Popup)
            self.setMinimumSize(260, 330)
            self.resize(260, 330)

            self._is_syncing = False
            self._value_items: list[tuple[QListWidgetItem, str, str]] = []

            layout = QVBoxLayout(self)
            layout.setContentsMargins(4, 4, 4, 4)
            layout.setSpacing(4)

            self.search_input = QLineEdit(self)
            self.search_input.setPlaceholderText("Search")
            self.search_input.setClearButtonEnabled(True)
            self.search_input.setObjectName("packshotRowCountInput")
            layout.addWidget(self.search_input)

            self.list_widget = QListWidget(self)
            self.list_widget.setSelectionMode(QListWidget.SelectionMode.NoSelection)
            layout.addWidget(self.list_widget, 1)

            self.item_select_all = QListWidgetItem("(Select All)")
            self.item_select_all.setFlags(self.item_select_all.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            self.item_select_all.setCheckState(Qt.CheckState.Checked)
            self.item_select_all.setData(Qt.ItemDataRole.UserRole, "__SELECT_ALL__")
            self.list_widget.addItem(self.item_select_all)

            selected = selected_values if selected_values is not None else {value for value, _label in values}

            for value, label in values:
                item = QListWidgetItem(label)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if value in selected else Qt.CheckState.Unchecked)
                item.setData(Qt.ItemDataRole.UserRole, value)
                self.list_widget.addItem(item)
                self._value_items.append((item, value, label))

            button_row = QHBoxLayout()
            button_row.addStretch(1)
            self.btn_apply = QPushButton("Apply")
            self.btn_apply.setObjectName("packshotUpdateRowsBtn")
            self.btn_cancel = QPushButton("Cancel")
            self.btn_cancel.setObjectName("packshotUpdateRowsBtn")
            button_row.addWidget(self.btn_apply, 0)
            button_row.addWidget(self.btn_cancel, 0)
            layout.addLayout(button_row)

            self.search_input.textChanged.connect(self._apply_search)
            self.list_widget.itemChanged.connect(self._on_item_changed)
            self.btn_apply.clicked.connect(self.accept)
            self.btn_cancel.clicked.connect(self.reject)

            self.search_input.setFocus()
            self._sync_select_all_state()

        def _apply_search(self, text: str) -> None:
            needle = text.strip().lower()
            for item, _value, label in self._value_items:
                item.setHidden(needle not in label.lower())
            self._sync_select_all_state()

        def _on_item_changed(self, item: QListWidgetItem) -> None:
            if self._is_syncing:
                return
            role = item.data(Qt.ItemDataRole.UserRole)
            if role == "__SELECT_ALL__":
                self._is_syncing = True
                try:
                    target_state = item.checkState()
                    for value_item, _value, _label in self._value_items:
                        if value_item.isHidden():
                            continue
                        value_item.setCheckState(target_state)
                finally:
                    self._is_syncing = False
                return
            self._sync_select_all_state()

        def _sync_select_all_state(self) -> None:
            visible_items = [item for item, _value, _label in self._value_items if not item.isHidden()]
            if not visible_items:
                state = Qt.CheckState.Unchecked
            else:
                all_checked = all(i.checkState() == Qt.CheckState.Checked for i in visible_items)
                any_checked = any(i.checkState() == Qt.CheckState.Checked for i in visible_items)
                if all_checked:
                    state = Qt.CheckState.Checked
                elif any_checked:
                    state = Qt.CheckState.PartiallyChecked
                else:
                    state = Qt.CheckState.Unchecked
            self._is_syncing = True
            try:
                self.item_select_all.setCheckState(state)
            finally:
                self._is_syncing = False

        def get_selected_values(self) -> set[str]:
            selected: set[str] = set()
            for item, value, _label in self._value_items:
                if item.checkState() == Qt.CheckState.Checked:
                    selected.add(value)
            return selected

    def __init__(self, parent: QWidget | None = None, start_dir: str = "", tsc_output_dir: str = "") -> None:
        super().__init__(parent)
        self._browse_start_dir = start_dir
        self._tsc_output_dir   = tsc_output_dir
        self.setWindowTitle("Open window")
        self.resize(1300, 620)
        self.setMinimumSize(1080, 520)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # --- Title row with count label ---
        title_row = QHBoxLayout()
        title_row.setSpacing(8)
        title = QLabel("Tracker Information")
        title.setObjectName("packshotTableTitle")
        title_row.addWidget(title, 0)
        title_row.addSpacing(530)
        self.label_row_count_display = QLabel("Count: <b>5</b>")
        self.label_row_count_display.setObjectName("tscCountLabel")
        self.label_row_count_display.setTextFormat(Qt.TextFormat.RichText)
        title_row.addWidget(self.label_row_count_display, 0)
        title_row.addStretch(1)
        layout.addLayout(title_row)

        # --- Config row: Row count, buttons, Import, Reset Filter, Export, Generate Chart ---
        config_row = QHBoxLayout()
        config_row.setSpacing(8)

        row_count_label = QLabel("Row count")
        row_count_label.setObjectName("packshotRowCountLabel")
        config_row.addWidget(row_count_label, 0)

        self.row_count_input = QLineEdit("5")
        self.row_count_input.setObjectName("packshotRowCountInput")
        self.row_count_input.setFixedWidth(45)
        config_row.addWidget(self.row_count_input, 0)

        self.btn_update_rows = QPushButton("Update")
        self.btn_update_rows.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_update_rows, 0)

        self.btn_reset_table = QPushButton("Reset")
        self.btn_reset_table.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_reset_table, 0)

        self.btn_delete_row = QPushButton("Delete")
        self.btn_delete_row.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_delete_row, 0)

        self.btn_undo_table = QPushButton("Undo")
        self.btn_undo_table.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_undo_table, 0)

        config_row.addSpacing(8)

        starting_row_label = QLabel("Starting Row:")
        starting_row_label.setObjectName("packshotRowCountLabel")
        config_row.addWidget(starting_row_label, 0)

        self.starting_row_input = QLineEdit("11")
        self.starting_row_input.setObjectName("packshotRowCountInput")
        self.starting_row_input.setFixedWidth(45)
        config_row.addWidget(self.starting_row_input, 0)


        # Place Import Trackers and Apply Cleanup immediately after Starting Row
        self.btn_import_tracker = QPushButton("Import Trackers")
        self.btn_import_tracker.setObjectName("tscPrimaryBtn")
        config_row.addWidget(self.btn_import_tracker, 0)

        self.btn_apply_cleanup = QPushButton("Apply Cleanup")
        self.btn_apply_cleanup.setObjectName("tscPrimaryBtn")
        config_row.addWidget(self.btn_apply_cleanup, 0)

        self.btn_load_tsc = QPushButton("Load TSC")
        self.btn_load_tsc.setObjectName("tscPrimaryBtn")
        config_row.addWidget(self.btn_load_tsc, 0)

        config_row.addSpacing(8)

        self.import_progress = QProgressBar(self)
        self.import_progress.setObjectName("sapImportProgress")
        self.import_progress.setFixedWidth(140)
        self.import_progress.setRange(0, 100)
        self.import_progress.setValue(0)
        self.import_progress.setVisible(False)
        config_row.addWidget(self.import_progress, 0)

        config_row.addSpacing(20)

        self.btn_reset_filter = QPushButton("Reset Filter")
        self.btn_reset_filter.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_reset_filter, 0)

        self.btn_export = QPushButton("Export")
        # Make Export use the gray style like Reset Filter
        self.btn_export.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_export, 0)

        config_row.addSpacing(20)

        self.btn_generate_chart = QPushButton("Generate Chart")
        self.btn_generate_chart.setObjectName("tscPrimaryBtn")
        config_row.addWidget(self.btn_generate_chart, 0)

        config_row.addStretch(1)
        layout.addLayout(config_row)

        # --- Second row: Status group, Work Type group, Custom Search ---
        action_row = QHBoxLayout()
        action_row.setSpacing(12)

        # -- Status group (gray rounded frame) --
        status_frame = QFrame(self)
        status_frame.setObjectName("tscGroupFrame")
        status_layout = QHBoxLayout(status_frame)
        status_layout.setContentsMargins(10, 6, 10, 6)
        status_layout.setSpacing(8)

        status_label = QLabel("Status:")
        status_label.setObjectName("tscGroupLabel")
        status_layout.addWidget(status_label, 0)

        self.tsc_cb_to_do = QCheckBox("To Do")
        self.tsc_cb_in_progress = QCheckBox("In Progress")
        self.tsc_cb_completed = QCheckBox("Completed")
        self.tsc_cb_on_hold = QCheckBox("On Hold")
        self.tsc_cb_cancelled = QCheckBox("Cancelled")
        for cb in (self.tsc_cb_to_do, self.tsc_cb_in_progress, self.tsc_cb_completed,
                   self.tsc_cb_on_hold, self.tsc_cb_cancelled):
            cb.setObjectName("tscGroupCheck")
            status_layout.addWidget(cb, 0)

        action_row.addWidget(status_frame, 0)

        action_row.addSpacing(8)

        # -- Work Type group (gray rounded frame) --
        worktype_frame = QFrame(self)
        worktype_frame.setObjectName("tscGroupFrame")
        worktype_layout = QHBoxLayout(worktype_frame)
        worktype_layout.setContentsMargins(10, 6, 10, 6)
        worktype_layout.setSpacing(8)

        worktype_label = QLabel("Build Type:")
        worktype_label.setObjectName("tscGroupLabel")
        worktype_layout.addWidget(worktype_label, 0)

        self.tsc_cb_clone = QCheckBox("Clone")
        self.tsc_cb_master = QCheckBox("Master")
        self.tsc_cb_resizing = QCheckBox("Resizing")
        self.tsc_cb_admin = QCheckBox("Admin")
        self.tsc_cb_upload = QCheckBox("Upload")
        for cb in (self.tsc_cb_clone, self.tsc_cb_master, self.tsc_cb_resizing,
                   self.tsc_cb_admin, self.tsc_cb_upload):
            cb.setObjectName("tscGroupCheck")
            worktype_layout.addWidget(cb, 0)

        action_row.addWidget(worktype_frame, 0)

        action_row.addSpacing(12)

        # Custom Search label + input
        self.tsc_custom_value_label = QLabel("Custom Search")
        self.tsc_custom_value_label.setObjectName("tscCustomValueLabel")
        action_row.addWidget(self.tsc_custom_value_label, 0)
        self.tsc_custom_value_input = QLineEdit()
        self.tsc_custom_value_input.setObjectName("tscCustomValueInput")
        self.tsc_custom_value_input.setFixedWidth(156)
        action_row.addWidget(self.tsc_custom_value_input, 0)

        action_row.addStretch(1)
        layout.addLayout(action_row)

        # --- Table with filtering ---
        col_count = len(self.TSC_HEADERS)
        self._column_filters: dict[int, set[str] | None] = {i: None for i in range(col_count)}
        self.table = _ClipboardTableWidget(5, col_count, self)
        self.table.setObjectName("packshotClipboardTable")
        self._refresh_header_labels()
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.horizontalHeader().setDefaultSectionSize(130)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionsClickable(True)
        self.table.horizontalHeader().sectionClicked.connect(self._on_header_clicked)
        self.table.horizontalHeader().sectionDoubleClicked.connect(self._on_filter_header_clicked)

        for row in range(5):
            for col in range(col_count):
                self.table.setItem(row, col, QTableWidgetItem(""))

        self._undo_stack: list[list[list[str]]] = []
        self._is_restoring_undo = False
        self._is_batch_edit = False

        layout.addWidget(self.table)

        self.btn_update_rows.clicked.connect(self._on_update_rows_clicked)
        self.btn_reset_table.clicked.connect(self._on_reset_table_clicked)
        self.btn_delete_row.clicked.connect(self._on_delete_rows_clicked)
        self.btn_undo_table.clicked.connect(self._on_undo_table_clicked)
        self.btn_reset_filter.clicked.connect(self._on_reset_filter_clicked)
        self.table.itemChanged.connect(self._on_table_item_changed)

        # Connect Status / Build Type checkboxes and Custom Search to filtering
        for _cb in (
            self.tsc_cb_to_do, self.tsc_cb_in_progress, self.tsc_cb_completed,
            self.tsc_cb_on_hold, self.tsc_cb_cancelled,
            self.tsc_cb_clone, self.tsc_cb_master, self.tsc_cb_resizing,
            self.tsc_cb_admin, self.tsc_cb_upload,
        ):
            _cb.stateChanged.connect(self._apply_filters)
        self.tsc_custom_value_input.textChanged.connect(self._apply_filters)
        self.btn_generate_chart.clicked.connect(self._on_generate_chart_clicked)

        # Flash Import Trackers, Apply Cleanup, and Load TSC buttons to #D02752 on click
        for _flash_btn in (self.btn_import_tracker, self.btn_apply_cleanup, self.btn_load_tsc):
            _flash_btn.clicked.connect(
                lambda _checked=False, b=_flash_btn: self._flash_primary_btn(b)
            )

        self._push_undo_snapshot(force=True)
        self._update_row_count_display()

        self.setStyleSheet(
            """
            QDialog {
                background-color: #F4F4F4;
            }

            QTableWidget#packshotClipboardTable {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                gridline-color: #B8B8B8;
                selection-background-color: #DCE6F5;
                selection-color: #111111;
                font-family: "Segoe UI";
                font-size: 12px;
            }

            QTableWidget#packshotClipboardTable QHeaderView::section {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #7D8694;
                padding: 6px 8px;
                font-weight: 700;
            }

            QTableWidget#packshotClipboardTable QTableCornerButton::section {
                background-color: #111F35;
                border: 1px solid #7D8694;
            }

            QLabel#packshotTableTitle {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 18px;
                font-weight: 800;
            }

            QLabel#tscCountLabel {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 600;
            }

            QLabel#packshotRowCountLabel {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 600;
            }

            QLineEdit#packshotRowCountInput {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 8px;
                font-family: "Segoe UI";
                font-size: 12px;
            }

            QProgressBar#sapImportProgress {
                background-color: #FFFFFF;
                border: 1px solid #A9A9A9;
                border-radius: 7px;
                text-align: center;
                color: #111111;
                font-family: "Segoe UI";
                font-size: 11px;
                min-height: 24px;
            }

            QProgressBar#sapImportProgress::chunk {
                background-color: #8A244B;
                border-radius: 6px;
            }

            QPushButton#packshotUpdateRowsBtn {
                background-color: #9EA3AB;
                color: #000000;
                border: 1px solid #8B9098;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }

            QPushButton#packshotUpdateRowsBtn:pressed {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #111F35;
            }

            QPushButton#mapperReformatBtn {
                background-color: #8A244B;
                color: #FFFFFF;
                border: 1px solid #8A244B;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }

            QPushButton#mapperReformatBtn:pressed {
                background-color: #D02752;
                color: #FFFFFF;
                border: 1px solid #D02752;
            }

            /* Primary TSC buttons: Import, Apply Cleanup, Generate Chart */
            QPushButton#tscPrimaryBtn {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #111F35;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }

            QPushButton#tscPrimaryBtn:pressed {
                background-color: #0e1726;
                color: #FFFFFF;
                border: 1px solid #0e1726;
            }

            QFrame#tscGroupFrame {
                background-color: #9EA3AB;
                border-radius: 10px;
            }

            QLabel#tscGroupLabel {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 700;
                background: transparent;
            }

            QCheckBox#tscGroupCheck {
                color: #333333;
                font-family: "Segoe UI";
                font-size: 13px;
                spacing: 8px;
                background: transparent;
            }

            QLabel#tscCustomValueLabel {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 600;
            }

            QLineEdit#tscCustomValueInput {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 8px;
                font-family: "Segoe UI";
                font-size: 12px;
            }
            """
        )

    # -- Table management helpers (same pattern as SAP window) --

    def _on_table_batch_edit_begin(self) -> None:
        self._is_batch_edit = True

    def _on_table_batch_edit_end(self) -> None:
        self._is_batch_edit = False
        self._push_undo_snapshot()

    def _push_undo_snapshot(self, force: bool = False) -> None:
        if self._is_restoring_undo or (self._is_batch_edit and not force):
            return
        snapshot: list[list[str]] = []
        for r in range(self.table.rowCount()):
            row_data: list[str] = []
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                row_data.append(item.text() if item else "")
            snapshot.append(row_data)
        if self._undo_stack and self._undo_stack[-1] == snapshot:
            return
        self._undo_stack.append(snapshot)
        if len(self._undo_stack) > 50:
            self._undo_stack = self._undo_stack[-50:]

    def _on_table_item_changed(self, item) -> None:
        self._push_undo_snapshot()

    def _on_update_rows_clicked(self) -> None:
        try:
            new_count = int(self.row_count_input.text())
        except ValueError:
            return
        new_count = max(1, min(new_count, 9999))
        col_count = self.table.columnCount()
        old_count = self.table.rowCount()
        self.table.setRowCount(new_count)
        if new_count > old_count:
            for r in range(old_count, new_count):
                for c in range(col_count):
                    self.table.setItem(r, c, QTableWidgetItem(""))
        self._push_undo_snapshot(force=True)
        self._update_row_count_display()

    def _on_reset_table_clicked(self) -> None:
        # Reset table to 5 blank rows
        self.table.blockSignals(True)
        self.table.setRowCount(5)
        col_count = self.table.columnCount()
        for r in range(5):
            for c in range(col_count):
                item = self.table.item(r, c)
                if item is None:
                    self.table.setItem(r, c, QTableWidgetItem(""))
                else:
                    item.setText("")
        self.table.blockSignals(False)

        # Reset row count input
        self.row_count_input.setText("5")

        # Uncheck all Status checkboxes
        for cb in (self.tsc_cb_to_do, self.tsc_cb_in_progress, self.tsc_cb_completed,
                   self.tsc_cb_on_hold, self.tsc_cb_cancelled):
            cb.setChecked(False)

        # Uncheck all Build Type checkboxes
        for cb in (self.tsc_cb_clone, self.tsc_cb_master, self.tsc_cb_resizing,
                   self.tsc_cb_admin, self.tsc_cb_upload):
            cb.setChecked(False)

        # Clear Custom Search input
        self.tsc_custom_value_input.clear()

        # Reset all column filters
        for column in self._column_filters:
            self._column_filters[column] = None
        self._apply_filters()

        self._push_undo_snapshot(force=True)
        self._update_row_count_display()

    def _on_delete_rows_clicked(self) -> None:
        selected = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not selected:
            return
        self._push_undo_snapshot(force=True)
        for r in selected:
            self.table.removeRow(r)
        self.row_count_input.setText(str(self.table.rowCount()))
        self._push_undo_snapshot(force=True)
        self._update_row_count_display()

    def _on_undo_table_clicked(self) -> None:
        if len(self._undo_stack) < 2:
            return
        self._undo_stack.pop()
        snapshot = self._undo_stack[-1]
        self._is_restoring_undo = True
        col_count = self.table.columnCount()
        self.table.setRowCount(len(snapshot))
        for r, row_data in enumerate(snapshot):
            for c in range(col_count):
                value = row_data[c] if c < len(row_data) else ""
                item = self.table.item(r, c)
                if item is None:
                    self.table.setItem(r, c, QTableWidgetItem(value))
                else:
                    item.setText(value)
        self.row_count_input.setText(str(self.table.rowCount()))
        self._is_restoring_undo = False
        self._update_row_count_display()

    # -- Filter / header helpers (same pattern as Reformatted table) --

    def _refresh_header_labels(self) -> None:
        labels = []
        for idx, base_label in enumerate(self.TSC_HEADERS):
            has_active_filter = self._column_filters.get(idx) is not None
            labels.append(f"{base_label} {'▾*' if has_active_filter else '▾'}")
        self.table.setHorizontalHeaderLabels(labels)

    def _on_header_clicked(self, column: int) -> None:
        self.table.selectColumn(column)

    def _on_filter_header_clicked(self, column: int) -> None:
        values: list[str] = []
        seen: set[str] = set()
        has_blank = False
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            item = self.table.item(row, column)
            cell_text = item.text().strip() if item is not None else ""
            if cell_text == "":
                has_blank = True
                continue
            if cell_text in seen:
                continue
            seen.add(cell_text)
            values.append(cell_text)

        value_pairs: list[tuple[str, str]] = [(v, v) for v in values]
        if has_blank:
            value_pairs.append((self._BLANK_FILTER, "(Blanks)"))

        popup = self._FilterPopup(value_pairs, self._column_filters.get(column), self)
        popup.move(QCursor.pos())
        if popup.exec() != QDialog.DialogCode.Accepted:
            return

        selected = popup.get_selected_values()
        if len(selected) == len(value_pairs):
            self._column_filters[column] = None
        else:
            self._column_filters[column] = selected
        self._apply_filters()

    def _apply_filters(self) -> None:
        # Determine which Status / Build Type checkbox values are active
        status_col = self.TSC_HEADERS.index("Status")
        build_type_col = self.TSC_HEADERS.index("Build Type")

        status_cb_map: dict[str, QCheckBox] = {
            "to do":       self.tsc_cb_to_do,
            "in progress": self.tsc_cb_in_progress,
            "completed":   self.tsc_cb_completed,
            "on hold":     self.tsc_cb_on_hold,
            "cancelled":   self.tsc_cb_cancelled,
        }
        build_type_cb_map: dict[str, QCheckBox] = {
            "clone":    self.tsc_cb_clone,
            "master":   self.tsc_cb_master,
            "resizing": self.tsc_cb_resizing,
            "admin":    self.tsc_cb_admin,
            "upload":   self.tsc_cb_upload,
        }

        checked_statuses    = {v for v, cb in status_cb_map.items()    if cb.isChecked()}
        checked_build_types = {v for v, cb in build_type_cb_map.items() if cb.isChecked()}
        custom_needle = self.tsc_custom_value_input.text().strip().lower()

        for row in range(self.table.rowCount()):
            row_matches = True

            # 1. Column header popup filters
            for col, filter_value in self._column_filters.items():
                if filter_value is None:
                    continue
                item = self.table.item(row, col)
                cell_text = item.text().strip() if item is not None else ""
                normalized = self._BLANK_FILTER if cell_text == "" else cell_text
                if normalized not in filter_value:
                    row_matches = False
                    break

            # 2. Status checkboxes – only active when at least one cb is checked
            if row_matches and checked_statuses:
                item = self.table.item(row, status_col)
                cell_val = item.text().strip().lower() if item is not None else ""
                if cell_val not in checked_statuses:
                    row_matches = False

            # 3. Build Type checkboxes – only active when at least one cb is checked
            if row_matches and checked_build_types:
                item = self.table.item(row, build_type_col)
                cell_val = item.text().strip().lower() if item is not None else ""
                if cell_val not in checked_build_types:
                    row_matches = False

            # 4. Custom Search – substring match across ALL columns (case-insensitive)
            if row_matches and custom_needle:
                found = any(
                    custom_needle in (self.table.item(row, c).text().lower()
                                      if self.table.item(row, c) is not None else "")
                    for c in range(self.table.columnCount())
                )
                if not found:
                    row_matches = False

            self.table.setRowHidden(row, not row_matches)

        self._refresh_header_labels()
        self._update_row_count_display()

    def _on_reset_filter_clicked(self) -> None:
        # Clear column popup filters
        for column in self._column_filters:
            self._column_filters[column] = None
        # Uncheck all Status checkboxes
        for cb in (self.tsc_cb_to_do, self.tsc_cb_in_progress, self.tsc_cb_completed,
                   self.tsc_cb_on_hold, self.tsc_cb_cancelled):
            cb.setChecked(False)
        # Uncheck all Build Type checkboxes
        for cb in (self.tsc_cb_clone, self.tsc_cb_master, self.tsc_cb_resizing,
                   self.tsc_cb_admin, self.tsc_cb_upload):
            cb.setChecked(False)
        # Clear custom search
        self.tsc_custom_value_input.clear()
        self._apply_filters()

    def _update_row_count_display(self) -> None:
        visible = sum(1 for r in range(self.table.rowCount()) if not self.table.isRowHidden(r))
        self.label_row_count_display.setText(f"Count: <b>{visible}</b>")

    def _on_generate_chart_clicked(self) -> None:
        """Collect visible table rows and open the analytics chart dialog."""
        rows: list[list[str]] = []
        for r in range(self.table.rowCount()):
            if self.table.isRowHidden(r):
                continue
            row_data = [
                (self.table.item(r, c).text() if self.table.item(r, c) is not None else "")
                for c in range(self.table.columnCount())
            ]
            rows.append(row_data)

        if not rows:
            QMessageBox.information(self, "No Data", "The table has no visible rows to chart.")
            return

        dlg = _TscChartDialog(list(self.TSC_HEADERS), rows, parent=self)
        dlg.exec()

    def _flash_primary_btn(self, btn: QPushButton) -> None:
        """Momentarily flash a tscPrimaryBtn to #D02752 on click."""
        btn.setStyleSheet(
            "QPushButton { background-color: #D02752; color: #FFFFFF; "
            "border: 1px solid #D02752; border-radius: 8px; min-height: 30px; "
            "padding: 0 12px; font-family: 'Segoe UI'; font-size: 12px; font-weight: 600; }"
        )
        QTimer.singleShot(250, lambda: btn.setStyleSheet(""))


class _TscChartDialog(QDialog):
    """Analytics chart window for the Tracker Information table."""

    _PALETTE = [
        "#8A244B", "#1F6FAE", "#2E9E5B", "#E69C2F", "#6B4DB5",
        "#D44C37", "#3AABB5", "#A5785D", "#C557A8", "#5C7D3E",
    ]

    _PALETTE_1 = ["#32175c","#003d8f","#0062b8","#0086cd","#00a8ca","#00c9b1","#00e689","#7cff58"]
    _PALETTE_2 = ["#5c0000","#760c25","#8b2047","#9b376d","#a45193","#a66cba","#9f88de","#90a4ff"]
    _PALETTE_3 = ["#4348b9","#9745af","#ca4a9e","#eb5c8b","#ff797b","#ff9b73","#ffbd76","#ffde88"]

    _COUNT_VARIABLES = [
        "SBU",
        "Status",
        "Build Type",
        "Year",
    ]

    _COMBO_VARIABLES = [
        "SBU",
        "Status",
        "Build Type",
        "Year",
    ]

    _COMBO_SEG_BASE = [
        "SBU",
        "Status",
        "Build Type",
        "Year",
    ]

    _TOTAL_COST_ALLOWED_VARS = {"SBU", "Build Type", "Status", "Year"}

    _COL_ALIAS: dict[str, str] = {"Year": "Actual Hand-In Date"}

    # Keep for backward compatibility
    _VARIABLES = _COMBO_VARIABLES

    _COUNT_CHART_TYPES = [
        "Pie Chart",
        "Bar Chart (h)",
        "Bar Chart (v)",
        "Donut Chart",
    ]

    _COMBO_CHART_TYPES = [
        "Stacked Bar (v)",
        "Stacked Bar (h)",
        "Grouped Bar (v)",
        "Grouped Bar (h)",
    ]

    def __init__(
        self,
        headers: list[str],
        table_data: list[list[str]],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Analytics")
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowMinimizeButtonHint
        )
        self.resize(1180, 700)
        self.setMinimumSize(900, 540)
        self._headers = headers
        self._data = table_data
        self._setup_ui()
        self._refresh_chart()

    # ------------------------------------------------------------------ UI --

    def _setup_ui(self) -> None:
        self.setStyleSheet("""
            QDialog  { background-color: #0F1B2E; }
            QWidget#sidebar {
                background-color: #111F35;
                border-right: 1px solid #1E3050;
            }
            QLabel#sideLabel {
                color: #A0B0C4;
                font-family: "Segoe UI"; font-size: 11px; font-weight: 600;
                text-transform: uppercase; letter-spacing: 1px;
            }
            QLabel#sideTitle {
                color: #FFFFFF;
                font-family: "Segoe UI"; font-size: 15px; font-weight: 800;
                padding-bottom: 6px;
            }
            QLabel#sideFieldLabel {
                color: #C8D5E4;
                font-family: "Segoe UI"; font-size: 12px; font-weight: 600;
            }
            QRadioButton#chartModeRadio {
                color: #FFFFFF;
                font-family: "Segoe UI"; font-size: 12px; font-weight: 600;
                spacing: 8px;
            }
            QRadioButton#chartModeRadio::indicator {
                width: 15px; height: 15px;
                border-radius: 8px;
                border: 2px solid #7A8BA0;
                background: transparent;
            }
            QRadioButton#chartModeRadio::indicator:checked {
                border: 2px solid #8A244B;
                background: #8A244B;
            }
            QComboBox#sideCombo {
                background-color: #1C2E47;
                color: #FFFFFF;
                border: 1px solid #2D4260;
                border-radius: 6px;
                min-height: 30px;
                padding: 0 10px;
                font-family: "Segoe UI"; font-size: 12px;
            }
            QComboBox#sideCombo:hover { border: 1px solid #8A244B; }
            QComboBox#sideCombo::drop-down { border: none; width: 22px; }
            QComboBox#sideCombo QAbstractItemView {
                background-color: #1C2E47;
                color: #FFFFFF;
                selection-background-color: #8A244B;
                border: 1px solid #2D4260;
                outline: none;
            }
            QCheckBox#sideCheck {
                color: #C8D5E4;
                font-family: "Segoe UI"; font-size: 13px;
                spacing: 8px;
            }
            QFrame#sideDivider {
                background: #1E3050; border: none;
                max-height: 1px; min-height: 1px;
            }
            QLabel#chartTitle {
                color: #FFFFFF;
                font-family: "Segoe UI"; font-size: 16px; font-weight: 700;
            }
            QLabel#chartSubtitle {
                color: #7A8BA0;
                font-family: "Segoe UI"; font-size: 11px;
            }
            QPushButton#chartSaveBtn {
                background-color: #8A244B; color: #FFFFFF;
                border: none; border-radius: 8px;
                min-height: 32px; padding: 0 20px;
                font-family: "Segoe UI"; font-size: 12px; font-weight: 600;
            }
            QPushButton#chartSaveBtn:hover   { background-color: #A82D5C; }
            QPushButton#chartSaveBtn:pressed { background-color: #6E1A3C; }
            QPushButton#chartCloseBtn {
                background-color: #2D3E58; color: #FFFFFF;
                border: none; border-radius: 8px;
                min-height: 32px; padding: 0 20px;
                font-family: "Segoe UI"; font-size: 12px; font-weight: 600;
            }
            QPushButton#chartCloseBtn:hover { background-color: #374D6B; }
            QRadioButton#textColorRadio {
                color: #C8D5E4;
                font-family: "Segoe UI"; font-size: 12px; font-weight: 500;
                spacing: 8px;
            }
            QRadioButton#textColorRadio::indicator {
                width: 13px; height: 13px;
                border-radius: 7px;
                border: 2px solid #7A8BA0;
                background: transparent;
            }
            QRadioButton#textColorRadio::indicator:checked {
                border: 2px solid #8A244B;
                background: #8A244B;
            }
            QPushButton#colorSwatchBtn {
                border: 2px solid #FFFFFF;
                border-radius: 4px;
                min-width: 28px; max-width: 28px;
                min-height: 20px; max-height: 20px;
            }
            QRadioButton#paletteRadio {
                color: #C8D5E4;
                font-family: "Segoe UI"; font-size: 12px;
                spacing: 8px;
            }
            QRadioButton#paletteRadio::indicator {
                width: 13px; height: 13px;
                border-radius: 7px;
                border: 2px solid #7A8BA0;
                background: transparent;
            }
            QRadioButton#paletteRadio::indicator:checked {
                border: 2px solid #8A244B;
                background: #8A244B;
            }
        """)

        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Left sidebar ─────────────────────────────────────────────────
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(240)
        sb = QVBoxLayout(sidebar)
        sb.setContentsMargins(16, 18, 16, 18)
        sb.setSpacing(0)

        # Title
        title_lbl = QLabel("Analytics")
        title_lbl.setObjectName("sideTitle")
        sb.addWidget(title_lbl)

        # ── Mode section label
        mode_sect = QLabel("MODE")
        mode_sect.setObjectName("sideLabel")
        mode_sect.setContentsMargins(0, 10, 0, 6)
        sb.addWidget(mode_sect)

        self._radio_count = QRadioButton("Count Chart")
        self._radio_count.setObjectName("chartModeRadio")
        self._radio_count.setChecked(True)
        self._radio_combo = QRadioButton("Combination Chart")
        self._radio_combo.setObjectName("chartModeRadio")
        mode_group = QButtonGroup(self)
        mode_group.addButton(self._radio_count)
        mode_group.addButton(self._radio_combo)
        sb.addWidget(self._radio_count)
        sb.addSpacing(4)
        sb.addWidget(self._radio_combo)

        # Divider
        div1 = QFrame()
        div1.setObjectName("sideDivider")
        div1.setContentsMargins(0, 14, 0, 14)
        sb.addSpacing(14)
        sb.addWidget(div1)
        sb.addSpacing(14)

        # ── Variable section
        var_sect = QLabel("VARIABLES")
        var_sect.setObjectName("sideLabel")
        var_sect.setContentsMargins(0, 0, 0, 6)
        sb.addWidget(var_sect)

        var_lbl = QLabel("Variable:")
        var_lbl.setObjectName("sideFieldLabel")
        sb.addWidget(var_lbl)
        self._var_combo = QComboBox()
        self._var_combo.setObjectName("sideCombo")
        self._var_combo.addItems(self._COUNT_VARIABLES)  # Count mode is default
        sb.addSpacing(4)
        sb.addWidget(self._var_combo)

        # Segment-by row (Combination only)
        sb.addSpacing(10)
        self._seg_lbl = QLabel("Segment by:")
        self._seg_lbl.setObjectName("sideFieldLabel")
        sb.addWidget(self._seg_lbl)
        self._seg_combo = QComboBox()
        self._seg_combo.setObjectName("sideCombo")
        self._seg_combo.addItems(self._COMBO_SEG_BASE)  # Total Cost added dynamically
        self._seg_combo.setCurrentIndex(2)   # default: Status
        sb.addSpacing(4)
        sb.addWidget(self._seg_combo)

        # Divider
        div2 = QFrame()
        div2.setObjectName("sideDivider")
        sb.addSpacing(14)
        sb.addWidget(div2)
        sb.addSpacing(14)

        # ── Chart Type section
        ct_sect = QLabel("CHART TYPE")
        ct_sect.setObjectName("sideLabel")
        ct_sect.setContentsMargins(0, 0, 0, 6)
        sb.addWidget(ct_sect)

        ct_lbl = QLabel("Chart Type:")
        ct_lbl.setObjectName("sideFieldLabel")
        sb.addWidget(ct_lbl)
        self._chart_type_combo = QComboBox()
        self._chart_type_combo.setObjectName("sideCombo")
        self._chart_type_combo.addItems(self._COUNT_CHART_TYPES)
        # Default for Count Chart should be Donut Chart
        try:
            self._chart_type_combo.setCurrentIndex(self._COUNT_CHART_TYPES.index("Donut Chart"))
        except ValueError:
            pass
        sb.addSpacing(4)
        sb.addWidget(self._chart_type_combo)

        # Divider
        div3 = QFrame()
        div3.setObjectName("sideDivider")
        sb.addSpacing(14)
        sb.addWidget(div3)
        sb.addSpacing(10)

        # ── Options section
        opts_sect = QLabel("OPTIONS")
        opts_sect.setObjectName("sideLabel")
        opts_sect.setContentsMargins(0, 0, 0, 8)
        sb.addWidget(opts_sect)

        self._include_blanks_cb = QCheckBox("Include blanks")
        self._include_blanks_cb.setObjectName("sideCheck")
        sb.addWidget(self._include_blanks_cb)

        # Theme is always "palette"; _theme_combo kept as internal state for _colors()
        self._theme_combo = QComboBox()
        self._theme_combo.addItems(["palette"])

        # Divider between OPTIONS and TEXT COLOR
        div_tc = QFrame()
        div_tc.setObjectName("sideDivider")
        sb.addSpacing(14)
        sb.addWidget(div_tc)
        sb.addSpacing(10)

        # TEXT COLOR
        tc_sect = QLabel("TEXT COLOR")
        tc_sect.setObjectName("sideLabel")
        tc_sect.setContentsMargins(0, 0, 0, 6)
        sb.addWidget(tc_sect)

        self._tc_radio_white = QRadioButton("White")
        self._tc_radio_white.setObjectName("textColorRadio")
        self._tc_radio_white.setChecked(True)
        self._tc_radio_black = QRadioButton("Black")
        self._tc_radio_black.setObjectName("textColorRadio")
        tc_group = QButtonGroup(self)
        tc_group.addButton(self._tc_radio_white)
        tc_group.addButton(self._tc_radio_black)
        sb.addWidget(self._tc_radio_white)
        sb.addSpacing(4)
        sb.addWidget(self._tc_radio_black)

        # stash initial single-hue / divergent colors
        self._single_hue_color = "#8A244B"
        self._divergent_color1 = "#488f31"
        self._divergent_color2 = "#de425b"
        self._palette_choice   = 1

        sb.addStretch(1)
        root.addWidget(sidebar)

        # Thin divider
        vdiv = QFrame()
        vdiv.setFrameShape(QFrame.Shape.VLine)
        vdiv.setFixedWidth(1)
        vdiv.setStyleSheet("background:#1E3050; border:none;")
        root.addWidget(vdiv)

        # ── Right chart panel ─────────────────────────────────────────────
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(20, 18, 20, 14)
        right_layout.setSpacing(10)

        # Title row
        title_row = QHBoxLayout()
        self._title_lbl = QLabel("")
        self._title_lbl.setObjectName("chartTitle")
        self._subtitle_lbl = QLabel("")
        self._subtitle_lbl.setObjectName("chartSubtitle")
        title_row.addWidget(self._title_lbl)
        title_row.addStretch(1)
        title_row.addWidget(self._subtitle_lbl)
        right_layout.addLayout(title_row)

        # Matplotlib canvas – wrapped in QScrollArea so tall charts are scrollable
        self._figure = Figure(facecolor="#FFFFFF")
        self._canvas = FigureCanvasQTAgg(self._figure)
        self._canvas.setMinimumHeight(450)
        self._chart_scroll = QScrollArea()
        self._chart_scroll.setWidgetResizable(True)
        self._chart_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._chart_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._chart_scroll.setWidget(self._canvas)
        self._chart_scroll.setStyleSheet(
            "QScrollArea { background:#FFFFFF; border:1px solid #D0D6DF; border-radius:8px; }"
            "QScrollBar:vertical { background:#E8E8E8; width:8px; margin:0; border-radius:4px; }"
            "QScrollBar::handle:vertical { background:#BBBBBB; border-radius:4px; min-height:24px; }"
            "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height:0px; }"
        )
        right_layout.addWidget(self._chart_scroll, 1)

        # ── Bottom dynamic panel (color pickers / palette radios) ─────────
        self._bottom_panel = QWidget()
        bp = QHBoxLayout(self._bottom_panel)
        bp.setContentsMargins(0, 0, 0, 0)
        bp.setSpacing(12)

        # single-hue: 1 swatch
        self._sh_swatch = QPushButton()
        self._sh_swatch.setObjectName("colorSwatchBtn")
        self._sh_swatch.setFixedSize(34, 24)
        self._sh_swatch.setStyleSheet("background:#8A244B; border:2px solid #FFFFFF; border-radius:4px;")
        self._sh_label = QLabel("Hue color:")
        self._sh_label.setObjectName("sideFieldLabel")

        # divergent: 2 swatches
        self._dv_swatch1 = QPushButton()
        self._dv_swatch1.setObjectName("colorSwatchBtn")
        self._dv_swatch1.setFixedSize(34, 24)
        self._dv_swatch1.setStyleSheet("background:#488f31; border:2px solid #FFFFFF; border-radius:4px;")
        self._dv_label1 = QLabel("Color A:")
        self._dv_label1.setObjectName("sideFieldLabel")
        self._dv_swatch2 = QPushButton()
        self._dv_swatch2.setObjectName("colorSwatchBtn")
        self._dv_swatch2.setFixedSize(34, 24)
        self._dv_swatch2.setStyleSheet("background:#de425b; border:2px solid #FFFFFF; border-radius:4px;")
        self._dv_label2 = QLabel("Color B:")
        self._dv_label2.setObjectName("sideFieldLabel")

        # palette: 3 radio buttons
        self._pal_radio1 = QRadioButton("palette-1")
        self._pal_radio1.setObjectName("paletteRadio")
        self._pal_radio1.setChecked(True)
        self._pal_radio2 = QRadioButton("palette-2")
        self._pal_radio2.setObjectName("paletteRadio")
        self._pal_radio3 = QRadioButton("palette-3")
        self._pal_radio3.setObjectName("paletteRadio")
        pal_grp = QButtonGroup(self)
        pal_grp.addButton(self._pal_radio1)
        pal_grp.addButton(self._pal_radio2)
        pal_grp.addButton(self._pal_radio3)

        # all hidden initially; shown by _on_theme_changed
        for w in (self._sh_label, self._sh_swatch,
                  self._dv_label1, self._dv_swatch1, self._dv_label2, self._dv_swatch2,
                  self._pal_radio1, self._pal_radio2, self._pal_radio3):
            bp.addWidget(w)
            w.hide()
        bp.addStretch(1)
        # Fixed height prevents the canvas from resizing when bottom panel contents change
        self._bottom_panel.setFixedHeight(36)
        right_layout.addWidget(self._bottom_panel)

        # Button bar
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        save_btn = QPushButton("Save Chart")
        save_btn.setObjectName("chartSaveBtn")
        close_btn = QPushButton("Close")
        close_btn.setObjectName("chartCloseBtn")
        btn_row.addWidget(save_btn)
        btn_row.addSpacing(8)
        btn_row.addWidget(close_btn)
        right_layout.addLayout(btn_row)

        root.addWidget(right, 1)

        # ── Wire up signals ───────────────────────────────────────────────
        self._radio_count.toggled.connect(self._on_mode_changed)
        self._var_combo.currentIndexChanged.connect(self._on_var_changed)
        self._seg_combo.currentIndexChanged.connect(self._refresh_chart)
        self._chart_type_combo.currentIndexChanged.connect(self._refresh_chart)
        self._include_blanks_cb.stateChanged.connect(self._refresh_chart)
        self._theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        self._tc_radio_white.toggled.connect(self._refresh_chart)
        self._tc_radio_black.toggled.connect(self._refresh_chart)
        self._sh_swatch.clicked.connect(self._on_pick_single_hue)
        self._dv_swatch1.clicked.connect(lambda: self._on_pick_divergent(1))
        self._dv_swatch2.clicked.connect(lambda: self._on_pick_divergent(2))
        self._pal_radio1.toggled.connect(lambda checked: self._on_palette_radio(1) if checked else None)
        self._pal_radio2.toggled.connect(lambda checked: self._on_palette_radio(2) if checked else None)
        self._pal_radio3.toggled.connect(lambda checked: self._on_palette_radio(3) if checked else None)
        save_btn.clicked.connect(self._on_save_chart)
        close_btn.clicked.connect(self.reject)

        # Initial mode state
        self._on_mode_changed()
        self._on_theme_changed()

    # ---------------------------------------------------------------- slots --

    def _update_seg_combo(self) -> None:
        """Rebuild Segment-by items based on current Variable — adds Total Cost only when allowed."""
        var = self._var_combo.currentText()
        current_seg = self._seg_combo.currentText()
        new_items = list(self._COMBO_SEG_BASE)
        if var in self._TOTAL_COST_ALLOWED_VARS:
            new_items.append("Total Cost")
        self._seg_combo.blockSignals(True)
        self._seg_combo.clear()
        self._seg_combo.addItems(new_items)
        idx = self._seg_combo.findText(current_seg)
        if idx >= 0:
            self._seg_combo.setCurrentIndex(idx)
        self._seg_combo.blockSignals(False)

    def _on_var_changed(self) -> None:
        """Called when Variable combo changes — updates Segment-by list then refreshes chart."""
        if not self._radio_count.isChecked():
            self._update_seg_combo()
        self._refresh_chart()

    def _resolve_header(self, col_name: str) -> str:
        """Resolve a chart variable display name to its actual table header."""
        return self._COL_ALIAS.get(col_name, col_name)

    def _apply_col_transform(self, col_name: str, val: str) -> str:
        """Apply column-specific display transform (date → year for Actual Hand-In Date)."""
        actual = self._COL_ALIAS.get(col_name, col_name)
        if actual == "Actual Hand-In Date" and val and len(val) >= 4:
            return val[:4]
        return val

    def _resize_canvas(self, n_items: int = 0, px_per_item: int = 34, min_px: int = 450) -> None:
        """Resize canvas height so a vertical scrollbar appears when many items are charted."""
        h_px = max(min_px, n_items * px_per_item + 130) if n_items > 0 else min_px
        self._canvas.setMinimumHeight(h_px)
        # Always size the figure to the height that will actually be displayed.
        # With setWidgetResizable(True) the canvas fills the viewport when h_px ≤
        # viewport_h, so we use max(h_px, viewport_h) as the figure height.  This
        # prevents the chart from rendering at a smaller-than-viewport height and
        # leaving blank white space below (which previously required a
        # minimize/maximize to trigger resizeEvent and correct the size).
        viewport_h = self._chart_scroll.viewport().height()
        figure_h = max(h_px, viewport_h) if viewport_h > 0 else h_px
        dpi = self._figure.dpi
        # figure_h is in logical (device-independent) pixels.  On HiDPI screens
        # the figure DPI may already be multiplied by the device pixel ratio, so
        # we must convert logical → physical pixels before dividing by dpi.
        dpr = self._canvas.devicePixelRatioF()
        self._figure.set_size_inches(self._figure.get_figwidth(), figure_h * dpr / dpi)
        self._canvas.updateGeometry()

    def _on_mode_changed(self) -> None:
        is_count = self._radio_count.isChecked()
        # Swap variable list – Count mode excludes Total Cost
        current_var = self._var_combo.currentText()
        self._var_combo.blockSignals(True)
        self._var_combo.clear()
        self._var_combo.addItems(self._COUNT_VARIABLES if is_count else self._COMBO_VARIABLES)
        idx = self._var_combo.findText(current_var)
        if idx >= 0:
            self._var_combo.setCurrentIndex(idx)
        self._var_combo.blockSignals(False)
        # Show/hide segment-by controls
        self._seg_lbl.setVisible(not is_count)
        self._seg_combo.setVisible(not is_count)
        # Rebuild segment-by items for new mode
        if not is_count:
            self._update_seg_combo()
        # Swap chart-type list
        self._chart_type_combo.blockSignals(True)
        self._chart_type_combo.clear()
        self._chart_type_combo.addItems(
            self._COUNT_CHART_TYPES if is_count else self._COMBO_CHART_TYPES
        )
        self._chart_type_combo.blockSignals(False)
        # Default to Donut Chart when in Count mode
        if is_count:
            try:
                self._chart_type_combo.setCurrentIndex(self._COUNT_CHART_TYPES.index("Donut Chart"))
            except ValueError:
                pass
        self._refresh_chart()

    def _refresh_chart(self) -> None:
        self._figure.clear()
        var      = self._var_combo.currentText()
        seg      = self._seg_combo.currentText()
        ctype    = self._chart_type_combo.currentText()
        blanks   = self._include_blanks_cb.isChecked()
        is_count = self._radio_count.isChecked()

        n_rows = len(self._data)
        self._subtitle_lbl.setText(f"Based on {n_rows} visible row(s)")

        if is_count:
            self._title_lbl.setText(f"{var}  —  {ctype}")
            if ctype == "Pie Chart":
                self._draw_pie(var, blanks, donut=False)
            elif ctype == "Donut Chart":
                self._draw_pie(var, blanks, donut=True)
            elif ctype == "Bar Chart (h)":
                self._draw_bar(var, blanks, horizontal=True)
            elif ctype == "Bar Chart (v)":
                self._draw_bar(var, blanks, horizontal=False)
            elif ctype == "Histogram":
                self._draw_histogram(var, blanks)
        else:
            seg = self._seg_combo.currentText()
            if seg == "Total Cost":
                self._title_lbl.setText(f"Sum of Total Cost by {var}  —  {ctype}")
                if "(h)" in ctype:
                    self._draw_sum_bar(var, seg, blanks, horizontal=True)
                else:
                    self._draw_sum_bar(var, seg, blanks, horizontal=False)
            else:
                self._title_lbl.setText(f"{seg} by {var}  —  {ctype}")
                if "(h)" in ctype:
                    self._draw_stacked(var, seg, blanks, horizontal=True,
                                       grouped="Grouped" in ctype)
                else:
                    self._draw_stacked(var, seg, blanks, horizontal=False,
                                       grouped="Grouped" in ctype)

        self._canvas.draw()
        self._canvas.update()

    # -------------------------------------------------------- helpers --------

    def _col_values(self, col_name: str, include_blanks: bool = False) -> list[str]:
        actual = self._resolve_header(col_name)
        if actual not in self._headers:
            return []
        idx = self._headers.index(actual)
        values = []
        for row in self._data:
            v = row[idx].strip() if idx < len(row) else ""
            v = self._apply_col_transform(col_name, v)
            if v or include_blanks:
                values.append(v if v else "(blank)")
        return values

    def _value_counts(self, col_name: str, include_blanks: bool = False) -> dict[str, int]:
        from collections import Counter
        return dict(Counter(self._col_values(col_name, include_blanks)).most_common())

    # ------------------------------------------------------ theme helpers --

    @staticmethod
    def _hex_to_rgb(h: str) -> tuple[int, int, int]:
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    @staticmethod
    def _rgb_to_hex(r: int, g: int, b: int) -> str:
        return f"#{r:02x}{g:02x}{b:02x}"

    @classmethod
    def _make_single_hue_colors(cls, base_hex: str, n: int) -> list[str]:
        """Produce n shades blending from base_hex toward near-white."""
        if n <= 0:
            return []
        r0, g0, b0 = cls._hex_to_rgb(base_hex)
        # target: a very light tint
        r1, g1, b1 = 255, 220, 240
        result = []
        for i in range(n):
            t = i / max(n - 1, 1)
            r = round(r0 + (r1 - r0) * t)
            g = round(g0 + (g1 - g0) * t)
            b = round(b0 + (b1 - b0) * t)
            result.append(cls._rgb_to_hex(r, g, b))
        return result

    @classmethod
    def _make_divergent_colors(cls, hex1: str, hex2: str, n: int) -> list[str]:
        """Produce n shades from hex1 → white/light-grey → hex2."""
        if n <= 0:
            return []
        r0, g0, b0 = cls._hex_to_rgb(hex1)
        r2, g2, b2 = cls._hex_to_rgb(hex2)
        mid_r, mid_g, mid_b = 241, 241, 241  # near-white midpoint
        result = []
        for i in range(n):
            t = i / max(n - 1, 1)  # 0 → 1
            if t <= 0.5:
                s = t / 0.5
                r = round(r0 + (mid_r - r0) * s)
                g = round(g0 + (mid_g - g0) * s)
                b = round(b0 + (mid_b - b0) * s)
            else:
                s = (t - 0.5) / 0.5
                r = round(mid_r + (r2 - mid_r) * s)
                g = round(mid_g + (g2 - mid_g) * s)
                b = round(mid_b + (b2 - mid_b) * s)
            result.append(cls._rgb_to_hex(r, g, b))
        return result

    def _colors(self, n: int) -> list[str]:
        theme = self._theme_combo.currentText()
        if theme == "single hue":
            return self._make_single_hue_colors(self._single_hue_color, n)
        if theme == "divergent":
            return self._make_divergent_colors(self._divergent_color1, self._divergent_color2, n)
        # palette (also the default fallback)
        p = {1: self._PALETTE_1, 2: self._PALETTE_2, 3: self._PALETTE_3}.get(self._palette_choice, self._PALETTE_1)
        return (p * ((n // len(p)) + 1))[:n]

    def _no_data(self) -> None:
        ax = self._figure.add_subplot(111)
        ax.text(0.5, 0.5, "No data available for this chart",
                ha="center", va="center", fontsize=13,
                color="#AAAAAA", transform=ax.transAxes)
        ax.axis("off")

    # -------------------------------------------------------- chart renderers

    def _draw_pie(self, col: str, blanks: bool, donut: bool = False) -> None:
        counts = self._value_counts(col, blanks)
        if not counts:
            self._resize_canvas(0)
            self._no_data()
            return
        labels = list(counts.keys())
        values = list(counts.values())
        # Fixed canvas height keeps pie/donut the same size regardless of variable
        self._resize_canvas(0)
        ax = self._figure.add_subplot(111)
        # Move pie/donut leftwards to leave space for details/legend
        try:
            ax.set_position([0.05, 0.12, 0.62, 0.8])
        except Exception:
            pass
        colors = self._colors(len(labels))
        wedgeprops = {"linewidth": 1.4, "edgecolor": "#FFFFFF"}
        if donut:
            wedgeprops["width"] = 0.52
        wedges, _tx, autotexts = ax.pie(
            values, labels=None, colors=colors,
            autopct=lambda p: f"{p:.1f}%" if p >= 2 else "",
            startangle=140, pctdistance=0.78 if not donut else 0.82,
            wedgeprops=wedgeprops,
        )
        pct_color = "#FFFFFF" if self._tc_radio_white.isChecked() else "#111111"
        for at in autotexts:
            at.set_fontsize(9)
            at.set_color(pct_color)
            at.set_fontweight("bold")
        if donut:
            total = sum(values)
            ax.text(0, 0, str(total), ha="center", va="center",
                    fontsize=16, fontweight="bold", color="#333333")
        legend_labels = [f"{lbl}  ({v})" for lbl, v in zip(labels, values)]
        ax.legend(wedges, legend_labels,
                  loc="center left", bbox_to_anchor=(0.93, 0.5),
                  fontsize=9, frameon=True, framealpha=0.95, edgecolor="#CCCCCC")
        ax.set_title(f"{col} Distribution  (total {sum(values)})",
                     fontsize=13, fontweight="bold", pad=14)

    def _on_theme_changed(self) -> None:
        """Show/hide bottom-panel controls — panel has fixed height so canvas never resizes."""
        theme = self._theme_combo.currentText()
        # single-hue controls
        for w in (self._sh_label, self._sh_swatch):
            w.setVisible(theme == "single hue")
        # divergent controls
        for w in (self._dv_label1, self._dv_swatch1, self._dv_label2, self._dv_swatch2):
            w.setVisible(theme == "divergent")
        # palette controls
        for w in (self._pal_radio1, self._pal_radio2, self._pal_radio3):
            w.setVisible(theme == "palette")
        self._refresh_chart()

    def _on_pick_single_hue(self) -> None:
        color = QColorDialog.getColor(QColor(self._single_hue_color), self, "Pick single hue color")
        if color.isValid():
            self._single_hue_color = color.name()
            self._sh_swatch.setStyleSheet(
                f"background:{self._single_hue_color}; border:2px solid #FFFFFF; border-radius:4px;"
            )
            self._refresh_chart()

    def _on_pick_divergent(self, which: int) -> None:
        start = self._divergent_color1 if which == 1 else self._divergent_color2
        color = QColorDialog.getColor(QColor(start), self, f"Pick divergent color {'A' if which==1 else 'B'}")
        if color.isValid():
            hexc = color.name()
            if which == 1:
                self._divergent_color1 = hexc
                self._dv_swatch1.setStyleSheet(
                    f"background:{hexc}; border:2px solid #FFFFFF; border-radius:4px;"
                )
            else:
                self._divergent_color2 = hexc
                self._dv_swatch2.setStyleSheet(
                    f"background:{hexc}; border:2px solid #FFFFFF; border-radius:4px;"
                )
            self._refresh_chart()

    def _on_palette_radio(self, choice: int) -> None:
        self._palette_choice = choice
        self._refresh_chart()

    def _draw_bar(self, col: str, blanks: bool, horizontal: bool = True,
                  max_items: int = 25) -> None:
        counts = self._value_counts(col, blanks)
        if not counts:
            self._no_data()
            return
        items = sorted(counts.items(), key=lambda x: -x[1])[:max_items]
        labels = [k for k, _ in items]
        values = [v for _, v in items]
        self._resize_canvas(len(labels) if horizontal else 0)
        colors = self._colors(len(labels))
        ax = self._figure.add_subplot(111)
        if horizontal:
            y_pos = list(range(len(labels)))
            ax.barh(y_pos, values[::-1], color=colors[::-1],
                    edgecolor="#FFFFFF", linewidth=0.5, height=0.65)
            ax.set_yticks(y_pos)
            ax.set_yticklabels(labels[::-1], fontsize=9)
            for i, val in enumerate(values[::-1]):
                ax.text(val + max(values) * 0.01, i, str(val),
                        va="center", ha="left", fontsize=9, color="#333333")
            ax.set_xlabel("Count", fontsize=10)
            ax.set_xlim(0, max(values) * 1.2)
            ax.tick_params(axis="y", length=0)
        else:
            x_pos = list(range(len(labels)))
            ax.bar(x_pos, values, color=colors,
                   edgecolor="#FFFFFF", linewidth=0.5, width=0.65)
            ax.set_xticks(x_pos)
            ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=9)
            for i, val in enumerate(values):
                ax.text(i, val + max(values) * 0.01, str(val),
                        ha="center", va="bottom", fontsize=9, color="#333333")
            ax.set_ylabel("Count", fontsize=10)
            ax.set_ylim(0, max(values) * 1.15)
            ax.tick_params(axis="x", length=0)
        ax.set_title(f"{col} Count", fontsize=13, fontweight="bold", pad=14)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        if horizontal:
            self._figure.subplots_adjust(left=0.22, right=0.88, top=0.92, bottom=0.08)
        else:
            self._figure.subplots_adjust(left=0.10, right=0.95, top=0.92, bottom=0.22)

    def _draw_histogram(self, col: str, blanks: bool) -> None:
        """Numeric histogram – falls back to bar chart if values are non-numeric."""
        self._resize_canvas(0)
        raw = self._col_values(col, blanks)
        if not raw:
            self._no_data()
            return
        numeric = []
        for v in raw:
            try:
                numeric.append(float(v))
            except ValueError:
                pass
        ax = self._figure.add_subplot(111)
        if numeric:
            ax.hist(numeric, bins="auto", color="#8A244B",
                    edgecolor="#FFFFFF", linewidth=0.6)
            ax.set_xlabel(col, fontsize=10)
            ax.set_ylabel("Frequency", fontsize=10)
        else:
            # Non-numeric – fall back to vertical bar
            from collections import Counter
            counts = dict(Counter(raw).most_common(25))
            labels = list(counts.keys())
            values = list(counts.values())
            colors = self._colors(len(labels))
            ax.bar(range(len(labels)), values, color=colors,
                   edgecolor="#FFFFFF", linewidth=0.5, width=0.65)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=9)
            ax.set_ylabel("Count", fontsize=10)
            ax.tick_params(axis="x", length=0)
        ax.set_title(f"{col} Histogram", fontsize=13, fontweight="bold", pad=14)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        self._figure.subplots_adjust(left=0.10, right=0.95, top=0.92, bottom=0.22)

    def _draw_sum_bar(self, x_col: str, sum_col: str, blanks: bool, horizontal: bool) -> None:
        """Draw a bar chart where bar height = sum of sum_col values per x_col category."""
        _ax = self._resolve_header(x_col)
        _as = self._resolve_header(sum_col)
        xi = self._headers.index(_ax) if _ax in self._headers else -1
        si = self._headers.index(_as) if _as in self._headers else -1
        if xi < 0 or si < 0 or not self._data:
            self._no_data()
            return

        sums: dict[str, float] = defaultdict(float)
        for row in self._data:
            xv = row[xi].strip() if xi < len(row) else ""
            xv = self._apply_col_transform(x_col, xv)
            sv = row[si].strip() if si < len(row) else ""
            if not blanks and not xv:
                continue
            xv = xv or "(blank)"
            try:
                sums[xv] += float(sv)
            except (ValueError, TypeError):
                sums[xv] += 0.0  # keep group even if value is non-numeric

        if not sums:
            self._no_data()
            return

        items = sorted(sums.items(), key=lambda x: -x[1])
        labels = [k for k, _ in items]
        values = [v for _, v in items]
        max_val = max(values) if values else 1

        self._resize_canvas(len(labels) if horizontal else 0)
        colors = self._colors(len(labels))
        ax = self._figure.add_subplot(111)

        if horizontal:
            y_pos = list(range(len(labels)))
            ax.barh(y_pos, values[::-1], color=colors[::-1],
                    edgecolor="#FFFFFF", linewidth=0.5, height=0.65)
            ax.set_yticks(y_pos)
            ax.set_yticklabels(labels[::-1], fontsize=9)
            for i, val in enumerate(values[::-1]):
                ax.text(val + max_val * 0.01, i, f"{val:,.2f}",
                        va="center", ha="left", fontsize=9, color="#333333")
            ax.set_xlabel(f"Total {sum_col}", fontsize=10)
            ax.set_xlim(0, max_val * 1.22)
            ax.tick_params(axis="y", length=0)
        else:
            x_pos = list(range(len(labels)))
            ax.bar(x_pos, values, color=colors,
                   edgecolor="#FFFFFF", linewidth=0.5, width=0.65)
            ax.set_xticks(x_pos)
            ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=9)
            for i, val in enumerate(values):
                ax.text(i, val + max_val * 0.01, f"{val:,.2f}",
                        ha="center", va="bottom", fontsize=9, color="#333333")
            ax.set_ylabel(f"Total {sum_col}", fontsize=10)
            ax.set_ylim(0, max_val * 1.18)
            ax.tick_params(axis="x", length=0)

        ax.set_title(f"Sum of {sum_col} by {x_col}  (total {sum(values):,.2f})",
                     fontsize=13, fontweight="bold", pad=14)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        if horizontal:
            self._figure.subplots_adjust(left=0.22, right=0.88, top=0.92, bottom=0.08)
        else:
            self._figure.subplots_adjust(left=0.10, right=0.95, top=0.92, bottom=0.22)

    def _draw_stacked(
        self, x_col: str, seg_col: str, blanks: bool,
        horizontal: bool = False, grouped: bool = False,
        max_x: int = 18,
    ) -> None:
        _ax = self._resolve_header(x_col)
        _as = self._resolve_header(seg_col)
        xi = self._headers.index(_ax) if _ax in self._headers else -1
        si = self._headers.index(_as) if _as in self._headers else -1
        if xi < 0 or si < 0 or not self._data:
            self._no_data()
            return

        matrix: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
        for row in self._data:
            xv = row[xi].strip() if xi < len(row) else ""
            xv = self._apply_col_transform(x_col, xv)
            sv = row[si].strip() if si < len(row) else ""
            sv = self._apply_col_transform(seg_col, sv)
            if not blanks and (not xv or not sv):
                continue
            xv = xv or "(blank)"
            sv = sv or "(blank)"
            matrix[xv][sv] += 1

        if not matrix:
            self._no_data()
            return

        x_labels = sorted(matrix.keys(), key=lambda k: -sum(matrix[k].values()))[:max_x]
        all_segs = sorted({s for m in matrix.values() for s in m})
        self._resize_canvas(len(x_labels) if horizontal else 0)
        colors = self._colors(len(all_segs))
        ax = self._figure.add_subplot(111)

        n_segs = len(all_segs)
        n_x    = len(x_labels)
        bar_w  = 0.65 if not grouped else max(0.12, 0.65 / max(n_segs, 1))

        if grouped:
            offsets = [(i - (n_segs - 1) / 2) * bar_w for i in range(n_segs)]
            for (seg, color), offset in zip(zip(all_segs, colors), offsets):
                vals = [matrix[x].get(seg, 0) for x in x_labels]
                if horizontal:
                    ax.barh([p + offset for p in range(n_x)], vals,
                            color=color, edgecolor="#FFFFFF", linewidth=0.4,
                            label=seg, height=bar_w)
                else:
                    ax.bar([p + offset for p in range(n_x)], vals,
                           color=color, edgecolor="#FFFFFF", linewidth=0.4,
                           label=seg, width=bar_w)
        else:
            bottoms = [0] * n_x
            for seg, color in zip(all_segs, colors):
                vals = [matrix[x].get(seg, 0) for x in x_labels]
                if horizontal:
                    ax.barh(range(n_x), vals, left=bottoms,
                            color=color, edgecolor="#FFFFFF", linewidth=0.4,
                            label=seg, height=0.68)
                else:
                    ax.bar(range(n_x), vals, bottom=bottoms,
                           color=color, edgecolor="#FFFFFF", linewidth=0.4,
                           label=seg, width=0.68)
                bottoms = [b + v for b, v in zip(bottoms, vals)]

        if horizontal:
            ax.set_yticks(range(n_x))
            ax.set_yticklabels(x_labels, fontsize=9)
            ax.set_xlabel("Count", fontsize=10)
            ax.tick_params(axis="y", length=0)
        else:
            ax.set_xticks(range(n_x))
            ax.set_xticklabels(x_labels, rotation=30, ha="right", fontsize=9)
            ax.set_ylabel("Count", fontsize=10)
            ax.tick_params(axis="x", length=0)

        ax.legend(loc="upper right", fontsize=9,
                  framealpha=0.95, edgecolor="#CCCCCC")
        ax.set_title(f"{seg_col} by {x_col}",
                     fontsize=13, fontweight="bold", pad=14)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        if horizontal:
            self._figure.subplots_adjust(left=0.22, right=0.82, top=0.92, bottom=0.08)
        else:
            self._figure.subplots_adjust(left=0.10, right=0.88, top=0.92, bottom=0.22)

    def _on_save_chart(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Chart", "",
            "PNG Image (*.png);;PDF Document (*.pdf);;SVG Image (*.svg)",
        )
        if path:
            self._figure.savefig(
                path, dpi=450, bbox_inches="tight", facecolor="white"
            )


class _ReviewTableDialog(QDialog):
    """Review Table — loads the 'body_map_IDH' sheet from an SDC or Review excel file."""

    _BLANK_FILTER   = "__BLANK__"
    _SHEET_NAME     = "body_map_IDH"
    _HBM_COL        = "Head Bom Mat"
    _ACCENT_HEADERS = {
        "Head Bom Mat", "Master IDH Build", "Hit Type",
        # Assessment Comments — canonical + common spelling variants
        "Assessment Comments", "Assesment Comments", "Asessment Comments",
        "Asesment Comments", "Comments", "Coments",
    }

    # Hit Type cell rules (key = lowercase cell value)
    _HIT_TYPE_RULES = {
        "total hit":   {"bg": "#C3CC9B", "fg": None,      "bold": False},
        "partial hit": {"bg": "#F08D39", "fg": None,      "bold": False},
        "0 hit":       {"bg": None,      "fg": "#CC0000", "bold": True},
    }
    # Master IDH Build cell rules
    _IDH_BUILD_RULES = {
        "3d": {"bg": "#C3CC9B", "fg": None, "bold": False},
        "2d": {"bg": "#F08D39", "fg": None, "bold": False},
    }
    _IDH_BUILD_NA_BG = "#9EA3AB"

    # ── nested filter popup (same as _TscOption1TableDialog) ──────────────
    class _FilterPopup(QDialog):
        def __init__(self, values, selected_values, parent=None):
            super().__init__(parent)
            self.setWindowFlags(Qt.WindowType.Popup)
            self.setMinimumSize(260, 330)
            self.resize(260, 330)
            self._is_syncing = False
            self._value_items: list = []
            layout = QVBoxLayout(self)
            layout.setContentsMargins(4, 4, 4, 4)
            layout.setSpacing(4)
            self.search_input = QLineEdit(self)
            self.search_input.setPlaceholderText("Search")
            self.search_input.setClearButtonEnabled(True)
            self.search_input.setObjectName("packshotRowCountInput")
            layout.addWidget(self.search_input)
            self.list_widget = QListWidget(self)
            self.list_widget.setSelectionMode(QListWidget.SelectionMode.NoSelection)
            layout.addWidget(self.list_widget, 1)
            self.item_select_all = QListWidgetItem("(Select All)")
            self.item_select_all.setFlags(self.item_select_all.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            self.item_select_all.setCheckState(Qt.CheckState.Checked)
            self.item_select_all.setData(Qt.ItemDataRole.UserRole, "__SELECT_ALL__")
            self.list_widget.addItem(self.item_select_all)
            selected = selected_values if selected_values is not None else {v for v, _ in values}
            for value, label in values:
                item = QListWidgetItem(label)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if value in selected else Qt.CheckState.Unchecked)
                item.setData(Qt.ItemDataRole.UserRole, value)
                self.list_widget.addItem(item)
                self._value_items.append((item, value, label))
            btn_row = QHBoxLayout()
            btn_row.addStretch(1)
            self.btn_apply = QPushButton("Apply")
            self.btn_apply.setObjectName("packshotUpdateRowsBtn")
            self.btn_cancel = QPushButton("Cancel")
            self.btn_cancel.setObjectName("packshotUpdateRowsBtn")
            btn_row.addWidget(self.btn_apply)
            btn_row.addWidget(self.btn_cancel)
            layout.addLayout(btn_row)
            self.search_input.textChanged.connect(self._apply_search)
            self.list_widget.itemChanged.connect(self._on_item_changed)
            self.btn_apply.clicked.connect(self.accept)
            self.btn_cancel.clicked.connect(self.reject)
            self.search_input.setFocus()
            self._sync_select_all_state()

        def _apply_search(self, text):
            needle = text.strip().lower()
            for item, _v, label in self._value_items:
                item.setHidden(needle not in label.lower())
            self._sync_select_all_state()

        def _on_item_changed(self, item):
            if self._is_syncing:
                return
            if item.data(Qt.ItemDataRole.UserRole) == "__SELECT_ALL__":
                self._is_syncing = True
                try:
                    state = item.checkState()
                    for vi, _v, _l in self._value_items:
                        if not vi.isHidden():
                            vi.setCheckState(state)
                finally:
                    self._is_syncing = False
                return
            self._sync_select_all_state()

        def _sync_select_all_state(self):
            visible = [i for i, _v, _l in self._value_items if not i.isHidden()]
            if not visible:
                state = Qt.CheckState.Unchecked
            elif all(i.checkState() == Qt.CheckState.Checked for i in visible):
                state = Qt.CheckState.Checked
            elif any(i.checkState() == Qt.CheckState.Checked for i in visible):
                state = Qt.CheckState.PartiallyChecked
            else:
                state = Qt.CheckState.Unchecked
            self._is_syncing = True
            try:
                self.item_select_all.setCheckState(state)
            finally:
                self._is_syncing = False

        def get_selected_values(self):
            return {v for item, v, _l in self._value_items if item.checkState() == Qt.CheckState.Checked}

    # ── nested: HBM lookup button delegate ────────────────────────────────
    class _HBMButtonDelegate(QStyledItemDelegate):
        """Draws a small ▶ button on the right edge of every cell in the HBM column.
        Clicking it calls on_lookup(hbm_value)."""
        _BTN_W = 22

        def __init__(self, on_lookup, parent=None):
            super().__init__(parent)
            self._on_lookup = on_lookup

        def paint(self, painter, option, index):
            # Paint cell text in a reduced rect (leave room for button)
            opt = QStyleOptionViewItem(option)
            self.initStyleOption(opt, index)
            opt.rect = option.rect.adjusted(0, 0, -self._BTN_W, 0)
            super().paint(painter, opt, index)
            # Draw small accent button on the right
            btn_rect = self._btn_rect(option.rect)
            painter.save()
            painter.setRenderHint(painter.RenderHint.Antialiasing)
            painter.setBrush(QColor("#D02752"))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(btn_rect, 3, 3)
            painter.setPen(QColor("#FFFFFF"))
            font = painter.font()
            font.setPixelSize(10)
            painter.setFont(font)
            painter.drawText(btn_rect, Qt.AlignmentFlag.AlignCenter, "▶")
            painter.restore()

        def editorEvent(self, event, model, option, index):
            if event.type() == QEvent.Type.MouseButtonRelease:
                if self._btn_rect(option.rect).contains(event.pos().toPoint()
                        if hasattr(event.pos(), "toPoint") else event.pos()):
                    val = index.data(Qt.ItemDataRole.DisplayRole) or ""
                    self._on_lookup(val.strip())
                    return True
            return super().editorEvent(event, model, option, index)

        @staticmethod
        def _btn_rect(cell_rect):
            return QRect(
                cell_rect.right() - _ReviewTableDialog._HBMButtonDelegate._BTN_W + 2,
                cell_rect.top() + 3,
                _ReviewTableDialog._HBMButtonDelegate._BTN_W - 4,
                cell_rect.height() - 6,
            )

    # ── nested: accent header view ─────────────────────────────────────────
    class _AccentHeaderView(QHeaderView):
        """Paints accent columns with #D02752 background, white bold text; normal columns unchanged."""
        _ACCENT_BG   = QColor("#D02752")
        _ACCENT_FG   = QColor("#FFFFFF")
        _NORMAL_BG   = QColor("#111F35")
        _BORDER      = QColor("#7D8694")

        def __init__(self, highlight_cols: set, parent=None):
            super().__init__(Qt.Orientation.Horizontal, parent)
            self._highlight_cols: set[int] = set(highlight_cols)
            self._btn_reserves: dict[int, int] = {}  # col -> px to reserve on right
            self._user_header_fmt: dict[int, dict] = {}  # col -> {"bg": "#hex", "fg": "#hex"}
            self.setSectionsClickable(True)

        def set_highlight_cols(self, cols: set) -> None:
            self._highlight_cols = set(cols)
            self.viewport().update()

        def set_section_btn_reserve(self, col: int, reserve_px: int) -> None:
            """Reserve right-side pixels in col's header text area for an overlaid button."""
            self._btn_reserves[col] = reserve_px
            self.viewport().update()

        def update_header_colors(self, fmt: dict) -> None:
            """Apply user-defined header colors; triggers repaint."""
            self._user_header_fmt = dict(fmt)
            self.viewport().update()

        def paintSection(self, painter, rect, logical_index):
            user_fmt = self._user_header_fmt.get(logical_index, {})
            is_accent = logical_index in self._highlight_cols

            if not is_accent and not user_fmt:
                super().paintSection(painter, rect, logical_index)
                return

            # Fully custom paint
            painter.save()

            # Background
            if user_fmt.get("bg"):
                bg_color = QColor(user_fmt["bg"])
            elif is_accent:
                bg_color = self._ACCENT_BG
            else:
                bg_color = self._NORMAL_BG
            painter.fillRect(rect, bg_color)
            painter.setPen(self._BORDER)
            painter.drawRect(rect.adjusted(0, 0, -1, -1))

            label = self.model().headerData(
                logical_index, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole
            )
            if label:
                font = painter.font()
                font.setBold(True)
                painter.setFont(font)
                # Foreground
                if user_fmt.get("fg"):
                    painter.setPen(QColor(user_fmt["fg"]))
                else:
                    painter.setPen(self._ACCENT_FG)
                right_margin = self._btn_reserves.get(logical_index, 0) + 4
                text_rect = rect.adjusted(8, 0, -right_margin, 0)
                painter.drawText(
                    text_rect,
                    int(Qt.AlignmentFlag.AlignCenter),
                    str(label),
                )
            painter.restore()

    # ── nested: Check Duplicate dialog ────────────────────────────────────
    class _CheckDuplicateDialog(QDialog):
        """Popup for configuring and running a duplicate check."""

        def __init__(self, parent=None) -> None:
            super().__init__(parent)
            self.setWindowTitle("Check Duplicate")
            self.resize(440, 320)
            self.setMinimumSize(380, 280)
            self._selected_tracker_files: list[str] = []

            root = QVBoxLayout(self)
            root.setContentsMargins(20, 20, 20, 20)
            root.setSpacing(12)

            # ── title ──────────────────────────────────────────────────────
            title = QLabel("Check Duplicate")
            title.setObjectName("_cdTitle")
            root.addWidget(title)

            # ── compare with self ──────────────────────────────────────────
            self.chk_compare_self = QCheckBox("Compare with self")
            self.chk_compare_self.setObjectName("_cdCheck")
            root.addWidget(self.chk_compare_self)

            # ── compare with other trackers ────────────────────────────────
            self.chk_compare_others = QCheckBox("Compare with other trackers")
            self.chk_compare_others.setObjectName("_cdCheck")
            root.addWidget(self.chk_compare_others)

            # select trackers row (indented, under 2nd checkbox)
            tracker_row = QHBoxLayout()
            tracker_row.setContentsMargins(22, 0, 0, 0)
            tracker_row.setSpacing(6)
            self.btn_select_trackers = QPushButton("Select Trackers")
            self.btn_select_trackers.setObjectName("_cdSecBtn")
            self.btn_select_trackers.setEnabled(False)
            tracker_row.addWidget(self.btn_select_trackers, 0)
            self.edit_selected_trackers = QLineEdit()
            self.edit_selected_trackers.setObjectName("_cdEdit")
            self.edit_selected_trackers.setReadOnly(True)
            self.edit_selected_trackers.setPlaceholderText("No trackers selected")
            tracker_row.addWidget(self.edit_selected_trackers, 1)
            root.addLayout(tracker_row)

            # ── field for checking ─────────────────────────────────────────
            field_row = QHBoxLayout()
            field_row.setSpacing(8)
            field_lbl = QLabel("Field for checking:")
            field_lbl.setObjectName("_cdLabel")
            field_row.addWidget(field_lbl, 0)
            self.combo_field = QComboBox()
            self.combo_field.setObjectName("_cdCombo")
            self.combo_field.addItems(["IDH", "Basic"])
            field_row.addWidget(self.combo_field, 1)
            root.addLayout(field_row)

            root.addStretch(1)

            # ── run button ─────────────────────────────────────────────────
            run_row = QHBoxLayout()
            run_row.addStretch(1)
            self.btn_run = QPushButton("Run Duplicate Check")
            self.btn_run.setObjectName("_cdRunBtn")
            self.btn_run.setFixedHeight(50)
            self.btn_run.setMinimumWidth(200)
            run_row.addWidget(self.btn_run)
            run_row.addStretch(1)
            root.addLayout(run_row)

            # ── signals ────────────────────────────────────────────────────
            self.chk_compare_others.toggled.connect(
                self.btn_select_trackers.setEnabled)
            self.btn_select_trackers.clicked.connect(self._on_select_trackers)

            self._apply_style()

        def _on_select_trackers(self) -> None:
            from pathlib import Path
            files, _ = QFileDialog.getOpenFileNames(
                self, "Select Tracker Files", "",
                "Excel files (*.xlsx *.xls)")
            if files:
                self._selected_tracker_files = files
                names = ", ".join(Path(f).name for f in files)
                self.edit_selected_trackers.setText(names)

        def _apply_style(self) -> None:
            self.setStyleSheet(
                """
                QDialog { background-color: #F4F4F4; }
                QLabel#_cdTitle {
                    color: #111F35; font-family: 'Segoe UI';
                    font-size: 16px; font-weight: 800;
                }
                QLabel#_cdLabel {
                    color: #111F35; font-family: 'Segoe UI';
                    font-size: 13px; font-weight: 600;
                }
                QCheckBox#_cdCheck {
                    color: #111F35; font-family: 'Segoe UI';
                    font-size: 13px; font-weight: 500;
                    spacing: 8px;
                }
                QCheckBox#_cdCheck::indicator {
                    width: 16px; height: 16px;
                    border: 1px solid #9EA3AB; border-radius: 4px;
                    background: #FFFFFF;
                }
                QCheckBox#_cdCheck::indicator:checked {
                    background: #111F35; border: 1px solid #111F35;
                }
                QPushButton#_cdSecBtn {
                    background-color: #9EA3AB; color: #000000;
                    border: 1px solid #8B9098; border-radius: 8px;
                    min-height: 30px; padding: 0 12px;
                    font-family: 'Segoe UI'; font-size: 12px; font-weight: 600;
                }
                QPushButton#_cdSecBtn:disabled {
                    background-color: #D0D3D8; color: #888888;
                    border: 1px solid #C0C3C8;
                }
                QPushButton#_cdSecBtn:pressed {
                    background-color: #111F35; color: #FFFFFF;
                }
                QLineEdit#_cdEdit {
                    background-color: #FFFFFF; color: #111111;
                    border: 1px solid #A9A9A9; border-radius: 8px;
                    min-height: 30px; padding: 0 8px;
                    font-family: 'Segoe UI'; font-size: 12px;
                }
                QComboBox#_cdCombo {
                    background-color: #FFFFFF; color: #111111;
                    border: 1px solid #A9A9A9; border-radius: 8px;
                    min-height: 30px; padding: 0 8px;
                    font-family: 'Segoe UI'; font-size: 12px;
                }
                QComboBox#_cdCombo::drop-down { border: none; width: 24px; }
                QPushButton#_cdRunBtn {
                    background-color: #111F35; color: #ffffff;
                    border: none; border-radius: 12px;
                    padding: 0 18px; font-family: 'Segoe UI';
                    font-size: 16px; font-weight: 700;
                }
                QPushButton#_cdRunBtn:pressed {
                    background-color: #D02752; color: #ffffff;
                }
                """
            )

    # ── nested: Check Missing IDHs dialog ─────────────────────────────────
    class _CheckMissingIDHsDialog(QDialog):
        """Popup for selecting a tracker and running the missing-IDH check."""

        def __init__(self, project_review_folder: str = "", initial_tracker: str = "", parent=None) -> None:
            super().__init__(parent)
            self.setWindowTitle("Check Missing IDHs")
            self.resize(480, 210)
            self.setMinimumSize(380, 190)
            self._tracker_file: str = initial_tracker

            root = QVBoxLayout(self)
            root.setContentsMargins(20, 20, 20, 20)
            root.setSpacing(12)

            # ── title ──────────────────────────────────────────────────────
            title = QLabel("Check Missing IDHs")
            title.setObjectName("_cdTitle")
            root.addWidget(title)

            # ── select tracker ─────────────────────────────────────────────
            tracker_row = QHBoxLayout()
            tracker_row.setSpacing(6)
            self.btn_select_tracker = QPushButton("Select Tracker")
            self.btn_select_tracker.setObjectName("_cdSecBtn")
            tracker_row.addWidget(self.btn_select_tracker, 0)
            self.edit_tracker = QLineEdit()
            self.edit_tracker.setObjectName("_cdEdit")
            self.edit_tracker.setReadOnly(True)
            self.edit_tracker.setPlaceholderText("No tracker selected")
            if initial_tracker:
                from pathlib import Path as _P
                self.edit_tracker.setText(_P(initial_tracker).name)
                self.edit_tracker.setToolTip(initial_tracker)
            tracker_row.addWidget(self.edit_tracker, 1)
            root.addLayout(tracker_row)

            root.addStretch(1)

            # ── run button ─────────────────────────────────────────────────
            run_row = QHBoxLayout()
            run_row.addStretch(1)
            self.btn_run = QPushButton("Check Missing IDHs")
            self.btn_run.setObjectName("_cdRunBtn")
            self.btn_run.setFixedHeight(50)
            self.btn_run.setMinimumWidth(220)
            run_row.addWidget(self.btn_run)
            run_row.addStretch(1)
            root.addLayout(run_row)

            # ── signals ────────────────────────────────────────────────────
            self.btn_select_tracker.clicked.connect(self._on_select_tracker)
            self.btn_run.clicked.connect(self._on_run_clicked)

            self._apply_style()

        @property
        def tracker_file(self) -> str:
            return self._tracker_file

        def _on_select_tracker(self) -> None:
            from pathlib import Path as _P
            start = str(_P(self._tracker_file).parent) if self._tracker_file else ""
            file, _ = QFileDialog.getOpenFileName(
                self, "Select Tracker File", start,
                "Excel files (*.xlsx *.xls *.xlsm)"
            )
            if not file:
                return
            if "tracker" not in _P(file).stem.lower():
                QMessageBox.warning(
                    self, "Invalid File",
                    "The selected file does not have 'tracker' in its name.\n"
                    "Please select a file with 'tracker' in the filename."
                )
                return
            self._tracker_file = file
            self.edit_tracker.setText(_P(file).name)
            self.edit_tracker.setToolTip(file)

        def _on_run_clicked(self) -> None:
            if not self._tracker_file:
                QMessageBox.warning(self, "No Tracker", "Please select a tracker file first.")
                return
            self.accept()

        def _apply_style(self) -> None:
            self.setStyleSheet(
                """
                QDialog { background-color: #F4F4F4; }
                QLabel#_cdTitle {
                    color: #111F35; font-family: 'Segoe UI';
                    font-size: 16px; font-weight: 800;
                }
                QPushButton#_cdSecBtn {
                    background-color: #9EA3AB; color: #000000;
                    border: 1px solid #8B9098; border-radius: 8px;
                    min-height: 30px; padding: 0 12px;
                    font-family: 'Segoe UI'; font-size: 12px; font-weight: 600;
                }
                QPushButton#_cdSecBtn:pressed {
                    background-color: #111F35; color: #FFFFFF;
                }
                QLineEdit#_cdEdit {
                    background-color: #FFFFFF; color: #111111;
                    border: 1px solid #A9A9A9; border-radius: 8px;
                    min-height: 30px; padding: 0 8px;
                    font-family: 'Segoe UI'; font-size: 12px;
                }
                QPushButton#_cdRunBtn {
                    background-color: #111F35; color: #ffffff;
                    border: none; border-radius: 12px;
                    padding: 0 18px; font-family: 'Segoe UI';
                    font-size: 16px; font-weight: 700;
                }
                QPushButton#_cdRunBtn:pressed {
                    background-color: #D02752; color: #ffffff;
                }
                """
            )

    # ── constructor ────────────────────────────────────────────────────────
    def __init__(self, source_file: str, sdc_folder: str = "", project_review_folder: str = "", tracker_files: list | None = None, load_colors: bool = False, parent=None) -> None:
        super().__init__(parent)
        from pathlib import Path
        self._source_file = source_file
        self._sdc_folder = sdc_folder
        self._project_review_folder = project_review_folder
        self._tracker_files: list[str] = tracker_files or []
        self._headers: list[str] = []
        self._original_snapshot: list[list[str]] = []
        self._original_headers: list[str] = []
        self._undo_stack: list[list[list[str]]] = []
        self._is_batch_edit = False
        self._is_restoring_undo = False
        self._column_filters: dict[int, set[str] | None] = {}
        # user-applied cell formatting: (row, col) -> {"bg": str|None, "fg": str|None, "bold": bool|None}
        self._user_fmt: dict[tuple[int, int], dict] = {}
        # user-applied header formatting: col -> {"bg": str|None, "fg": str|None}
        self._header_fmt: dict[int, dict] = {}
        self._tracker_info_dialogs: list = []  # keep refs to non-modal Tracker Info windows

        file_name = Path(source_file).name if source_file else "No file"
        self.setWindowTitle(f"Review Table — {file_name}")
        # Allow maximize / restore to full screen
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowMinimizeButtonHint
        )
        self.resize(1400, 700)
        self.setMinimumSize(900, 500)

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(10)

        # ── title row ─────────────────────────────────────────────────────
        title_row = QHBoxLayout()
        title_row.setSpacing(12)
        title_lbl = QLabel("Review Table")
        title_lbl.setObjectName("packshotTableTitle")
        title_row.addWidget(title_lbl, 0)
        self.lbl_file_name = QLabel(file_name)
        self.lbl_file_name.setObjectName("tscCountLabel")
        title_row.addWidget(self.lbl_file_name, 0)
        title_row.addStretch(1)
        self.lbl_row_count = QLabel("Rows: <b>0</b>")
        self.lbl_row_count.setObjectName("tscCountLabel")
        self.lbl_row_count.setTextFormat(Qt.TextFormat.RichText)
        title_row.addWidget(self.lbl_row_count, 0)
        root_layout.addLayout(title_row)

        # ── main toolbar ──────────────────────────────────────────────────
        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        self.btn_rt_delete       = QPushButton("Delete")
        self.btn_rt_reset        = QPushButton("Reset")
        self.btn_rt_undo         = QPushButton("Undo")
        self.btn_rt_reset_filter = QPushButton("Reset Filter")
        self.btn_rt_add_row      = QPushButton("Add Row")
        self.btn_rt_add_col      = QPushButton("Add Column")
        for _b in (self.btn_rt_delete, self.btn_rt_reset, self.btn_rt_undo,
                   self.btn_rt_reset_filter, self.btn_rt_add_row, self.btn_rt_add_col):
            _b.setObjectName("packshotUpdateRowsBtn")
            toolbar.addWidget(_b, 0)

        toolbar.addSpacing(16)
        toolbar.addStretch(1)

        self.btn_rt_check_missing_idhs = QPushButton("Check Missing IDHs")
        self.btn_rt_check_missing_idhs.setObjectName("tscPrimaryBtn")
        toolbar.addWidget(self.btn_rt_check_missing_idhs, 0)

        self.btn_rt_import_bma = QPushButton("Import another BMA")
        self.btn_rt_import_bma.setObjectName("tscPrimaryBtn")
        toolbar.addWidget(self.btn_rt_import_bma, 0)

        self.btn_rt_save = QPushButton("Save")
        self.btn_rt_save.setObjectName("tscPrimaryBtn")
        toolbar.addWidget(self.btn_rt_save, 0)

        root_layout.addLayout(toolbar)

        # ── formatting toolbar ────────────────────────────────────────────
        fmt_toolbar = QHBoxLayout()
        fmt_toolbar.setSpacing(8)
        fmt_lbl = QLabel("Format:")
        fmt_lbl.setObjectName("packshotRowCountLabel")
        fmt_toolbar.addWidget(fmt_lbl, 0)
        self.btn_rt_bold           = QPushButton("Bold")
        self.btn_rt_cell_color     = QPushButton("Cell Color")
        self.btn_rt_text_color     = QPushButton("Text Color")
        self.btn_rt_header_bg      = QPushButton("Header BG")
        self.btn_rt_header_fg      = QPushButton("Header Text")
        for _b in (self.btn_rt_bold, self.btn_rt_cell_color, self.btn_rt_text_color,
                   self.btn_rt_header_bg, self.btn_rt_header_fg):
            _b.setObjectName("packshotUpdateRowsBtn")
            fmt_toolbar.addWidget(_b, 0)
        fmt_toolbar.addStretch(1)
        root_layout.addLayout(fmt_toolbar)

        # ── table ──────────────────────────────────────────────────────────
        headers, rows = self._load_sheet(source_file)
        headers, rows = self._ensure_assessment_col(headers, rows)
        self._headers = headers
        # Pre-load cell and header colors when continuing from an existing saved file
        if load_colors:
            self._user_fmt, self._header_fmt = self._load_sheet_colors(source_file, headers)
        col_count = max(len(headers), 1)
        self._column_filters = {i: None for i in range(col_count)}

        self.table = _ClipboardTableWidget(0, col_count, self)
        self.table.setObjectName("packshotClipboardTable")

        # Accent header view (replaces default horizontal header)
        self._header_view = self._AccentHeaderView(self._compute_accent_cols(), self.table)
        self._header_view.setDefaultSectionSize(140)
        self._header_view.setStretchLastSection(True)
        self._header_view.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self._header_view.sectionClicked.connect(self._on_header_clicked)
        self._header_view.sectionDoubleClicked.connect(self._on_filter_header_clicked)
        self.table.setHorizontalHeader(self._header_view)
        if self._header_fmt:
            self._header_view.update_header_colors(self._header_fmt)

        self._refresh_header_labels()
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultSectionSize(34)

        self._populate_table(rows)
        root_layout.addWidget(self.table)

        # Install HBM lookup delegate on the Head Bom Mat column
        self._install_hbm_delegate()

        # ── "Run Initial Assessment" button overlaid on Assessment Comments header ──
        self._assess_run_btn = QPushButton("clone check", self._header_view)
        self._assess_run_btn.setObjectName("rtAssessRunBtn")
        self._assess_run_btn.setFixedSize(90, 22)
        self._assess_run_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._assess_run_btn.setToolTip("Run Initial Assessment — detect clones by Basic Number")
        self._assess_run_btn.setStyleSheet(
            "QPushButton#rtAssessRunBtn {"
            "  background: rgba(255,255,255,0.20);"
            "  color: #FFFFFF;"
            "  border: 1px solid rgba(255,255,255,0.50);"
            "  border-radius: 4px;"
            "  font-size: 12px;"
            "  padding: 0;"
            "}"
            "QPushButton#rtAssessRunBtn:hover {"
            "  background: rgba(255,255,255,0.38);"
            "}"
            "QPushButton#rtAssessRunBtn:pressed {"
            "  background: rgba(255,255,255,0.60);"
            "}"
        )
        self._assess_run_btn.clicked.connect(self._on_run_initial_assessment)
        self._header_view.sectionResized.connect(lambda *_: self._position_assess_btn())
        self._header_view.geometriesChanged.connect(self._position_assess_btn)
        QTimer.singleShot(0, self._position_assess_btn)

        # Store original data for full Reset
        self._original_snapshot = [list(r) for r in self._get_table_snapshot()]
        self._original_headers  = list(self._headers)

        self._push_undo_snapshot(force=True)
        self._update_row_count()

        # ── signals ────────────────────────────────────────────────────────
        self.btn_rt_delete.clicked.connect(self._on_delete_clicked)
        self.btn_rt_reset.clicked.connect(self._on_reset_clicked)
        self.btn_rt_undo.clicked.connect(self._on_undo_clicked)
        self.btn_rt_reset_filter.clicked.connect(self._on_reset_filter_clicked)
        self.btn_rt_add_row.clicked.connect(self._on_add_row)
        self.btn_rt_add_col.clicked.connect(self._on_add_col)
        self.btn_rt_check_missing_idhs.clicked.connect(self._on_check_missing_idhs_clicked)
        self.btn_rt_import_bma.clicked.connect(self._on_import_bma_clicked)
        self.btn_rt_save.clicked.connect(self._on_save_clicked)
        self.btn_rt_bold.clicked.connect(self._on_format_bold)
        self.btn_rt_cell_color.clicked.connect(self._on_format_cell_color)
        self.btn_rt_text_color.clicked.connect(self._on_format_text_color)
        self.btn_rt_header_bg.clicked.connect(self._on_format_header_bg)
        self.btn_rt_header_fg.clicked.connect(self._on_format_header_fg)
        self.table.itemChanged.connect(self._on_item_changed)

        self._apply_stylesheet()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        QTimer.singleShot(0, self._position_assess_btn)

    def _on_check_duplicate_clicked(self) -> None:
        dlg = self._CheckDuplicateDialog(self)
        dlg.exec()

    def _on_check_missing_idhs_clicked(self) -> None:
        """Open the Check Missing IDHs dialog, then append missing entries to the table."""
        from pathlib import Path

        # Find initial tracker: scan project_review_folder for any Excel with "tracker" in stem
        initial_tracker = ""
        if self._project_review_folder:
            pr = Path(self._project_review_folder)
            if pr.is_dir():
                for f in sorted(pr.iterdir()):
                    if (f.suffix.lower() in {".xlsx", ".xls", ".xlsm"}
                            and "tracker" in f.stem.lower()):
                        initial_tracker = str(f)
                        break
        # Fall back to already-loaded tracker_files
        if not initial_tracker:
            for tf in self._tracker_files:
                if "tracker" in Path(tf).stem.lower() and Path(tf).exists():
                    initial_tracker = tf
                    break

        dlg = self._CheckMissingIDHsDialog(
            project_review_folder=self._project_review_folder,
            initial_tracker=initial_tracker,
            parent=self,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        tracker_file = dlg.tracker_file
        if not tracker_file:
            return

        # Find Head Bom Mat column index
        lower_headers = [h.strip().lower() for h in self._headers]
        hbm_idx = next((i for i, h in enumerate(lower_headers) if h == "head bom mat"), None)
        if hbm_idx is None:
            QMessageBox.warning(self, "Column Not Found",
                                "Could not find the 'Head Bom Mat' column in the table.")
            return

        # Collect existing HBM values
        existing_hbm: set[str] = set()
        for r in range(self.table.rowCount()):
            item = self.table.item(r, hbm_idx)
            if item:
                val = item.text().strip()
                if val:
                    existing_hbm.add(val)

        # Find missing IDHs
        missing = review_project.find_missing_idhs(tracker_file, existing_hbm)
        if not missing:
            QMessageBox.information(self, "No Missing IDHs",
                                    "All IDHs in the tracker are already present in the table.")
            return

        # Locate Basic Number and Hit Type column indices
        basic_idx = next(
            (i for i, h in enumerate(lower_headers)
             if h in {"basic number", "basic num", "basic no", "basic"}),
            None,
        )
        hit_type_idx = next(
            (i for i, h in enumerate(lower_headers) if h == "hit type"),
            None,
        )

        # Append missing IDHs as new rows
        self._push_undo_snapshot(force=True)
        self.table.blockSignals(True)
        col_count = self.table.columnCount()
        for idh_val in missing:
            row_pos = self.table.rowCount()
            self.table.insertRow(row_pos)
            for c in range(col_count):
                if c == hbm_idx:
                    val = idh_val
                elif basic_idx is not None and c == basic_idx:
                    val = "missing"
                elif hit_type_idx is not None and c == hit_type_idx:
                    val = "na"
                else:
                    val = ""
                item = QTableWidgetItem(val)
                self._apply_rule_formatting(item, c, val)
                self.table.setItem(row_pos, c, item)
        self.table.blockSignals(False)
        self._push_undo_snapshot(force=True)
        self._update_row_count()

        QMessageBox.information(
            self, "Done",
            f"Added {len(missing)} missing IDH(s) to the review table."
        )

    # ── sheet loader ───────────────────────────────────────────────────────
    @staticmethod
    def _load_sheet(file_path: str) -> tuple[list[str], list[list[str]]]:
        """Read body_map_IDH sheet; return (headers, data_rows) as plain strings."""
        try:
            import pandas as pd
            sheet_name = _ReviewTableDialog._SHEET_NAME
            for engine in ("calamine", "openpyxl", None):
                try:
                    kw: dict = {} if engine is None else {"engine": engine}
                    df = pd.read_excel(
                        file_path,
                        sheet_name=sheet_name,
                        dtype=str,
                        **kw,
                    )
                    df = df.fillna("")
                    return list(df.columns), [list(r) for r in df.values.tolist()]
                except Exception:
                    continue
        except Exception:
            pass
        return [], []

    @staticmethod
    def _load_sheet_colors(
        file_path: str, headers: list[str]
    ) -> tuple[dict[tuple[int, int], dict], dict[int, dict]]:
        """Read cell and header colors from a saved review xlsx (openpyxl).

        Returns:
            cell_fmt  – {(row, col): {"bg": "#rrggbb", "fg": "#rrggbb", "bold": bool}}
            header_fmt – {col: {"bg": "#rrggbb", "fg": "#rrggbb"}}
        """
        cell_fmt: dict[tuple[int, int], dict] = {}
        header_fmt: dict[int, dict] = {}

        def _parse_rgb(raw: str | None) -> str | None:
            """Convert openpyxl ARGB/RGB string → '#rrggbb', or None if default/absent."""
            if not raw:
                return None
            raw = raw.upper().lstrip("#")
            if len(raw) == 8:
                raw = raw[2:]   # strip alpha
            if len(raw) != 6:
                return None
            if raw in ("000000", "FFFFFF"):
                return None     # ignore default black/white as 'no color'
            return "#" + raw

        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheet_name = _ReviewTableDialog._SHEET_NAME
            if sheet_name not in wb.sheetnames:
                wb.close()
                return cell_fmt, header_fmt
            ws = wb[sheet_name]

            # Build xlsx-column → table-column mapping from the header row
            lower_headers = [h.strip().lower() for h in headers]
            xl_to_tbl: dict[int, int] = {}   # 0-based xlsx col → 0-based table col
            # Default accent header colors to ignore (they are not user-set)
            _ACCENT_BG_NORM = "D02752"
            _NORMAL_BG_NORM = "111F35"
            _DEFAULT_FG_NORM = "FFFFFF"

            header_xl_row = next(ws.iter_rows(min_row=1, max_row=1), None)
            if header_xl_row:
                for xl_c, cell in enumerate(header_xl_row):
                    val = str(cell.value or "").strip().lower()
                    if val in lower_headers:
                        tbl_c = lower_headers.index(val)
                        xl_to_tbl[xl_c] = tbl_c
                        # Read header color — only store if different from defaults
                        h_bg: str | None = None
                        h_fg: str | None = None
                        if cell.fill and cell.fill.fill_type == "solid":
                            raw = getattr(cell.fill.fgColor, "rgb", None)
                            raw_norm = (raw or "").upper().lstrip("#")
                            if len(raw_norm) == 8:
                                raw_norm = raw_norm[2:]
                            if raw_norm not in (_ACCENT_BG_NORM, _NORMAL_BG_NORM, "000000", "FFFFFF", ""):
                                h_bg = "#" + raw_norm if len(raw_norm) == 6 else None
                        if cell.font and cell.font.color and cell.font.color.type == "rgb":
                            raw = cell.font.color.rgb
                            raw_norm = (raw or "").upper().lstrip("#")
                            if len(raw_norm) == 8:
                                raw_norm = raw_norm[2:]
                            if raw_norm not in (_DEFAULT_FG_NORM, "000000", "FFFFFF", ""):
                                h_fg = "#" + raw_norm if len(raw_norm) == 6 else None
                        if h_bg or h_fg:
                            entry: dict = {}
                            if h_bg:
                                entry["bg"] = h_bg
                            if h_fg:
                                entry["fg"] = h_fg
                            header_fmt[tbl_c] = entry

            # Read data rows (row 2 onwards = table row 0 onwards)
            for xl_r, row_cells in enumerate(ws.iter_rows(min_row=2)):
                tbl_r = xl_r
                for xl_c, cell in enumerate(row_cells):
                    tbl_c = xl_to_tbl.get(xl_c)
                    if tbl_c is None:
                        continue
                    entry: dict = {}
                    if cell.fill and cell.fill.fill_type == "solid":
                        bg = _parse_rgb(getattr(cell.fill.fgColor, "rgb", None))
                        if bg:
                            entry["bg"] = bg
                    if cell.font:
                        if cell.font.color and cell.font.color.type == "rgb":
                            fg = _parse_rgb(cell.font.color.rgb)
                            if fg:
                                entry["fg"] = fg
                        if cell.font.bold:
                            entry["bold"] = True
                    if entry:
                        cell_fmt[(tbl_r, tbl_c)] = entry

            wb.close()
        except Exception:
            pass

        return cell_fmt, header_fmt

    # ── table helpers ──────────────────────────────────────────────────────
    def _compute_accent_cols(self) -> set[int]:
        lower_headers = [h.strip().lower() for h in self._headers]
        result: set[int] = set()
        for name in self._ACCENT_HEADERS:
            key = name.strip().lower()
            if key in lower_headers:
                result.add(lower_headers.index(key))
        return result

    def _apply_rule_formatting(self, item: "QTableWidgetItem", col_idx: int, value: str) -> None:
        """Apply rule-based background / foreground / bold to item."""
        if col_idx >= len(self._headers):
            return
        header_key = self._headers[col_idx].strip().lower()
        val_lower  = value.strip().lower()

        # Reset to defaults
        item.setBackground(Qt.GlobalColor.transparent)
        item.setForeground(QColor("#111111"))
        font = item.font()
        font.setBold(False)
        item.setFont(font)

        if header_key == "hit type":
            rule = self._HIT_TYPE_RULES.get(val_lower)
            if rule:
                if rule["bg"]:
                    item.setBackground(QColor(rule["bg"]))
                if rule["fg"]:
                    item.setForeground(QColor(rule["fg"]))
                if rule["bold"]:
                    font.setBold(True)
                    item.setFont(font)
        elif header_key == "master idh build":
            if val_lower in self._IDH_BUILD_RULES:
                rule = self._IDH_BUILD_RULES[val_lower]
                if rule["bg"]:
                    item.setBackground(QColor(rule["bg"]))
            elif "na" in val_lower and val_lower:
                item.setBackground(QColor(self._IDH_BUILD_NA_BG))

    def _apply_user_formatting(self, item: "QTableWidgetItem", r: int, c: int) -> None:
        """Apply user-stored overrides on top of rule-based styling."""
        fmt = self._user_fmt.get((r, c))
        if not fmt:
            return
        if fmt.get("bg"):
            item.setBackground(QColor(fmt["bg"]))
        if fmt.get("fg"):
            item.setForeground(QColor(fmt["fg"]))
        if fmt.get("bold") is not None:
            font = item.font()
            font.setBold(fmt["bold"])
            item.setFont(font)

    def _populate_table(self, rows: list[list[str]]) -> None:
        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        col_count = self.table.columnCount()
        for r, row_data in enumerate(rows):
            for c in range(col_count):
                val = str(row_data[c]) if c < len(row_data) else ""
                item = QTableWidgetItem(val)
                self._apply_rule_formatting(item, c, val)
                self._apply_user_formatting(item, r, c)
                self.table.setItem(r, c, item)
        self.table.blockSignals(False)

    def _get_table_snapshot(self) -> list[list[str]]:
        return [
            [(self.table.item(r, c).text() if self.table.item(r, c) else "")
             for c in range(self.table.columnCount())]
            for r in range(self.table.rowCount())
        ]

    def _refresh_header_labels(self) -> None:
        labels = []
        for idx, base in enumerate(self._headers):
            has_filter = self._column_filters.get(idx) is not None
            labels.append(f"{base} {'▾*' if has_filter else '▾'}")
        self.table.setHorizontalHeaderLabels(labels)
        # Update accent columns on the custom header view (if already created)
        if hasattr(self, "_header_view"):
            self._header_view.set_highlight_cols(self._compute_accent_cols())

    def _update_row_count(self) -> None:
        visible = sum(1 for r in range(self.table.rowCount()) if not self.table.isRowHidden(r))
        self.lbl_row_count.setText(f"Rows: <b>{visible}</b>")

    def _install_hbm_delegate(self) -> None:
        """Set the HBM button delegate on the Head Bom Mat column (if present)."""
        lower = [h.strip().lower() for h in self._headers]
        if self._HBM_COL.lower() in lower:
            col = lower.index(self._HBM_COL.lower())
            delegate = self._HBMButtonDelegate(self._on_hbm_lookup, self.table)
            self.table.setItemDelegateForColumn(col, delegate)

    # ── Assessment Comments column helpers ─────────────────────────────────
    _ASSESS_COL_VARIANTS: frozenset[str] = frozenset({
        "assessment comments", "assesment comments", "asessment comments",
        "asesment comments", "comments", "coments",
    })

    @staticmethod
    def _ensure_assessment_col(
        headers: list[str], rows: list[list[str]]
    ) -> tuple[list[str], list[list[str]]]:
        """Insert 'Assessment Comments' after 'Lib IDH' if no variant is present."""
        lower = [h.strip().lower() for h in headers]
        if any(h in _ReviewTableDialog._ASSESS_COL_VARIANTS for h in lower):
            return headers, rows
        insert_after = next((i for i, h in enumerate(lower) if h == "lib idh"), None)
        insert_pos = (insert_after + 1) if insert_after is not None else len(headers)
        headers = list(headers)
        headers.insert(insert_pos, "Assessment Comments")
        new_rows: list[list[str]] = []
        for r in rows:
            new_r = list(r)
            while len(new_r) < insert_pos:
                new_r.append("")
            new_r.insert(insert_pos, "")
            new_rows.append(new_r)
        return headers, new_rows

    def _assess_col_index(self) -> int | None:
        lower = [h.strip().lower() for h in self._headers]
        return next((i for i, h in enumerate(lower) if h in self._ASSESS_COL_VARIANTS), None)

    def _position_assess_btn(self) -> None:
        assess_col = self._assess_col_index()
        if assess_col is None:
            self._assess_run_btn.hide()
            return
        btn_w = self._assess_run_btn.width()
        btn_h = self._assess_run_btn.height()
        # Tell the header painter to leave room for the button on the right
        self._header_view.set_section_btn_reserve(assess_col, btn_w + 4)
        col_x = self._header_view.sectionViewportPosition(assess_col)
        col_w = self._header_view.sectionSize(assess_col)
        hdr_h = self._header_view.height()
        x = col_x + col_w - btn_w - 6
        y = max(0, (hdr_h - btn_h) // 2)
        self._assess_run_btn.move(x, y)
        self._assess_run_btn.raise_()
        self._assess_run_btn.show()

    def _on_run_initial_assessment(self) -> None:
        """For each Basic Number group, mark subsequent rows as 'Clone of: <HBM>'."""
        lower = [h.strip().lower() for h in self._headers]
        basic_variants = {"basic number", "basic num", "basic no", "basic"}
        basic_col = next((i for i, h in enumerate(lower) if h in basic_variants), None)
        hbm_col   = next((i for i, h in enumerate(lower) if h == "head bom mat"), None)
        assess_col = self._assess_col_index()

        if basic_col is None or assess_col is None:
            QMessageBox.warning(
                self, "Missing Columns",
                "Could not find 'Basic Number' or 'Assessment Comments' column.",
            )
            return

        self._push_undo_snapshot(force=True)
        self.table.blockSignals(True)

        # first_row_for[basic_val] = row index of first occurrence
        first_row_for: dict[str, int] = {}
        clones_written = 0

        for row in range(self.table.rowCount()):
            basic_item = self.table.item(row, basic_col)
            basic_val = basic_item.text().strip() if basic_item else ""
            if not basic_val:
                continue

            if basic_val not in first_row_for:
                first_row_for[basic_val] = row
                # Leave first occurrence untouched
            else:
                first_row = first_row_for[basic_val]
                hbm_val = ""
                if hbm_col is not None:
                    hbm_item = self.table.item(first_row, hbm_col)
                    hbm_val = hbm_item.text().strip() if hbm_item else ""
                comment = f"Clone of: {hbm_val}" if hbm_val else "Clone of: (unknown)"
                item = self.table.item(row, assess_col)
                if item is None:
                    item = QTableWidgetItem(comment)
                    self.table.setItem(row, assess_col, item)
                else:
                    item.setText(comment)
                clones_written += 1

        self.table.blockSignals(False)
        self._push_undo_snapshot(force=True)

        if clones_written:
            QMessageBox.information(
                self, "Assessment Complete",
                f"Initial assessment done — {clones_written} clone(s) marked.",
            )
        else:
            QMessageBox.information(
                self, "Assessment Complete",
                "No duplicate Basic Numbers found — no clones to mark.",
            )

    def _on_hbm_lookup(self, hbm_value: str) -> None:
        """Look up hbm_value in all tracker files and show a popup with matching row data."""
        if not hbm_value:
            QMessageBox.information(self, "Lookup", "Cell is empty — nothing to look up.")
            return
        if not self._tracker_files:
            QMessageBox.information(self, "No Trackers",
                                    "No Project Tracker files loaded.\n"
                                    "Set the Project Tracker field in the Review Project panel first.")
            return

        try:
            import pandas as pd
        except ImportError:
            QMessageBox.warning(self, "Error", "pandas is required for tracker lookup.")
            return

        IDH_ALIASES = {"idh number", "idh", "idh no", "idh no.", "idh_number"}
        results: list[tuple[str, list[str], list[str]]] = []  # (filename, headers, values)

        def _normalise_id(v: str) -> str:
            """Strip whitespace and trailing .0 so numeric IDs match regardless of type."""
            v = v.strip()
            if v.endswith(".0") and v[:-2].isdigit():
                v = v[:-2]
            return v.lower()

        def _find_header_row(xl, sheet: str) -> int | None:
            """Scan the first 20 rows to find the row whose cells contain an IDH-like header."""
            raw = xl.parse(sheet, header=None, dtype=str).fillna("")
            for row_idx in range(min(20, len(raw))):
                row_vals = [str(v).strip().lower() for v in raw.iloc[row_idx]]
                if any(v in IDH_ALIASES for v in row_vals):
                    return row_idx
            return None

        for tf in self._tracker_files:
            from pathlib import Path as _P
            if not _P(tf).exists():
                continue
            try:
                xl = pd.ExcelFile(tf, engine="calamine")
                tracker_sheets = [s for s in xl.sheet_names if "tracker" in s.lower()]
                for sheet in tracker_sheets:
                    try:
                        header_row = _find_header_row(xl, sheet)
                        if header_row is None:
                            continue
                        df = xl.parse(sheet, header=header_row, dtype=str).fillna("")
                        # Find IDH column
                        idh_col = next(
                            (c for c in df.columns if str(c).strip().lower() in IDH_ALIASES),
                            None,
                        )
                        if idh_col is None:
                            continue
                        target = _normalise_id(hbm_value)
                        match = df[df[idh_col].apply(lambda x: _normalise_id(str(x))) == target]
                        if match.empty:
                            continue
                        for _, row in match.iterrows():
                            results.append((
                                f"{_P(tf).name} / {sheet}",
                                list(df.columns),
                                [str(row[c]) for c in df.columns],
                            ))
                    except Exception:
                        continue
            except Exception:
                continue

        if not results:
            QMessageBox.information(self, "Not Found",
                                    f"No matching row found for IDH: {hbm_value}\n\n"
                                    f"Searched {len(self._tracker_files)} tracker file(s).")
            return

        # Show popup dialog with results
        _ACCENT_COLS  = {"idh number", "idh name"}
        _YELLOW_COLS  = {"packaging size", "packaging type"}
        _HDR_BG       = QColor("#111F35")
        _ACCENT_BG    = QColor("#D02752")
        _YELLOW_BG    = QColor("#FFD150")
        _HDR_FG       = QColor("#FFFFFF")
        _DARK_FG      = QColor("#000000")

        dlg = QDialog()
        dlg.setWindowTitle(f"Tracker Info — {hbm_value}")
        dlg.setMinimumWidth(800)
        dlg.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        dlg.setWindowFlags(
            dlg.windowFlags()
            | Qt.WindowType.WindowMaximizeButtonHint
            | Qt.WindowType.WindowMinimizeButtonHint
        )
        v = QVBoxLayout(dlg)
        v.setContentsMargins(0, 0, 0, 10)
        v.setSpacing(0)

        for source_label, headers, values in results:
            tbl = QTableWidget(1, len(headers), dlg)
            tbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            tbl.verticalHeader().setVisible(False)
            tbl.horizontalHeader().setStretchLastSection(True)
            tbl.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
            tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
            tbl.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
            tbl.setShowGrid(True)

            # Style header items
            for c, col_name in enumerate(headers):
                hdr_item = QTableWidgetItem(col_name)
                key = col_name.strip().lower()
                if key in _ACCENT_COLS:
                    hdr_item.setBackground(_ACCENT_BG)
                    hdr_item.setForeground(_HDR_FG)
                elif key in _YELLOW_COLS:
                    hdr_item.setBackground(_YELLOW_BG)
                    hdr_item.setForeground(_DARK_FG)
                else:
                    hdr_item.setBackground(_HDR_BG)
                    hdr_item.setForeground(_HDR_FG)
                f = hdr_item.font()
                f.setBold(True)
                hdr_item.setFont(f)
                tbl.setHorizontalHeaderItem(c, hdr_item)

            for c, val in enumerate(values):
                tbl.setItem(0, c, QTableWidgetItem(val))

            tbl.resizeColumnsToContents()
            tbl.setRowHeight(0, 34)
            # Use frameWidth so borders are included in the height
            frame = tbl.frameWidth() * 2
            tbl.setFixedHeight(tbl.horizontalHeader().sizeHint().height() + 34 + frame + 2)
            v.addWidget(tbl)

        v.addSpacing(8)
        btn_close = QPushButton("Close")
        btn_close.setObjectName("packshotUpdateRowsBtn")
        btn_close.clicked.connect(dlg.close)
        h = QHBoxLayout()
        h.setContentsMargins(0, 0, 10, 0)
        h.addStretch(1)
        h.addWidget(btn_close)
        v.addLayout(h)
        dlg.adjustSize()
        self._tracker_info_dialogs.append(dlg)
        dlg.destroyed.connect(lambda: self._tracker_info_dialogs.remove(dlg) if dlg in self._tracker_info_dialogs else None)
        dlg.show()

    # ── undo stack ─────────────────────────────────────────────────────────
    def _push_undo_snapshot(self, force: bool = False) -> None:
        if self._is_restoring_undo or (self._is_batch_edit and not force):
            return
        snap = self._get_table_snapshot()
        if self._undo_stack and self._undo_stack[-1] == snap:
            return
        self._undo_stack.append(snap)
        if len(self._undo_stack) > 50:
            self._undo_stack = self._undo_stack[-50:]

    def _on_table_batch_edit_begin(self) -> None:
        self._is_batch_edit = True

    def _on_table_batch_edit_end(self) -> None:
        self._is_batch_edit = False
        self._push_undo_snapshot()

    def _on_item_changed(self, item: "QTableWidgetItem") -> None:
        self.table.blockSignals(True)
        try:
            r, c = item.row(), item.column()
            self._apply_rule_formatting(item, c, item.text())
            self._apply_user_formatting(item, r, c)
        finally:
            self.table.blockSignals(False)
        self._push_undo_snapshot()

    # ── toolbar handlers ───────────────────────────────────────────────────
    def _on_delete_clicked(self) -> None:
        """Delete selected rows; if a column header was clicked, delete that column."""
        sel_model = self.table.selectionModel()
        # full columns selected?
        selected_cols = sorted(
            {idx.column() for idx in self.table.selectedIndexes()
             if sel_model.isColumnSelected(idx.column(), self.table.rootIndex())},
            reverse=True,
        )
        if selected_cols:
            self._push_undo_snapshot(force=True)
            for c in selected_cols:
                self.table.removeColumn(c)
                self._headers.pop(c) if c < len(self._headers) else None
            # rebuild filter keys
            self._column_filters = {i: None for i in range(self.table.columnCount())}
            self._refresh_header_labels()
            self._push_undo_snapshot(force=True)
            return
        # rows
        selected_rows = sorted(
            {idx.row() for idx in self.table.selectedIndexes()},
            reverse=True,
        )
        if not selected_rows:
            return
        self._push_undo_snapshot(force=True)
        for r in selected_rows:
            self.table.removeRow(r)
        self._push_undo_snapshot(force=True)
        self._update_row_count()

    def _on_reset_clicked(self) -> None:
        """Restore the table to its original loaded state."""
        reply = QMessageBox.question(
            self, "Reset Table",
            "Reset the table to its original loaded state? All unsaved changes will be lost.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._headers  = list(self._original_headers)
        self._user_fmt = {}
        col_count = len(self._headers)
        self.table.blockSignals(True)
        self.table.setColumnCount(col_count)
        self._column_filters = {i: None for i in range(col_count)}
        self._refresh_header_labels()
        self._populate_table(self._original_snapshot)
        self.table.blockSignals(False)
        self._undo_stack.clear()
        self._push_undo_snapshot(force=True)
        self._apply_filters()
        self._update_row_count()

    def _on_undo_clicked(self) -> None:
        if len(self._undo_stack) < 2:
            return
        self._undo_stack.pop()
        snapshot = self._undo_stack[-1]
        self._is_restoring_undo = True
        self.table.blockSignals(True)
        col_count = self.table.columnCount()
        self.table.setRowCount(len(snapshot))
        for r, row_data in enumerate(snapshot):
            for c in range(col_count):
                val = row_data[c] if c < len(row_data) else ""
                item = self.table.item(r, c)
                if item is None:
                    item = QTableWidgetItem(val)
                    self.table.setItem(r, c, item)
                else:
                    item.setText(val)
                self._apply_rule_formatting(item, c, val)
                self._apply_user_formatting(item, r, c)
        self.table.blockSignals(False)
        self._is_restoring_undo = False
        self._update_row_count()

    def _on_add_row(self) -> None:
        """Insert a blank row AFTER the currently selected row (or at end if none selected)."""
        self._push_undo_snapshot(force=True)
        current_row = self.table.currentRow()
        insert_at = (current_row + 1) if current_row >= 0 else self.table.rowCount()
        self.table.insertRow(insert_at)
        for c in range(self.table.columnCount()):
            self.table.setItem(insert_at, c, QTableWidgetItem(""))
        self._push_undo_snapshot(force=True)
        self._update_row_count()

    def _on_add_col(self) -> None:
        """Insert a blank column AFTER the currently selected column (or at end if none selected)."""
        col_name, ok = QInputDialog.getText(self, "Add Column", "Column name:")
        if not ok or not col_name.strip():
            return
        col_name = col_name.strip()
        self._push_undo_snapshot(force=True)
        current_col = self.table.currentColumn()
        insert_at = (current_col + 1) if current_col >= 0 else self.table.columnCount()
        self.table.blockSignals(True)
        self.table.insertColumn(insert_at)
        self._headers.insert(insert_at, col_name)
        # Rebuild filter dict, shifting indices for columns at/after insert_at
        old_filters = self._column_filters.copy()
        self._column_filters = {}
        for old_c, fval in old_filters.items():
            new_c = old_c if old_c < insert_at else old_c + 1
            self._column_filters[new_c] = fval
        self._column_filters[insert_at] = None
        self._refresh_header_labels()
        for r in range(self.table.rowCount()):
            self.table.setItem(r, insert_at, QTableWidgetItem(""))
        self.table.blockSignals(False)
        self._push_undo_snapshot(force=True)

    def _on_reset_filter_clicked(self) -> None:
        for k in self._column_filters:
            self._column_filters[k] = None
        self._apply_filters()

    # ── formatting handlers ────────────────────────────────────────────────
    def _on_format_bold(self) -> None:
        """Toggle bold on all selected cells."""
        selected = self.table.selectedItems()
        if not selected:
            return
        all_bold = all(item.font().bold() for item in selected)
        target = not all_bold
        self.table.blockSignals(True)
        for item in selected:
            r, c = item.row(), item.column()
            fmt = self._user_fmt.setdefault((r, c), {})
            fmt["bold"] = target
            font = item.font()
            font.setBold(target)
            item.setFont(font)
        self.table.blockSignals(False)

    def _on_format_cell_color(self) -> None:
        """Pick a background color and apply to selected cells."""
        color = QColorDialog.getColor(parent=self)
        if not color.isValid():
            return
        hex_color = color.name()
        self.table.blockSignals(True)
        for item in self.table.selectedItems():
            r, c = item.row(), item.column()
            self._user_fmt.setdefault((r, c), {})["bg"] = hex_color
            item.setBackground(QColor(hex_color))
        self.table.blockSignals(False)

    def _on_format_text_color(self) -> None:
        """Pick a text color and apply to selected cells."""
        color = QColorDialog.getColor(parent=self)
        if not color.isValid():
            return
        hex_color = color.name()
        self.table.blockSignals(True)
        for item in self.table.selectedItems():
            r, c = item.row(), item.column()
            self._user_fmt.setdefault((r, c), {})["fg"] = hex_color
            item.setForeground(QColor(hex_color))
        self.table.blockSignals(False)

    def _on_format_header_bg(self) -> None:
        """Pick a background color for the currently selected column's header."""
        col = self.table.currentColumn()
        if col < 0:
            QMessageBox.information(self, "No Column Selected",
                                    "Click on a column header to select it first.")
            return
        color = QColorDialog.getColor(parent=self)
        if not color.isValid():
            return
        self._header_fmt.setdefault(col, {})["bg"] = color.name()
        self._header_view.update_header_colors(self._header_fmt)

    def _on_format_header_fg(self) -> None:
        """Pick a text color for the currently selected column's header."""
        col = self.table.currentColumn()
        if col < 0:
            QMessageBox.information(self, "No Column Selected",
                                    "Click on a column header to select it first.")
            return
        color = QColorDialog.getColor(parent=self)
        if not color.isValid():
            return
        self._header_fmt.setdefault(col, {})["fg"] = color.name()
        self._header_view.update_header_colors(self._header_fmt)

    def _on_import_bma_clicked(self) -> None:
        """Open another BMA file and append rows whose Head Bom Mat doesn't yet exist."""
        from pathlib import Path
        start = self._sdc_folder or ""
        file, _ = QFileDialog.getOpenFileName(
            self, "Select BMA/SDC file to import", start,
            "Excel Files (*.xlsx *.xls *.xlsm)",
        )
        if not file or not Path(file).exists():
            return

        new_headers, new_rows = self._load_sheet(file)
        if not new_headers:
            QMessageBox.warning(self, "Import Error",
                                f"Could not read sheet '{self._SHEET_NAME}' from:\n{file}")
            return

        if not new_rows:
            QMessageBox.information(self, "Import", "The selected file contains no data rows.")
            return

        # Build column mapping: new-file col index → current table col index (case-insensitive)
        current_lower = [h.strip().lower() for h in self._headers]
        header_map: dict[int, int] = {}
        for new_c, new_h in enumerate(new_headers):
            key = new_h.strip().lower()
            if key in current_lower:
                header_map[new_c] = current_lower.index(key)

        if not header_map:
            QMessageBox.warning(self, "Import Error",
                                "No matching column headers found between the files.")
            return

        # Collect existing Head Bom Mat values to skip duplicates
        existing_hbm_col = next(
            (i for i, h in enumerate(self._headers) if h.strip().lower() == self._HBM_COL.lower()),
            None,
        )
        existing_hbm_vals: set[str] = set()
        if existing_hbm_col is not None:
            for r in range(self.table.rowCount()):
                item = self.table.item(r, existing_hbm_col)
                val = item.text().strip() if item else ""
                if val:
                    existing_hbm_vals.add(val.lower())

        new_hbm_col = next(
            (i for i, h in enumerate(new_headers) if h.strip().lower() == self._HBM_COL.lower()),
            None,
        )

        self._push_undo_snapshot(force=True)
        self.table.blockSignals(True)

        added = 0
        skipped = 0
        col_count = self.table.columnCount()
        for row_data in new_rows:
            # Skip rows whose Head Bom Mat already exists in the table
            if new_hbm_col is not None:
                hbm_val = str(row_data[new_hbm_col]).strip() if new_hbm_col < len(row_data) else ""
                if hbm_val.lower() in existing_hbm_vals:
                    skipped += 1
                    continue
                if hbm_val:
                    existing_hbm_vals.add(hbm_val.lower())

            insert_row = self.table.rowCount()
            self.table.insertRow(insert_row)
            for c in range(col_count):
                self.table.setItem(insert_row, c, QTableWidgetItem(""))
            for new_c, cur_c in header_map.items():
                val = str(row_data[new_c]).strip() if new_c < len(row_data) else ""
                item = QTableWidgetItem(val)
                self._apply_rule_formatting(item, cur_c, val)
                self.table.setItem(insert_row, cur_c, item)
            added += 1

        self.table.blockSignals(False)
        self._push_undo_snapshot(force=True)
        self._update_row_count()
        msg = f"{added} row(s) added from {Path(file).name}."
        if skipped:
            msg += f"\n{skipped} row(s) skipped (Head Bom Mat already exists)."
        QMessageBox.information(self, "Import Complete", msg)

    def _on_save_clicked(self) -> None:
        """Save table as review_YYYY_MM_DD_HH_MM.xlsx with full formatting in the Project Review folder."""
        from pathlib import Path
        from datetime import datetime
        try:
            import openpyxl
            from openpyxl.styles import (PatternFill, Font, Alignment,
                                         Border, Side, GradientFill)
            from openpyxl.utils import get_column_letter
        except ImportError:
            QMessageBox.warning(self, "Error", "openpyxl is required to save.")
            return

        # ── output path ────────────────────────────────────────────────────
        if self._project_review_folder and Path(self._project_review_folder).is_dir():
            out_dir = Path(self._project_review_folder)
        elif self._source_file:
            out_dir = Path(self._source_file).parent
        else:
            out_dir = Path.cwd()
        ts = datetime.now().strftime("%Y_%m_%d_%H_%M")
        out_path = out_dir / f"review_{ts}.xlsx"

        # ── helpers ────────────────────────────────────────────────────────
        def _hex_fill(hex_color: str | None) -> PatternFill | None:
            if not hex_color:
                return None
            return PatternFill("solid", fgColor=hex_color.lstrip("#"))

        def _hex_font(hex_color: str | None = None, bold: bool = False,
                      name: str = "Calibri", size: int = 11) -> Font:
            kw: dict = {"name": name, "size": size, "bold": bold}
            if hex_color:
                kw["color"] = hex_color.lstrip("#")
            return Font(**kw)

        thin = Side(style="thin", color="D0D0D0")
        cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Accent header cols (0-based indices)
        accent_col_indices: set[int] = self._compute_accent_cols()
        hbm_col = next(
            (i for i, h in enumerate(self._headers)
             if h.strip().lower() == self._HBM_COL.lower()),
            None,
        )

        # ── build workbook ─────────────────────────────────────────────────
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = self._SHEET_NAME

        col_count = self.table.columnCount()
        row_count = self.table.rowCount()

        # ── header row (row 1) ─────────────────────────────────────────────
        for c_idx, col_name in enumerate(self._headers):
            xl_col = c_idx + 1
            cell = ws.cell(row=1, column=xl_col, value=col_name)
            # User-set header color takes priority; fall back to accent/normal defaults
            h_fmt = self._header_fmt.get(c_idx, {})
            h_bg = h_fmt.get("bg") or ("#D02752" if c_idx in accent_col_indices else "#111F35")
            h_fg = h_fmt.get("fg") or "#FFFFFF"
            cell.fill      = _hex_fill(h_bg)
            cell.font      = _hex_font(h_fg, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
            cell.border    = cell_border

        ws.row_dimensions[1].height = 22

        # ── data rows ──────────────────────────────────────────────────────
        xl_row = 2
        for r in range(row_count):
            if self.table.isRowHidden(r):
                continue
            for c in range(col_count):
                tbl_item = self.table.item(r, c)
                raw_text = tbl_item.text() if tbl_item else ""
                # Strip the delegate arrow hint (not visible in xlsx anyway)
                xl_cell = ws.cell(row=xl_row, column=c + 1, value=raw_text)

                # Start with defaults
                bg_hex: str | None = None
                fg_hex: str | None = None
                bold = False

                # Rule-based formatting (same logic as _apply_rule_formatting)
                col_name = self._headers[c] if c < len(self._headers) else ""
                col_lower = col_name.strip().lower()
                val_lower  = raw_text.strip().lower()

                if col_lower == "hit type":
                    rule = self._HIT_TYPE_RULES.get(val_lower)
                    if rule:
                        bg_hex = rule.get("bg")
                        fg_hex = rule.get("fg")
                        bold   = rule.get("bold", False)

                elif col_lower == "master idh build":
                    rule = self._IDH_BUILD_RULES.get(val_lower)
                    if rule:
                        bg_hex = rule.get("bg")
                        fg_hex = rule.get("fg")
                        bold   = rule.get("bold", False)
                    elif val_lower and "na" in val_lower:
                        bg_hex = self._IDH_BUILD_NA_BG

                # User formatting overrides
                user_fmt = self._user_fmt.get((r, c), {})
                if "bg" in user_fmt and user_fmt["bg"] is not None:
                    bg_hex = user_fmt["bg"]  # already a "#RRGGBB" hex string
                if "fg" in user_fmt and user_fmt["fg"] is not None:
                    fg_hex = user_fmt["fg"]  # already a "#RRGGBB" hex string
                if "bold" in user_fmt:
                    bold = user_fmt["bold"]

                if bg_hex:
                    xl_cell.fill = _hex_fill(bg_hex)
                if fg_hex or bold:
                    xl_cell.font = _hex_font(fg_hex, bold=bold)
                xl_cell.alignment = Alignment(vertical="center", wrap_text=False)
                xl_cell.border    = cell_border

            ws.row_dimensions[xl_row].height = 18
            xl_row += 1

        # ── column widths (based on content) ───────────────────────────────
        for c_idx in range(col_count):
            col_letter = get_column_letter(c_idx + 1)
            # Measure header + all data cells
            max_len = len(self._headers[c_idx]) if c_idx < len(self._headers) else 8
            for r in range(row_count):
                if self.table.isRowHidden(r):
                    continue
                item = self.table.item(r, c_idx)
                if item:
                    max_len = max(max_len, len(item.text()))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 50)

        # ── freeze top row + autofilter ────────────────────────────────────
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        try:
            wb.save(str(out_path))
            QMessageBox.information(self, "Saved", f"Saved to:\n{out_path}")
        except Exception as exc:
            QMessageBox.warning(self, "Save Error", str(exc))

    # ── filter helpers ─────────────────────────────────────────────────────
    def _on_header_clicked(self, column: int) -> None:
        self.table.selectColumn(column)

    def _on_filter_header_clicked(self, column: int) -> None:
        values: list[str] = []
        seen: set[str] = set()
        has_blank = False
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            item = self.table.item(row, column)
            text = item.text().strip() if item else ""
            if text == "":
                has_blank = True
                continue
            if text in seen:
                continue
            seen.add(text)
            values.append(text)

        value_pairs = [(v, v) for v in values]
        if has_blank:
            value_pairs.append((self._BLANK_FILTER, "(Blanks)"))

        popup = self._FilterPopup(value_pairs, self._column_filters.get(column), self)
        popup.move(QCursor.pos())
        if popup.exec() != QDialog.DialogCode.Accepted:
            return

        selected = popup.get_selected_values()
        self._column_filters[column] = None if len(selected) == len(value_pairs) else selected
        self._apply_filters()

    def _apply_filters(self) -> None:
        for row in range(self.table.rowCount()):
            visible = True
            for col, fval in self._column_filters.items():
                if fval is None:
                    continue
                item = self.table.item(row, col)
                text = item.text().strip() if item else ""
                key = self._BLANK_FILTER if text == "" else text
                if key not in fval:
                    visible = False
                    break
            self.table.setRowHidden(row, not visible)
        self._refresh_header_labels()
        self._update_row_count()

    # ── stylesheet ─────────────────────────────────────────────────────────
    def _apply_stylesheet(self) -> None:
        self.setStyleSheet(
            """
            QDialog { background-color: #F4F4F4; }

            QTableWidget#packshotClipboardTable {
                background-color: #FFFFFF; color: #111111;
                border: 1px solid #A9A9A9; gridline-color: #B8B8B8;
                selection-background-color: #DCE6F5; selection-color: #111111;
                font-family: "Segoe UI"; font-size: 12px;
            }
            QTableWidget#packshotClipboardTable QHeaderView::section {
                background-color: #111F35; color: #FFFFFF;
                border: 1px solid #7D8694; padding: 6px 8px; font-weight: 700;
            }
            QTableWidget#packshotClipboardTable QTableCornerButton::section {
                background-color: #111F35; border: 1px solid #7D8694;
            }

            QLabel#packshotTableTitle {
                color: #111F35; font-family: "Segoe UI";
                font-size: 18px; font-weight: 800;
            }
            QLabel#tscCountLabel {
                color: #111F35; font-family: "Segoe UI";
                font-size: 13px; font-weight: 600;
            }
            QLabel#packshotRowCountLabel {
                color: #111F35; font-family: "Segoe UI";
                font-size: 13px; font-weight: 600;
            }
            QLineEdit#packshotRowCountInput {
                background-color: #FFFFFF; color: #111111;
                border: 1px solid #A9A9A9; border-radius: 8px;
                min-height: 30px; padding: 0 8px;
                font-family: "Segoe UI"; font-size: 12px;
            }

            QPushButton#packshotUpdateRowsBtn {
                background-color: #9EA3AB; color: #000000;
                border: 1px solid #8B9098; border-radius: 8px;
                min-height: 30px; padding: 0 12px;
                font-family: "Segoe UI"; font-size: 12px; font-weight: 600;
            }
            QPushButton#packshotUpdateRowsBtn:pressed {
                background-color: #111F35; color: #FFFFFF; border: 1px solid #111F35;
            }

            QPushButton#tscPrimaryBtn {
                background-color: #111F35; color: #FFFFFF;
                border: 1px solid #111F35; border-radius: 8px;
                min-height: 30px; padding: 0 12px;
                font-family: "Segoe UI"; font-size: 12px; font-weight: 600;
            }
            QPushButton#tscPrimaryBtn:pressed {
                background-color: #0e1726; color: #FFFFFF; border: 1px solid #0e1726;
            }
            """
        )


class _MapperOption1TableDialog(QDialog):
    def __init__(self, parent: QWidget | None = None, start_dir: str = "", export_dir: str = "") -> None:
        super().__init__(parent)
        self._browse_start_dir = start_dir
        self._export_dir = export_dir
        self.setWindowTitle("Open window")
        self.resize(1160, 620)
        self.setMinimumSize(1080, 520)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        title = QLabel("SAP Information")
        title.setObjectName("packshotTableTitle")
        layout.addWidget(title)

        config_row = QHBoxLayout()
        config_row.setSpacing(8)
        row_count_label = QLabel("Row count")
        row_count_label.setObjectName("packshotRowCountLabel")
        config_row.addWidget(row_count_label, 0)

        self.row_count_input = QLineEdit("5")
        self.row_count_input.setObjectName("packshotRowCountInput")
        self.row_count_input.setFixedWidth(90)
        config_row.addWidget(self.row_count_input, 0)

        self.btn_update_rows = QPushButton("Update")
        self.btn_update_rows.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_update_rows, 0)

        self.btn_reset_table = QPushButton("Reset")
        self.btn_reset_table.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_reset_table, 0)

        self.btn_delete_row = QPushButton("Delete")
        self.btn_delete_row.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_delete_row, 0)

        self.btn_undo_table = QPushButton("Undo")
        self.btn_undo_table.setObjectName("packshotUpdateRowsBtn")
        config_row.addWidget(self.btn_undo_table, 0)

        self.radio_cleanup_1 = QRadioButton("Cleanup 1")
        self.radio_cleanup_1.setChecked(True)
        self.radio_cleanup_1.setObjectName("collectorCleanupRadio")
        self.radio_cleanup_1.setToolTip(
            "remove non-SMU and unneeded packaging\n"
            "(retains accl and flex SMU rows):\n"
            "sal, pal, film, ship, wgl, sheet, shee,\n"
            "pl, t-secur, saco, rbosac, acco_pe, tear, bulk"
        )
        config_row.addWidget(self.radio_cleanup_1, 0)

        self.radio_cleanup_2 = QRadioButton("Cleanup 2")
        self.radio_cleanup_2.setObjectName("collectorCleanupRadio")
        self.radio_cleanup_2.setToolTip(
            "remove non-SMU and unneeded packaging\n"
            "(including accl and flex):\n"
            "sal, flex, pall, accl, film, ship, wgl,\n"
            "sheet, shee, pl, t-secur, saco, acco_pe,\n"
            "tear, bulk"
        )
        config_row.addWidget(self.radio_cleanup_2, 0)

        self.radio_cleanup_3 = QRadioButton("Cleanup 3")
        self.radio_cleanup_3.setObjectName("collectorCleanupRadio")
        self.radio_cleanup_3.setToolTip(
            "remove non-SMU, un-identifiable basic name\n"
            "and unneeded packaging:\n"
            "sal, flex, pall, accl, film, ship, wgl,\n"
            "sheet, shee, pl, t-secur, saco, bag,\n"
            "rbosac, leaflet, acco, paco, tear, bulk"
        )
        config_row.addWidget(self.radio_cleanup_3, 0)

        self.radio_cleanup_4 = QRadioButton("Cleanup 4")
        self.radio_cleanup_4.setObjectName("collectorCleanupRadio")
        self.radio_cleanup_4.setToolTip("no cleanup, all info retained")
        config_row.addWidget(self.radio_cleanup_4, 0)

        self.cleanup_radio_group = QButtonGroup(self)
        self.cleanup_radio_group.setExclusive(True)
        for rb in (self.radio_cleanup_1, self.radio_cleanup_2, self.radio_cleanup_3, self.radio_cleanup_4):
            self.cleanup_radio_group.addButton(rb)

        self.btn_import_tracker = QPushButton("Import SAP Data")
        self.btn_import_tracker.setObjectName("packshotUpdateRowsBtn")

        self.import_progress = QProgressBar(self)
        self.import_progress.setObjectName("sapImportProgress")
        self.import_progress.setFixedWidth(140)
        self.import_progress.setRange(0, 100)
        self.import_progress.setValue(0)
        self.import_progress.setVisible(False)
        config_row.addWidget(self.import_progress, 0)

        config_row.addWidget(self.btn_import_tracker, 0)

        self.btn_reformat_table = QPushButton("Reformat")
        self.btn_reformat_table.setObjectName("mapperReformatBtn")
        config_row.addWidget(self.btn_reformat_table, 0)

        config_row.addStretch(1)
        layout.addLayout(config_row)

        self.table = _ClipboardTableWidget(5, 6, self)
        self.table.setObjectName("packshotClipboardTable")
        self._set_headers()
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.horizontalHeader().setDefaultSectionSize(154)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        for row in range(5):
            for col in range(6):
                self.table.setItem(row, col, QTableWidgetItem(""))

        self.sap_table_reformatter = SapTableReformatter()
        self._last_imported_stem: str = ""
        self._undo_stack: list[list[list[str]]] = []
        self._is_restoring_undo = False
        self._is_batch_edit = False

        layout.addWidget(self.table)
        self.btn_update_rows.clicked.connect(self._on_update_rows_clicked)
        self.btn_reset_table.clicked.connect(self._on_reset_table_clicked)
        self.btn_delete_row.clicked.connect(self._on_delete_rows_clicked)
        self.btn_undo_table.clicked.connect(self._on_undo_table_clicked)
        self.btn_import_tracker.clicked.connect(self._on_import_tracker_clicked)
        self.btn_reformat_table.clicked.connect(self._on_reformat_table_clicked)
        self.table.itemChanged.connect(self._on_table_item_changed)
        self._push_undo_snapshot(force=True)
        self._update_row_count_label()

        self.setStyleSheet(
            """
            QDialog {
                background-color: #F4F4F4;
            }

            QTableWidget#packshotClipboardTable {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                gridline-color: #B8B8B8;
                selection-background-color: #DCE6F5;
                selection-color: #111111;
                font-family: "Segoe UI";
                font-size: 12px;
            }

            QTableWidget#packshotClipboardTable QHeaderView::section {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #7D8694;
                padding: 6px 8px;
                font-weight: 700;
            }

            QTableWidget#packshotClipboardTable QTableCornerButton::section {
                background-color: #111F35;
                border: 1px solid #7D8694;
            }

            QLabel#packshotTableTitle {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 18px;
                font-weight: 800;
            }

            QLabel#packshotRowCountLabel {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 600;
            }

            QLineEdit#packshotRowCountInput {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 8px;
                font-family: "Segoe UI";
                font-size: 12px;
            }

            QProgressBar#sapImportProgress {
                background-color: #FFFFFF;
                border: 1px solid #A9A9A9;
                border-radius: 7px;
                text-align: center;
                color: #111111;
                font-family: "Segoe UI";
                font-size: 11px;
                min-height: 24px;
            }

            QProgressBar#sapImportProgress::chunk {
                background-color: #8A244B;
                border-radius: 6px;
            }

            QPushButton#packshotUpdateRowsBtn {
                background-color: #9EA3AB;
                color: #000000;
                border: 1px solid #8B9098;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }

            QPushButton#packshotUpdateRowsBtn:pressed {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #111F35;
            }

            QPushButton#mapperReformatBtn {
                background-color: #8A244B;
                color: #FFFFFF;
                border: 1px solid #8A244B;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }

            QPushButton#mapperReformatBtn:pressed {
                background-color: #D02752;
                color: #FFFFFF;
                border: 1px solid #D02752;
            }

            QRadioButton {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }

            QRadioButton::indicator {
                width: 14px;
                height: 14px;
            }

            QRadioButton::indicator:unchecked {
                border: 1px solid #8B9098;
                border-radius: 7px;
                background-color: #FFFFFF;
            }

            QRadioButton::indicator:checked {
                border: 1px solid #8A244B;
                border-radius: 7px;
                background-color: #8A244B;
            }
            """
        )

    def _set_headers(self) -> None:
        self.table.setHorizontalHeaderLabels(
            [
                "Head Bom Mat",
                "BOM COMPONENT",
                "Sort String",
                "Component Desc",
                "Basic Number",
                "Basic Name",
            ]
        )

    def _update_row_count_label(self) -> None:
        if hasattr(self, "label_table_count"):
            self.label_table_count.setText(f"Count: <b>{self.table.rowCount()}</b>")

    def _set_import_progress(self, value: int, visible: bool = True) -> None:
        self.import_progress.setVisible(visible)
        if visible:
            self.import_progress.setValue(max(0, min(100, value)))
        QApplication.processEvents()

    @staticmethod
    def _normalize_header_text(value: object) -> str:
        if value is None:
            return ""
        text = str(value).strip().lower()
        text = text.replace("\n", " ").replace("\r", " ")
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return " ".join(text.split())

    @staticmethod
    def _cell_to_text(value: object) -> str:
        if pd.isna(value):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def _read_excel_raw(self, file_path: str) -> pd.DataFrame:
        lower_path = file_path.lower()

        if lower_path.endswith(".xls"):
            # Some SAP exports are text-based files with .xls extension.
            # Try as real Excel first, then fallback to auto text->xlsx conversion.
            try:
                excel_data = pd.ExcelFile(file_path, engine="openpyxl")
                if excel_data.sheet_names:
                    return pd.read_excel(
                        excel_data,
                        sheet_name=excel_data.sheet_names[0],
                        header=None,
                        dtype=object,
                    )
            except Exception:
                pass

            try:
                excel_data = pd.ExcelFile(file_path, engine="xlrd")
            except ImportError as exc:
                excel_data = None
            except Exception as exc:
                excel_data = None

            if excel_data is not None:
                if not excel_data.sheet_names:
                    raise ValueError("No worksheet found in selected file.")
                return pd.read_excel(
                    excel_data,
                    sheet_name=excel_data.sheet_names[0],
                    header=None,
                    dtype=object,
                )

            text_df = self._read_text_style_xls(file_path)
            if text_df is None:
                raise ValueError(
                    "Unable to read .xls file. File may be corrupted or in unsupported SAP export format."
                )

            tmp_xlsx_path = self._convert_dataframe_to_temp_xlsx(text_df)
            try:
                excel_data = pd.ExcelFile(tmp_xlsx_path, engine="openpyxl")
                if not excel_data.sheet_names:
                    raise ValueError("No worksheet found in converted SAP file.")
                return pd.read_excel(
                    excel_data,
                    sheet_name=excel_data.sheet_names[0],
                    header=None,
                    dtype=object,
                )
            finally:
                try:
                    Path(tmp_xlsx_path).unlink(missing_ok=True)
                except Exception:
                    pass
        else:
            last_error: Exception | None = None
            excel_data = None
            for engine in ("openpyxl", None):
                try:
                    excel_data = pd.ExcelFile(file_path, engine=engine)
                    break
                except Exception as exc:
                    last_error = exc

            if excel_data is None:
                if last_error is not None:
                    raise last_error
                raise ValueError("Unable to read selected file.")

        if not excel_data.sheet_names:
            raise ValueError("No worksheet found in selected file.")

        return pd.read_excel(
            excel_data,
            sheet_name=excel_data.sheet_names[0],
            header=None,
            dtype=object,
        )

    def _read_text_style_xls(self, file_path: str) -> pd.DataFrame | None:
        parsed = self._try_read_sap_text_file(file_path)
        if parsed is not None:
            return parsed

        # Fallback: internally perform the manual trick (rename .xls -> .csv) and parse again.
        tmp_csv = tempfile.NamedTemporaryFile(suffix=".csv", delete=False)
        tmp_csv_path = tmp_csv.name
        tmp_csv.close()

        try:
            Path(tmp_csv_path).write_bytes(Path(file_path).read_bytes())
            return self._try_read_sap_text_file(tmp_csv_path)
        except Exception:
            return None
        finally:
            try:
                Path(tmp_csv_path).unlink(missing_ok=True)
            except Exception:
                pass

    def _try_read_sap_text_file(self, file_path: str) -> pd.DataFrame | None:
        parse_configs = [
            {"encoding": "utf-16", "sep": "\t", "skiprows": 3},
            {"encoding": "utf-16", "sep": r"\s{2,}", "skiprows": 3},
            {"encoding": "utf-16le", "sep": "\t", "skiprows": 3},
            {"encoding": "utf-16le", "sep": r"\s{2,}", "skiprows": 3},
            {"encoding": "utf-8-sig", "sep": "\t", "skiprows": 3},
            {"encoding": "cp1252", "sep": "\t", "skiprows": 3},
            {"encoding": "utf-16", "sep": "\t", "skiprows": 0},
            {"encoding": "utf-16", "sep": r"\s{2,}", "skiprows": 0},
        ]

        for config in parse_configs:
            try:
                df = pd.read_csv(
                    file_path,
                    header=None,
                    dtype=object,
                    engine="python",
                    on_bad_lines="skip",
                    quotechar='"',
                    quoting=csv.QUOTE_MINIMAL,
                    **config,
                )
            except Exception:
                continue

            if df.empty:
                continue

            df = df.dropna(how="all").reset_index(drop=True)
            if df.empty:
                continue

            # Must be reasonably tabular for SAP header detection to work.
            if df.shape[1] < 4:
                continue

            return df

        return None

    @staticmethod
    def _convert_dataframe_to_temp_xlsx(df: pd.DataFrame) -> str:
        tmp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_path = tmp_file.name
        tmp_file.close()

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "SAP Data"

        safe_df = df.fillna("")
        for row_values in safe_df.itertuples(index=False, name=None):
            worksheet.append(list(row_values))

        workbook.save(tmp_path)
        return tmp_path

    def _detect_header_row_and_columns(self, raw_df: pd.DataFrame) -> tuple[int, dict[str, int]]:
        target_columns: dict[str, list[str]] = {
            "Head Bom Mat": ["head bom mat", "head bom", "headbommat"],
            "BOM COMPONENT": ["bom component", "bom comp", "bomcomponent"],
            "Sort String": ["sort string", "sortstring"],
            "Component Desc": ["component desc", "component description", "component"],
            "Basic Number": ["basic number", "basic num", "basic no", "basic"],
            "Basic Name": ["basic name", "basicname"],
        }
        normalized_targets = {
            target: {self._normalize_header_text(alias) for alias in aliases}
            for target, aliases in target_columns.items()
        }

        max_scan_rows = min(120, raw_df.shape[0])
        for row_idx in range(max_scan_rows):
            row_values = raw_df.iloc[row_idx].tolist()
            row_normalized = [self._normalize_header_text(value) for value in row_values]

            found_columns: dict[str, int] = {}
            for target_name, aliases in normalized_targets.items():
                for col_idx, cell_value in enumerate(row_normalized):
                    if cell_value in aliases:
                        found_columns[target_name] = col_idx
                        break

            if len(found_columns) == len(target_columns):
                return row_idx, found_columns

        missing = ", ".join(target_columns.keys())
        raise ValueError(f"Unable to detect all required headers. Required headers: {missing}")

    def _extract_sap_rows(self, raw_df: pd.DataFrame, header_row_idx: int, col_map: dict[str, int]) -> list[list[str]]:
        ordered_headers = [
            "Head Bom Mat",
            "BOM COMPONENT",
            "Sort String",
            "Component Desc",
            "Basic Number",
            "Basic Name",
        ]
        out_rows: list[list[str]] = []

        for row_idx in range(header_row_idx + 1, raw_df.shape[0]):
            row_values: list[str] = []
            has_value = False
            for header in ordered_headers:
                col_idx = col_map[header]
                value = raw_df.iat[row_idx, col_idx] if col_idx < raw_df.shape[1] else ""
                text_value = self._cell_to_text(value)
                if text_value:
                    has_value = True
                row_values.append(text_value)

            if has_value:
                out_rows.append(row_values)

        return out_rows

    def _on_import_tracker_clicked(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select SAP File",
            self._browse_start_dir,
            "Excel Files (*.xls)",
        )
        if not file_path:
            return

        self._last_imported_stem = Path(file_path).stem

        self._set_import_progress(5, visible=True)
        try:
            raw_df = self._read_excel_raw(file_path)
            self._set_import_progress(35, visible=True)
            header_row_idx, col_map = self._detect_header_row_and_columns(raw_df)
            self._set_import_progress(60, visible=True)
            extracted_rows = self._extract_sap_rows(raw_df, header_row_idx, col_map)
            self._set_import_progress(80, visible=True)
        except Exception as exc:
            self._set_import_progress(0, visible=False)
            msg = QMessageBox(self)
            msg.setWindowTitle("Error")
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText(f"Failed to import SAP data. {exc}")
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.exec()
            return

        self.table.blockSignals(True)
        try:
            row_count = max(1, len(extracted_rows))
            self.table.setRowCount(row_count)
            for row_idx in range(row_count):
                values = extracted_rows[row_idx] if row_idx < len(extracted_rows) else [""] * self.table.columnCount()
                for col_idx in range(self.table.columnCount()):
                    text_value = values[col_idx] if col_idx < len(values) else ""
                    item = self.table.item(row_idx, col_idx)
                    if item is None:
                        item = QTableWidgetItem("")
                        self.table.setItem(row_idx, col_idx, item)
                    item.setText(text_value)
        finally:
            self.table.blockSignals(False)

        self.table.clearSelection()
        self.row_count_input.setText(str(self.table.rowCount()))
        self._update_row_count_label()
        self._push_undo_snapshot(force=True)
        self._set_import_progress(100, visible=True)
        self.import_progress.setVisible(False)

    def _on_update_rows_clicked(self) -> None:
        raw_value = self.row_count_input.text().strip()
        try:
            new_count = int(raw_value)
        except ValueError:
            self.row_count_input.setText(str(self.table.rowCount()))
            return

        new_count = max(1, min(new_count, 5000))
        self.table.setRowCount(new_count)
        for row in range(new_count):
            for col in range(self.table.columnCount()):
                if self.table.item(row, col) is None:
                    self.table.setItem(row, col, QTableWidgetItem(""))
        self.row_count_input.setText(str(new_count))
        self._update_row_count_label()
        self._push_undo_snapshot()

    def _on_reset_table_clicked(self) -> None:
        self.row_count_input.setText("5")
        self.table.clearContents()
        self.table.setRowCount(5)
        self.table.clearSelection()
        for row in range(5):
            for col in range(self.table.columnCount()):
                self.table.setItem(row, col, QTableWidgetItem(""))
        self._update_row_count_label()
        self._push_undo_snapshot()

    def _on_delete_rows_clicked(self) -> None:
        selected_rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        if not selected_rows:
            current_row = self.table.currentRow()
            if current_row >= 0:
                selected_rows = [current_row]
            else:
                return

        total_rows = self.table.rowCount()
        if total_rows <= 0:
            return

        if len(selected_rows) >= total_rows:
            self.table.setRowCount(1)
            for col in range(self.table.columnCount()):
                self.table.setItem(0, col, QTableWidgetItem(""))
        else:
            for row in selected_rows:
                if 0 <= row < self.table.rowCount():
                    self.table.removeRow(row)

        self.table.clearSelection()
        self.row_count_input.setText(str(self.table.rowCount()))
        self._update_row_count_label()
        self._push_undo_snapshot()

    def _capture_table_state(self) -> list[list[str]]:
        state: list[list[str]] = []
        for row in range(self.table.rowCount()):
            row_values: list[str] = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_values.append(item.text() if item is not None else "")
            state.append(row_values)
        return state

    def _restore_table_state(self, state: list[list[str]]) -> None:
        self._is_restoring_undo = True
        self.table.blockSignals(True)
        try:
            row_count = max(1, len(state))
            col_count = self.table.columnCount()
            self.table.setRowCount(row_count)

            for row in range(row_count):
                row_values = state[row] if row < len(state) else [""] * col_count
                for col in range(col_count):
                    value = row_values[col] if col < len(row_values) else ""
                    item = self.table.item(row, col)
                    if item is None:
                        item = QTableWidgetItem("")
                        self.table.setItem(row, col, item)
                    item.setText(value)
        finally:
            self.table.blockSignals(False)
            self._is_restoring_undo = False

        self.row_count_input.setText(str(self.table.rowCount()))
        self._update_row_count_label()

    def _push_undo_snapshot(self, force: bool = False) -> None:
        if self._is_restoring_undo or self._is_batch_edit:
            return

        snapshot = self._capture_table_state()
        if not force and self._undo_stack and self._undo_stack[-1] == snapshot:
            return

        self._undo_stack.append(snapshot)
        if len(self._undo_stack) > 200:
            self._undo_stack.pop(0)

    def _on_table_item_changed(self, _item: QTableWidgetItem) -> None:
        if self._is_batch_edit:
            return
        self._push_undo_snapshot()

    def _on_table_batch_edit_begin(self) -> None:
        if self._is_restoring_undo:
            return
        self._push_undo_snapshot(force=False)
        self._is_batch_edit = True

    def _on_table_batch_edit_end(self) -> None:
        if self._is_restoring_undo:
            return
        self._is_batch_edit = False
        self._push_undo_snapshot(force=False)

    def _on_undo_table_clicked(self) -> None:
        if self._is_restoring_undo:
            return

        current_snapshot = self._capture_table_state()
        if not self._undo_stack:
            self._undo_stack.append(current_snapshot)
            return

        if self._undo_stack[-1] != current_snapshot:
            self._undo_stack.append(current_snapshot)

        if len(self._undo_stack) <= 1:
            return

        self._undo_stack.pop()
        previous_snapshot = self._undo_stack[-1]
        self._restore_table_state(previous_snapshot)

    def _on_reformat_table_clicked(self) -> None:
        cleanup_mode = next(
            (i for i, rb in enumerate(
                [self.radio_cleanup_1, self.radio_cleanup_2, self.radio_cleanup_3, self.radio_cleanup_4],
                start=1,
            ) if rb.isChecked()),
            1,
        )

        try:
            reformatted_rows, basic_comb_count = self.sap_table_reformatter.reformat_from_table(self.table, cleanup_mode)
        except SapTableReformatError as exc:
            msg = QMessageBox(self)
            msg.setWindowTitle(f"Cleanup {cleanup_mode} Error")
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText(str(exc))
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.exec()
            return

        if not hasattr(self, "_reformatted_dialogs"):
            self._reformatted_dialogs: list = []

        dlg = _MapperReformattedTableDialog(
            reformatted_rows,
            basic_comb_count,
            cleanup_mode,
            self,
            source_stem=self._last_imported_stem,
            export_dir=self._export_dir,
        )
        self._reformatted_dialogs.append(dlg)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()


class _MapperReformattedTableDialog(QDialog):
    _BLANK_FILTER = "__BLANK__"

    class _FilterPopup(QDialog):
        def __init__(
            self,
            values: list[tuple[str, str]],
            selected_values: set[str] | None,
            parent: QWidget | None = None,
        ) -> None:
            super().__init__(parent)
            self.setWindowFlags(Qt.WindowType.Popup)
            self.setMinimumSize(260, 330)
            self.resize(260, 330)

            self._is_syncing = False
            self._value_items: list[tuple[QListWidgetItem, str, str]] = []

            layout = QVBoxLayout(self)
            layout.setContentsMargins(4, 4, 4, 4)
            layout.setSpacing(4)

            self.search_input = QLineEdit(self)
            self.search_input.setPlaceholderText("Search")
            self.search_input.setClearButtonEnabled(True)
            self.search_input.setObjectName("packshotRowCountInput")
            layout.addWidget(self.search_input)

            self.list_widget = QListWidget(self)
            self.list_widget.setSelectionMode(QListWidget.SelectionMode.NoSelection)
            layout.addWidget(self.list_widget, 1)

            self.item_select_all = QListWidgetItem("(Select All)")
            self.item_select_all.setFlags(self.item_select_all.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            self.item_select_all.setCheckState(Qt.CheckState.Checked)
            self.item_select_all.setData(Qt.ItemDataRole.UserRole, "__SELECT_ALL__")
            self.list_widget.addItem(self.item_select_all)

            selected = selected_values if selected_values is not None else {value for value, _label in values}

            for value, label in values:
                item = QListWidgetItem(label)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if value in selected else Qt.CheckState.Unchecked)
                item.setData(Qt.ItemDataRole.UserRole, value)
                self.list_widget.addItem(item)
                self._value_items.append((item, value, label))

            button_row = QHBoxLayout()
            button_row.addStretch(1)
            self.btn_apply = QPushButton("Apply")
            self.btn_apply.setObjectName("packshotUpdateRowsBtn")
            self.btn_cancel = QPushButton("Cancel")
            self.btn_cancel.setObjectName("packshotUpdateRowsBtn")
            button_row.addWidget(self.btn_apply, 0)
            button_row.addWidget(self.btn_cancel, 0)
            layout.addLayout(button_row)

            self.search_input.textChanged.connect(self._apply_search)
            self.list_widget.itemChanged.connect(self._on_item_changed)
            self.btn_apply.clicked.connect(self.accept)
            self.btn_cancel.clicked.connect(self.reject)

            self.search_input.setFocus()
            self._sync_select_all_state()

        def _apply_search(self, text: str) -> None:
            needle = text.strip().lower()
            for item, _value, label in self._value_items:
                item.setHidden(needle not in label.lower())
            self._sync_select_all_state()

        def _on_item_changed(self, item: QListWidgetItem) -> None:
            if self._is_syncing:
                return

            role = item.data(Qt.ItemDataRole.UserRole)
            if role == "__SELECT_ALL__":
                self._is_syncing = True
                try:
                    target_state = item.checkState()
                    for value_item, _value, _label in self._value_items:
                        if value_item.isHidden():
                            continue
                        value_item.setCheckState(target_state)
                finally:
                    self._is_syncing = False
                return

            self._sync_select_all_state()

        def _sync_select_all_state(self) -> None:
            visible_items = [item for item, _value, _label in self._value_items if not item.isHidden()]
            if not visible_items:
                state = Qt.CheckState.Unchecked
            else:
                all_checked = all(i.checkState() == Qt.CheckState.Checked for i in visible_items)
                any_checked = any(i.checkState() == Qt.CheckState.Checked for i in visible_items)
                if all_checked:
                    state = Qt.CheckState.Checked
                elif any_checked:
                    state = Qt.CheckState.PartiallyChecked
                else:
                    state = Qt.CheckState.Unchecked

            self._is_syncing = True
            try:
                self.item_select_all.setCheckState(state)
            finally:
                self._is_syncing = False

        def get_selected_values(self) -> set[str]:
            selected: set[str] = set()
            for item, value, _label in self._value_items:
                if item.checkState() == Qt.CheckState.Checked:
                    selected.add(value)
            return selected

    def __init__(
        self,
        rows: list[list[str]],
        basic_comb_count: int,
        cleanup_mode: int,
        parent: QWidget | None = None,
        source_stem: str = "",
        export_dir: str = "",
    ) -> None:
        super().__init__(parent)
        self._cleanup_mode = cleanup_mode if cleanup_mode in (1, 2, 3, 4) else 1
        self._source_stem = source_stem
        self._export_dir = export_dir
        self.setWindowTitle("Reformatted SAP Data")
        self.resize(1120, 620)
        self.setMinimumSize(1020, 520)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        header_row = QHBoxLayout()
        header_row.setSpacing(8)

        title = QLabel("Reformatted SAP Data")
        title.setObjectName("mapperReformattedTitle")
        header_row.addWidget(title, 0)
        header_row.addSpacing(16)

        self.btn_reset_filter = QPushButton("Reset Filter")
        self.btn_reset_filter.setObjectName("packshotUpdateRowsBtn")
        header_row.addWidget(self.btn_reset_filter, 0)
        header_row.addSpacing(8)

        self.btn_export_table = QPushButton("Export")
        self.btn_export_table.setObjectName("packshotUpdateRowsBtn")
        header_row.addWidget(self.btn_export_table, 0)
        header_row.addSpacing(6)

        self.export_progress = QProgressBar(self)
        self.export_progress.setObjectName("mapperExportProgress")
        self.export_progress.setFixedWidth(140)
        self.export_progress.setRange(0, 100)
        self.export_progress.setValue(0)
        self.export_progress.setVisible(False)
        header_row.addWidget(self.export_progress, 0)
        header_row.addSpacing(8)

        self.btn_grouping_count = QPushButton("Grouping Count")
        self.btn_grouping_count.setObjectName("packshotUpdateRowsBtn")
        header_row.addWidget(self.btn_grouping_count, 0)

        header_row.addSpacing(14)

        self._selected_column_for_count = 0
        self.label_selection_count = QLabel("Count: <b>0</b>")
        self.label_selection_count.setObjectName("mapperReformattedOverview")
        self.label_selection_count.setTextFormat(Qt.TextFormat.RichText)
        header_row.addWidget(self.label_selection_count, 0)
        header_row.addSpacing(20)

        self.label_hsi_count = QLabel("HSI Count: <b>0</b>")
        self.label_hsi_count.setObjectName("mapperReformattedOverview")
        self.label_hsi_count.setTextFormat(Qt.TextFormat.RichText)
        self.label_hsi_count.setVisible(self._cleanup_mode != 4)
        header_row.addWidget(self.label_hsi_count, 0)
        header_row.addSpacing(20)

        self.label_basic_comb_overview = QLabel(
            f"Total Count of Basic Combinations: <b>{basic_comb_count}</b>"
        )
        self.label_basic_comb_overview.setObjectName("mapperReformattedOverview")
        self.label_basic_comb_overview.setTextFormat(Qt.TextFormat.RichText)
        header_row.addWidget(self.label_basic_comb_overview, 0)
        header_row.addSpacing(24)

        _cu_tag = f"cu{self._cleanup_mode}"
        self.label_cleanup_tag = QLabel(_cu_tag)
        self.label_cleanup_tag.setObjectName("mapperCleanupTag")
        header_row.addWidget(self.label_cleanup_tag, 0)
        header_row.addStretch(1)
        layout.addLayout(header_row)

        self.table = _ClipboardTableWidget(max(len(rows), 1), 7, self)
        self.table.setObjectName("mapperReformattedTable")
        self._header_labels = [
            "Head Bom Mat",
            "HSI",
            "BOM COMPONENT",
            "Component Desc",
            "Basic Number",
            "BC",
            "Basic Name",
        ]
        self._column_filters: dict[int, set[str] | None] = {i: None for i in range(len(self._header_labels))}
        self._refresh_header_labels()
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.horizontalHeader().setDefaultSectionSize(200)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 100)  # Head Bom Mat
        self.table.setColumnWidth(1, 100)  # HSI
        self.table.setColumnWidth(5, 100)  # BC
        self.table.horizontalHeader().setSectionsClickable(True)
        self.table.horizontalHeader().sectionClicked.connect(self._on_header_clicked)
        self.table.horizontalHeader().sectionDoubleClicked.connect(self._on_filter_header_clicked)

        for row_idx, values in enumerate(rows):
            for col_idx, value in enumerate(values):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(value))

        if not rows:
            for col_idx in range(7):
                self.table.setItem(0, col_idx, QTableWidgetItem(""))

        layout.addWidget(self.table)
        self.table.selectColumn(0)
        self._update_selection_count_label()
        self._update_hsi_count_label()

        self.btn_export_table.clicked.connect(self._on_export_table_clicked)
        self.btn_reset_filter.clicked.connect(self._on_reset_filter_clicked)
        self.btn_grouping_count.clicked.connect(self._on_grouping_count_clicked)

        self.setStyleSheet(
            """
            QDialog {
                background-color: #F4F4F4;
            }

            QLabel#mapperReformattedTitle {
                color: #8A244B;
                font-family: "Segoe UI";
                font-size: 18px;
                font-weight: 800;
            }

            QLabel#mapperReformattedOverview {
                color: #111F35;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 600;
            }

            QLabel#mapperCleanupTag {
                color: #D02752;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 700;
            }

            QTableWidget#mapperReformattedTable {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                gridline-color: #B8B8B8;
                selection-background-color: #DCE6F5;
                selection-color: #111111;
                font-family: "Segoe UI";
                font-size: 12px;
            }

            QTableWidget#mapperReformattedTable QHeaderView::section {
                background-color: #8A244B;
                color: #FFFFFF;
                border: 1px solid #7D8694;
                padding: 6px 8px;
                font-weight: 700;
            }

            QTableWidget#mapperReformattedTable QTableCornerButton::section {
                background-color: #8A244B;
                border: 1px solid #7D8694;
            }

            QPushButton#packshotUpdateRowsBtn {
                background-color: #9EA3AB;
                color: #000000;
                border: 1px solid #8B9098;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }

            QPushButton#packshotUpdateRowsBtn:pressed {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #111F35;
            }

            QLineEdit#packshotRowCountInput {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 8px;
                font-family: "Segoe UI";
                font-size: 12px;
            }

            QProgressBar#mapperExportProgress {
                background-color: #FFFFFF;
                border: 1px solid #A9A9A9;
                border-radius: 7px;
                text-align: center;
                color: #111111;
                font-family: "Segoe UI";
                font-size: 11px;
                min-height: 24px;
            }

            QProgressBar#mapperExportProgress::chunk {
                background-color: #8A244B;
                border-radius: 6px;
            }
            """
        )

    def _refresh_header_labels(self) -> None:
        labels = []
        for idx, base_label in enumerate(self._header_labels):
            has_active_filter = self._column_filters.get(idx) is not None
            labels.append(f"{base_label} {'▾*' if has_active_filter else '▾'}")
        self.table.setHorizontalHeaderLabels(labels)

    def _on_header_clicked(self, column: int) -> None:
        self._selected_column_for_count = column
        self.table.selectColumn(column)
        self._update_selection_count_label()

    def _on_filter_header_clicked(self, column: int) -> None:
        values: list[str] = []
        seen: set[str] = set()
        has_blank = False
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            item = self.table.item(row, column)
            cell_text = item.text().strip() if item is not None else ""
            if cell_text == "":
                has_blank = True
                continue
            if cell_text in seen:
                continue
            seen.add(cell_text)
            values.append(cell_text)

        value_pairs: list[tuple[str, str]] = [(value, value) for value in values]
        if has_blank:
            value_pairs.append((self._BLANK_FILTER, "(Blanks)"))

        popup = self._FilterPopup(value_pairs, self._column_filters.get(column), self)
        popup.move(QCursor.pos())
        if popup.exec() != QDialog.DialogCode.Accepted:
            return

        selected = popup.get_selected_values()
        if len(selected) == len(value_pairs):
            self._column_filters[column] = None
        else:
            self._column_filters[column] = selected
        self._apply_filters()

    def _update_selection_count_label(self) -> None:
        count = 0
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            count += 1
        self.label_selection_count.setText(f"Count: <b>{count}</b>")

    def _update_hsi_count_label(self) -> None:
        yes_count = 0
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            item = self.table.item(row, 1)
            hsi_value = item.text().strip().upper() if item is not None else ""
            if hsi_value == "YES":
                yes_count += 1
        self.label_hsi_count.setText(f"HSI Count: <b>{yes_count}</b>")

    def _apply_filters(self) -> None:
        for row in range(self.table.rowCount()):
            row_matches = True
            for col, filter_value in self._column_filters.items():
                if filter_value is None:
                    continue
                item = self.table.item(row, col)
                cell_text = item.text().strip() if item is not None else ""
                normalized = self._BLANK_FILTER if cell_text == "" else cell_text
                if normalized not in filter_value:
                    row_matches = False
                    break
            self.table.setRowHidden(row, not row_matches)

        self._refresh_header_labels()
        self._update_selection_count_label()
        self._update_hsi_count_label()

    def _on_reset_filter_clicked(self) -> None:
        for column in self._column_filters:
            self._column_filters[column] = None
        self._apply_filters()

    def _on_export_table_clicked(self) -> None:
        import os as _os
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
        stem = self._source_stem if self._source_stem else "manual_entry"
        cu_tag = f"cu{self._cleanup_mode}"
        _filename = f"rsd_{stem}_{cu_tag}_{timestamp}.xlsx"
        _start_path = _os.path.join(self._export_dir, _filename) if self._export_dir else _filename
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Reformatted SAP Data",
            _start_path,
            "Excel Files (*.xlsx)",
        )
        if not output_path:
            return

        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"

        header = [
            "Head Bom Mat",
            "HSI",
            "BOM COMPONENT",
            "Component Desc",
            "Basic Number",
            "BC",
            "Basic Name",
        ]

        self.export_progress.setVisible(True)
        self.export_progress.setValue(10)
        QApplication.processEvents()

        try:
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Reformatted SAP Data"

            output_rows: list[list[str]] = []
            for row in range(self.table.rowCount()):
                if self.table.isRowHidden(row):
                    continue
                values: list[str] = []
                row_has_value = False
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    text = item.text().strip() if item is not None else ""
                    if text:
                        row_has_value = True
                    values.append(text)
                if row_has_value:
                    output_rows.append(values)

            self.export_progress.setValue(45)
            QApplication.processEvents()

            worksheet.append(header)
            for values in output_rows:
                worksheet.append(values)

            accent_fill = PatternFill(fill_type="solid", fgColor="8A244B")
            white_font = Font(color="FFFFFF", bold=True)
            normal_font = Font(color="111111")
            border = Border(
                left=Side(style="thin", color="B8B8B8"),
                right=Side(style="thin", color="B8B8B8"),
                top=Side(style="thin", color="B8B8B8"),
                bottom=Side(style="thin", color="B8B8B8"),
            )

            max_row = worksheet.max_row
            max_col = worksheet.max_column

            for col in range(1, max_col + 1):
                cell = worksheet.cell(row=1, column=col)
                cell.fill = accent_fill
                cell.font = white_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

            for row in range(2, max_row + 1):
                for col in range(1, max_col + 1):
                    data_cell = worksheet.cell(row=row, column=col)
                    data_cell.font = normal_font
                    data_cell.alignment = Alignment(horizontal="left", vertical="center")
                    data_cell.border = border

            self.export_progress.setValue(80)
            QApplication.processEvents()

            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = f"A1:{worksheet.cell(row=1, column=max_col).coordinate}"

            for col_cells in worksheet.columns:
                col_letter = col_cells[0].column_letter
                max_len = 0
                for cell in col_cells:
                    value = "" if cell.value is None else str(cell.value)
                    if len(value) > max_len:
                        max_len = len(value)
                worksheet.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 56)

            # ── Grouping Count sheet (second sheet) ──────────────────────
            gc_counts: dict[str, int] = {}
            gc_bn_map: dict[str, dict[str, int]] = {}
            gc_bname_map: dict[str, dict[str, int]] = {}
            for row_vals in output_rows:
                comb = row_vals[5].strip() if len(row_vals) > 5 else ""
                if not comb:
                    continue
                gc_counts[comb] = gc_counts.get(comb, 0) + 1
                bn = row_vals[4].strip() if len(row_vals) > 4 else ""
                gc_bn_map.setdefault(comb, {})
                gc_bn_map[comb][bn] = gc_bn_map[comb].get(bn, 0) + 1
                bname = row_vals[6].strip() if len(row_vals) > 6 else ""
                gc_bname_map.setdefault(comb, {})
                gc_bname_map[comb][bname] = gc_bname_map[comb].get(bname, 0) + 1

            gc_rows: list[tuple[str, int, str, str]] = []
            for comb, cnt in gc_counts.items():
                best_bn = max(gc_bn_map[comb].items(), key=lambda p: (p[1], p[0]))[0] if gc_bn_map.get(comb) else ""
                best_name = max(gc_bname_map[comb].items(), key=lambda p: (p[1], p[0]))[0] if gc_bname_map.get(comb) else ""
                gc_rows.append((comb, cnt, best_bn, best_name))

            def _gc_sort(r):
                m = re.search(r"(\d+)", r[0])
                return (-r[1], int(m.group(1)) if m else 10**9, r[0])

            gc_rows.sort(key=_gc_sort)

            grouping_ws = workbook.create_sheet("Grouping Count")
            grouping_ws.append(["BC", "Count", "Basic Number", "Basic Name"])
            for comb_name, count, bn, bname in gc_rows:
                grouping_ws.append([comb_name, count, bn, bname])

            gc_max_row = grouping_ws.max_row
            gc_max_col = grouping_ws.max_column
            for col in range(1, gc_max_col + 1):
                cell = grouping_ws.cell(row=1, column=col)
                cell.fill = accent_fill
                cell.font = white_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border
            for row in range(2, gc_max_row + 1):
                for col in range(1, gc_max_col + 1):
                    c = grouping_ws.cell(row=row, column=col)
                    c.font = normal_font
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    c.border = border
            grouping_ws.freeze_panes = "A2"
            grouping_ws.auto_filter.ref = f"A1:{grouping_ws.cell(row=1, column=gc_max_col).coordinate}"
            for col_cells in grouping_ws.columns:
                letter = col_cells[0].column_letter
                mx = max((len(str(cell.value)) for cell in col_cells if cell.value is not None), default=0)
                grouping_ws.column_dimensions[letter].width = min(max(mx + 2, 12), 56)

            workbook.save(output_path)
            self.export_progress.setValue(100)
            QApplication.processEvents()
        except Exception as exc:
            self.export_progress.setVisible(False)
            msg = QMessageBox(self)
            msg.setWindowTitle("Error")
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText(f"Failed to export file. {exc}")
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.exec()
            return

        self.export_progress.setVisible(False)

        msg = QMessageBox(self)
        msg.setWindowTitle("Done")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText("Reformatted SAP data exported successfully.")
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def _on_grouping_count_clicked(self) -> None:
        grouping_counts: dict[str, int] = {}
        grouping_basic_numbers: dict[str, dict[str, int]] = {}
        grouping_basic_names: dict[str, dict[str, int]] = {}
        combination_column = 5
        basic_number_column = 4
        basic_name_column = 6

        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            item = self.table.item(row, combination_column)
            combination = item.text().strip() if item is not None else ""
            if not combination:
                continue
            grouping_counts[combination] = grouping_counts.get(combination, 0) + 1

            basic_number_item = self.table.item(row, basic_number_column)
            basic_number = basic_number_item.text().strip() if basic_number_item is not None else ""
            if combination not in grouping_basic_numbers:
                grouping_basic_numbers[combination] = {}
            grouping_basic_numbers[combination][basic_number] = grouping_basic_numbers[combination].get(basic_number, 0) + 1

            basic_name_item = self.table.item(row, basic_name_column)
            basic_name = basic_name_item.text().strip() if basic_name_item is not None else ""
            if combination not in grouping_basic_names:
                grouping_basic_names[combination] = {}
            grouping_basic_names[combination][basic_name] = grouping_basic_names[combination].get(basic_name, 0) + 1

        grouped_rows: list[tuple[str, int, str, str]] = []
        for combination, count in grouping_counts.items():
            best_basic_number = ""
            if combination in grouping_basic_numbers and grouping_basic_numbers[combination]:
                best_basic_number = max(
                    grouping_basic_numbers[combination].items(),
                    key=lambda pair: (pair[1], pair[0]),
                )[0]

            best_basic_name = ""
            if combination in grouping_basic_names and grouping_basic_names[combination]:
                best_basic_name = max(
                    grouping_basic_names[combination].items(),
                    key=lambda pair: (pair[1], pair[0]),
                )[0]
            grouped_rows.append((combination, count, best_basic_number, best_basic_name))

        sorted_counts = sorted(
            grouped_rows,
            key=lambda row: (-row[1], self._combination_sort_key(row[0])),
        )

        self._grouping_count_dialog = _MapperGroupingCountDialog(sorted_counts, self)
        self._grouping_count_dialog.show()
        self._grouping_count_dialog.raise_()
        self._grouping_count_dialog.activateWindow()

    @staticmethod
    def _combination_sort_key(value: str) -> tuple[int, str]:
        # "comb 0" sorts first (numeric), then BC codes sort alphabetically.
        match = re.search(r"(\d+)", value)
        if match:
            return int(match.group(1)), value
        return 10**9, value


class _MapperGroupingCountDialog(QDialog):
    def __init__(self, grouping_counts: list[tuple[str, int, str, str]], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Grouping Count")
        self.resize(420, 560)
        self.setMinimumSize(360, 460)

        self._grouping_counts = grouping_counts

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        header_row = QHBoxLayout()
        header_row.setSpacing(8)

        title = QLabel("Grouping Count")
        title.setObjectName("mapperReformattedTitle")
        header_row.addWidget(title, 0)

        self.btn_export = QPushButton("Export")
        self.btn_export.setObjectName("packshotUpdateRowsBtn")
        header_row.addWidget(self.btn_export, 0)

        header_row.addStretch(1)

        layout.addLayout(header_row)

        self.table = _ClipboardTableWidget(max(len(grouping_counts), 1), 4, self)
        self.table._paste_disabled = True
        self.table.setObjectName("mapperReformattedTable")
        self.table.setHorizontalHeaderLabels(["BC", "Count", "Basic Number", "Basic Name"])
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.horizontalHeader().setDefaultSectionSize(140)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        if grouping_counts:
            for row_idx, (name, count, basic_number, basic_name) in enumerate(grouping_counts):
                self.table.setItem(row_idx, 0, QTableWidgetItem(name))
                self.table.setItem(row_idx, 1, QTableWidgetItem(str(count)))
                self.table.setItem(row_idx, 2, QTableWidgetItem(basic_number))
                self.table.setItem(row_idx, 3, QTableWidgetItem(basic_name))
        else:
            self.table.setItem(0, 0, QTableWidgetItem(""))
            self.table.setItem(0, 1, QTableWidgetItem("0"))
            self.table.setItem(0, 2, QTableWidgetItem(""))
            self.table.setItem(0, 3, QTableWidgetItem(""))

        layout.addWidget(self.table)

        self.btn_export.clicked.connect(self._on_export_clicked)

        self.setStyleSheet(
            """
            QDialog {
                background-color: #F4F4F4;
            }

            QLabel#mapperReformattedTitle {
                color: #8A244B;
                font-family: "Segoe UI";
                font-size: 16px;
                font-weight: 800;
            }

            QTableWidget#mapperReformattedTable {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                gridline-color: #B8B8B8;
                selection-background-color: #DCE6F5;
                selection-color: #111111;
                font-family: "Segoe UI";
                font-size: 11px;
            }

            QTableWidget#mapperReformattedTable QHeaderView::section {
                background-color: #8A244B;
                color: #FFFFFF;
                border: 1px solid #7D8694;
                padding: 6px 8px;
                font-weight: 700;
            }

            QTableWidget#mapperReformattedTable QTableCornerButton::section {
                background-color: #8A244B;
                border: 1px solid #7D8694;
            }

            QPushButton#packshotUpdateRowsBtn {
                background-color: #9EA3AB;
                color: #000000;
                border: 1px solid #8B9098;
                border-radius: 8px;
                min-height: 30px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 12px;
                font-weight: 600;
            }

            QPushButton#packshotUpdateRowsBtn:pressed {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #111F35;
            }
            """
        )

    def _on_export_clicked(self) -> None:
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
        output_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Save Grouping Count",
            f"grouping_count_{timestamp}.xlsx",
            "Excel Files (*.xlsx);;Text Files (*.txt)",
        )
        if not output_path:
            return

        is_txt = selected_filter.startswith("Text Files") or output_path.lower().endswith(".txt")
        is_xlsx = selected_filter.startswith("Excel Files") or output_path.lower().endswith(".xlsx")

        if not is_txt and not is_xlsx:
            output_path += ".xlsx"
            is_xlsx = True
        elif is_txt and not output_path.lower().endswith(".txt"):
            output_path += ".txt"
        elif is_xlsx and not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"

        try:
            if is_txt:
                with open(output_path, "w", encoding="utf-8") as handle:
                    handle.write("BC\tCount\tBasic Number\tBasic Name\n")
                    for name, count, basic_number, basic_name in self._grouping_counts:
                        handle.write(f"{name}\t{count}\t{basic_number}\t{basic_name}\n")
            else:
                workbook = Workbook()
                worksheet = workbook.active
                worksheet.title = "Grouping Count"

                worksheet.append(["BC", "Count", "Basic Number", "Basic Name"])
                for name, count, basic_number, basic_name in self._grouping_counts:
                    worksheet.append([name, str(count), basic_number, basic_name])

                accent_fill = PatternFill(fill_type="solid", fgColor="8A244B")
                white_font = Font(color="FFFFFF", bold=True)
                normal_font = Font(color="111111")
                border = Border(
                    left=Side(style="thin", color="B8B8B8"),
                    right=Side(style="thin", color="B8B8B8"),
                    top=Side(style="thin", color="B8B8B8"),
                    bottom=Side(style="thin", color="B8B8B8"),
                )

                for col in range(1, 5):
                    header_cell = worksheet.cell(row=1, column=col)
                    header_cell.fill = accent_fill
                    header_cell.font = white_font
                    header_cell.alignment = Alignment(horizontal="center", vertical="center")
                    header_cell.border = border

                for row in range(2, worksheet.max_row + 1):
                    for col in range(1, 5):
                        body_cell = worksheet.cell(row=row, column=col)
                        body_cell.font = normal_font
                        if col == 2:
                            body_cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            body_cell.alignment = Alignment(horizontal="left", vertical="center")
                        body_cell.border = border

                worksheet.freeze_panes = "A2"
                worksheet.auto_filter.ref = "A1:D1"
                worksheet.column_dimensions["A"].width = 30
                worksheet.column_dimensions["B"].width = 12
                worksheet.column_dimensions["C"].width = 30
                worksheet.column_dimensions["D"].width = 48
                workbook.save(output_path)
        except Exception as exc:
            msg = QMessageBox(self)
            msg.setWindowTitle("Error")
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText(f"Failed to export file. {exc}")
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.exec()
            return

        msg = QMessageBox(self)
        msg.setWindowTitle("Done")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText("Grouping count exported successfully.")
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()


class _SettingsDialog(QDialog):
    """Settings window — Configure Root Folders and Projects."""

    _ROOT_HINT = (
        "HAT Dashboard will create a 'HAT DASHBOARD ROOT' folder at the chosen location "
        "containing a MASTER sub-folder with: Briefing Trackers, Tracker Status Collector, "
        "Thumbnail Generator, and SAP Data Reformat. "
        "Settings are saved to AppData and persist across sessions."
    )
    _PROJECT_HINT = (
        "HAT Dashboard will create the named project folder at the chosen location "
        "with sub-folders: Packshot Naming Generator, SAP Data Reformat, "
        "SAP Data Compare, and Project Review. "
        "The project folder path is saved to AppData."
    )

    def __init__(self, config: "HatConfig", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("HAT Dashboard — Settings")
        self.resize(650, 340)
        self.setMinimumSize(520, 280)
        self._config = config

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(28, 24, 28, 24)
        root_layout.setSpacing(14)

        # ── section title ──────────────────────────────────────────────
        section_label = QLabel("Configure root folders")
        section_label.setObjectName("settingsSectionTitle")
        root_layout.addWidget(section_label)

        # ── mode radio buttons ─────────────────────────────────────────
        radio_row = QHBoxLayout()
        radio_row.setSpacing(24)
        self._radio_root = QRadioButton("Setup root")
        self._radio_root.setObjectName("settingsRadio")
        self._radio_root.setChecked(True)
        self._radio_project = QRadioButton("Setup a project")
        self._radio_project.setObjectName("settingsRadio")
        radio_row.addWidget(self._radio_root)
        radio_row.addWidget(self._radio_project)
        radio_row.addStretch(1)
        root_layout.addLayout(radio_row)

        # ── location row ───────────────────────────────────────────────
        loc_row = QHBoxLayout()
        loc_row.setSpacing(10)
        self._lbl_location = QLabel("Root folder location:")
        self._lbl_location.setObjectName("settingsLabel")
        loc_row.addWidget(self._lbl_location, 0)

        self._btn_browse = QPushButton("Browse…")
        self._btn_browse.setObjectName("settingsBrowseBtn")
        self._btn_browse.setFixedHeight(34)
        self._btn_browse.setMinimumWidth(90)
        loc_row.addWidget(self._btn_browse, 0)

        self._input_location = QLineEdit()
        self._input_location.setObjectName("settingsLineEdit")
        self._input_location.setPlaceholderText("Select a folder…")
        self._input_location.setText(config.root_folder())
        loc_row.addWidget(self._input_location, 1)
        root_layout.addLayout(loc_row)

        # ── project name row (project mode only) ──────────────────────
        self._proj_name_row = QWidget()
        proj_name_layout = QHBoxLayout(self._proj_name_row)
        proj_name_layout.setContentsMargins(0, 0, 0, 0)
        proj_name_layout.setSpacing(10)
        lbl_proj = QLabel("Project name:")
        lbl_proj.setObjectName("settingsLabel")
        proj_name_layout.addWidget(lbl_proj, 0)
        self._input_project_name = QLineEdit()
        self._input_project_name.setObjectName("settingsLineEdit")
        self._input_project_name.setPlaceholderText("Enter project folder name…")
        proj_name_layout.addWidget(self._input_project_name, 1)
        self._proj_name_row.setVisible(False)
        root_layout.addWidget(self._proj_name_row)

        # ── hint ───────────────────────────────────────────────────────
        self._hint = QLabel(self._ROOT_HINT)
        self._hint.setObjectName("settingsHint")
        self._hint.setWordWrap(True)
        root_layout.addWidget(self._hint)

        # ── config path label ──────────────────────────────────────────
        self._path_label = QLabel(f"Config file: {config.config_path()}")
        self._path_label.setObjectName("settingsPathLabel")
        self._path_label.setWordWrap(True)
        root_layout.addWidget(self._path_label)

        root_layout.addStretch(1)

        # ── bottom buttons ─────────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)

        self._btn_open_folder = QPushButton("Open config folder")
        self._btn_open_folder.setObjectName("settingsSecondaryBtn")
        self._btn_open_folder.setFixedHeight(34)
        btn_row.addWidget(self._btn_open_folder)

        self._btn_save = QPushButton("Save and Apply")
        self._btn_save.setObjectName("settingsSaveBtn")
        self._btn_save.setFixedHeight(34)
        self._btn_save.setMinimumWidth(120)
        btn_row.addWidget(self._btn_save)

        root_layout.addLayout(btn_row)

        # ── connections ────────────────────────────────────────────────
        self._radio_root.toggled.connect(self._sync_mode)
        self._btn_browse.clicked.connect(self._browse_folder)
        self._btn_save.clicked.connect(self._save)
        self._btn_open_folder.clicked.connect(self._open_config_folder)

        self.setStyleSheet("""
            QDialog { background-color: #F4F4F4; }
            #settingsSectionTitle {
                font-family: 'Segoe UI'; font-size: 17px; font-weight: 700;
                color: #111F35;
            }
            #settingsLabel {
                font-family: 'Segoe UI'; font-size: 13px; color: #333333;
            }
            #settingsHint {
                font-family: 'Segoe UI'; font-size: 11px; color: #666666;
            }
            #settingsPathLabel {
                font-family: 'Segoe UI'; font-size: 10px; color: #999999;
            }
            #settingsRadio {
                font-family: 'Segoe UI'; font-size: 13px; color: #333333;
            }
            #settingsLineEdit {
                background-color: #FFFFFF; color: #111111;
                border: 1px solid #BBBBBB; border-radius: 6px;
                padding: 0 8px; min-height: 34px;
                font-family: 'Segoe UI'; font-size: 12px;
            }
            #settingsBrowseBtn, #settingsSecondaryBtn {
                background-color: #9EA3AB; color: #000000;
                border: 1px solid #8B9098; border-radius: 6px;
                padding: 0 12px;
                font-family: 'Segoe UI'; font-size: 12px; font-weight: 600;
            }
            #settingsBrowseBtn:hover, #settingsSecondaryBtn:hover {
                background-color: #ACB1B8;
            }
            #settingsBrowseBtn:pressed, #settingsSecondaryBtn:pressed {
                background-color: #111F35; color: #ffffff;
            }
            #settingsSaveBtn {
                background-color: #111F35; color: #ffffff;
                border: none; border-radius: 6px; padding: 0 16px;
                font-family: 'Segoe UI'; font-size: 13px; font-weight: 700;
            }
            #settingsSaveBtn:pressed { background-color: #D02752; }
        """)

    # ── slots ──────────────────────────────────────────────────────────────

    def _sync_mode(self, root_checked: bool) -> None:
        if root_checked:
            self._lbl_location.setVisible(True)
            self._btn_browse.setVisible(True)
            self._input_location.setVisible(True)
            self._lbl_location.setText("Root folder location:")
            self._input_location.setPlaceholderText("Select a folder…")
            self._input_location.setText(self._config.root_folder())
            self._proj_name_row.setVisible(False)
            self._hint.setText(self._ROOT_HINT)
        else:
            self._lbl_location.setVisible(False)
            self._btn_browse.setVisible(False)
            self._input_location.setVisible(False)
            self._proj_name_row.setVisible(True)
            self._hint.setText(self._PROJECT_HINT)
        self.adjustSize()

    def _browse_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "Select folder", self._input_location.text() or ""
        )
        if folder:
            self._input_location.setText(folder)

    def _save(self) -> None:
        from pathlib import Path

        if self._radio_root.isChecked():
            location = self._input_location.text().strip()
            if not location:
                QMessageBox.warning(self, "No folder selected",
                                    "Please select a folder before saving.")
                return
            if not Path(location).exists():
                QMessageBox.warning(self, "Folder not found",
                                    f"The folder does not exist:\n{location}")
                return
            created = self._config.create_folder_hierarchy(location)
            self._config.set_root_folder(location)
            self._config.save()
            msg = QMessageBox(self)
            msg.setWindowTitle("Settings saved")
            msg.setIcon(QMessageBox.Icon.Information)
            if created:
                msg.setText(
                    f"Root folder set to:\n{location}\n\n"
                    f"Created {len(created)} new sub-folder(s)."
                )
            else:
                msg.setText(f"Root folder set to:\n{location}\n\nAll folders already exist.")
            msg.exec()
        else:
            project_name = self._input_project_name.text().strip()
            if not project_name:
                QMessageBox.warning(self, "No project name",
                                    "Please enter a project name.")
                return
            root = self._config.root_folder()
            if not root or not Path(root).exists():
                QMessageBox.warning(self, "Root folder not found",
                                    "Root folder not found, configure in Setup root.")
                return
            # Place project at <root>/HAT DASHBOARD ROOT/ — same level as MASTER
            location = str(Path(root) / "HAT DASHBOARD ROOT")
            created, project_path = self._config.create_project_folder_hierarchy(
                location, project_name
            )
            self._config.save()
            msg = QMessageBox(self)
            msg.setWindowTitle("Project created")
            msg.setIcon(QMessageBox.Icon.Information)
            if created:
                msg.setText(
                    f"Project folder created:\n{project_path}\n\n"
                    f"Created {len(created)} sub-folder(s)."
                )
            else:
                msg.setText(
                    f"Project folder:\n{project_path}\n\nAll folders already exist."
                )
            msg.exec()

        self._path_label.setText(f"Config file: {self._config.config_path()}")

    def _open_config_folder(self) -> None:
        import subprocess
        folder = str(self._config.config_path().parent)
        subprocess.Popen(f'explorer "{folder}"')


class _SearchMultiValueDialog(QDialog):
    """Table-based multi-value input for a single expanded-search field."""

    def __init__(
        self,
        field_label: str,
        initial_values: list[str],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Multiple Values — {field_label}")
        self.resize(420, 440)
        self.setMinimumSize(320, 300)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        title = QLabel(f"Values for: <b>{field_label}</b>")
        title.setObjectName("searchMvTitle")
        layout.addWidget(title)

        # Toolbar
        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)
        self._btn_add   = QPushButton("+ Add Row")
        self._btn_del   = QPushButton("Delete")
        self._btn_reset = QPushButton("Reset")
        self._btn_paste = QPushButton("Paste")
        for b in (self._btn_add, self._btn_del, self._btn_reset, self._btn_paste):
            b.setObjectName("searchMvBtn")
            btn_row.addWidget(b)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        # Single-column table
        self.table = QTableWidget(0, 1)
        self.table.setObjectName("searchMvTable")
        self.table.setHorizontalHeaderLabels(["Value"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self.table)

        # Seed with initial values (always at least one blank row)
        seeds = [v for v in initial_values if v.strip()] or ["", "", "", "", ""]
        for v in seeds:
            self._add_row(v)

        # OK / Cancel
        ok_row = QHBoxLayout()
        ok_row.setSpacing(8)
        self._btn_ok     = QPushButton("OK")
        self._btn_cancel = QPushButton("Cancel")
        self._btn_ok.setObjectName("collectorRunBtn")
        self._btn_cancel.setObjectName("collectorGrayBtn")
        self._btn_ok.setFixedHeight(38)
        self._btn_cancel.setFixedHeight(38)
        self._btn_ok.setMinimumWidth(90)
        self._btn_cancel.setMinimumWidth(90)
        ok_row.addStretch(1)
        ok_row.addWidget(self._btn_cancel)
        ok_row.addWidget(self._btn_ok)
        layout.addLayout(ok_row)

        self._btn_add.clicked.connect(self._on_add)
        self._btn_del.clicked.connect(self._on_delete)
        self._btn_reset.clicked.connect(self._on_reset)
        self._btn_paste.clicked.connect(self._on_paste)
        self._btn_ok.clicked.connect(self.accept)
        self._btn_cancel.clicked.connect(self.reject)

        # Allow Ctrl+V directly on the table
        orig_key_press = self.table.keyPressEvent
        def _table_key_press(event, _orig=orig_key_press):
            if event.matches(QKeySequence.StandardKey.Paste):
                self._on_paste()
            elif event.matches(QKeySequence.StandardKey.Copy):
                selected = self.table.selectedIndexes()
                lines = []
                for idx in sorted(selected, key=lambda i: i.row()):
                    item = self.table.item(idx.row(), 0)
                    lines.append(item.text() if item else "")
                QApplication.clipboard().setText("\n".join(lines))
            else:
                _orig(event)
        self.table.keyPressEvent = _table_key_press

        self.setStyleSheet("""
            QDialog { background-color: #F4F4F4; }
            QLabel#searchMvTitle {
                color: #111F35;
                font-family: "Segoe UI"; font-size: 14px; font-weight: 700;
            }
            QPushButton#searchMvBtn {
                background-color: #9EA3AB;
                color: #000000;
                border: 1px solid #8B9098;
                border-radius: 8px;
                padding: 0 12px;
                min-height: 30px;
                font-family: "Segoe UI"; font-size: 12px; font-weight: 600;
            }
            QPushButton#searchMvBtn:hover   { background-color: #ACB1B8; }
            QPushButton#searchMvBtn:pressed { background-color: #111F35; color: #FFFFFF; border: 1px solid #111F35; }
            QTableWidget#searchMvTable {
                background-color: #FFFFFF;
                color: #111111;
                border: 1px solid #A9A9A9;
                gridline-color: #D0D0D0;
                selection-background-color: #DCE6F5;
                font-family: "Segoe UI"; font-size: 13px;
            }
            QTableWidget#searchMvTable QHeaderView::section {
                background-color: #111F35;
                color: #FFFFFF;
                border: 1px solid #7D8694;
                padding: 4px 8px;
                font-weight: 700;
                font-family: "Segoe UI"; font-size: 12px;
            }
        """)

    # ── helpers ──────────────────────────────────────────────────────────

    def _add_row(self, value: str = "") -> None:
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(value))

    def _on_add(self) -> None:
        self._add_row()
        self.table.scrollToBottom()
        last = self.table.rowCount() - 1
        self.table.setCurrentCell(last, 0)
        self.table.editItem(self.table.item(last, 0))

    def _on_delete(self) -> None:
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)

    def _on_reset(self) -> None:
        self.table.setRowCount(0)
        self._add_row()

    def _on_paste(self) -> None:
        """Paste clipboard text into the table.

        Handles all common Excel/spreadsheet copy formats:
        - Single or multi-column Excel paste → one value per *row* (first column only)
        - Each Excel row is kept intact as a single value (commas within a cell are preserved)
        - Plain newline-separated text       → one value per line
        - Blank lines are discarded
        """
        text = QApplication.clipboard().text().strip()
        if not text:
            return

        values: list[str] = []
        for line in text.splitlines():
            # For Excel multi-column pastes the columns are tab-separated;
            # only take the first column but keep any commas/text within that cell.
            if "\t" in line:
                cell = line.split("\t")[0].strip()
            else:
                cell = line.strip()
            if cell:
                values.append(cell)

        if not values:
            return

        # If there is currently a selected cell, insert at that position;
        # otherwise replace all rows
        selected_rows = sorted({idx.row() for idx in self.table.selectedIndexes()})
        if selected_rows and len(values) > 0:
            insert_at = selected_rows[0]
            # Remove selected rows first, then insert new values starting there
            for r in sorted(selected_rows, reverse=True):
                self.table.removeRow(r)
            for offset, val in enumerate(values):
                self.table.insertRow(insert_at + offset)
                self.table.setItem(insert_at + offset, 0, QTableWidgetItem(val))
        else:
            self.table.setRowCount(0)
            for val in values:
                self._add_row(val)

    def get_values(self) -> list[str]:
        """Return non-empty values entered in the table."""
        result = []
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            val = item.text().strip() if item else ""
            if val:
                result.append(val)
        return result


class _SearchResultWindow(QDialog):
    """Non-modal window showing search results as tiles or table."""

    MINIMAL_COLS: list[str] = [
        "IDH", "Pack Name", "Pack Type", "Pack Size", "Basic", "Project Name",
    ]
    ALL_COLS: list[str] = [
        "IDH", "Pack Name", "Pack Type", "Pack Size", "Basic", "Project Name",
        "SBU", "Label Size", "Color",
    ]

    # Detail fields for tile descriptions
    MINIMAL_TILE_FIELDS: list[str] = ["IDH", "Pack Name", "Basic", "Image Name", "Build Type"]
    ALL_TILE_FIELDS: list[str] = [
        "IDH", "Pack Name", "Basic", "Image Name", "Build Type", "SBU",
        "Pack Type", "Pack Size", "Label Size", "Project Name",
    ]

    _ACTIVE_HEADER_BG   = "#F63049"
    _ACTIVE_HEADER_FG   = "#000000"
    _INACTIVE_HEADER_BG = "#555555"
    _INACTIVE_HEADER_FG = "#AAAAAA"

    _HIGHLIGHT_COLOR = "#E8151B"   # red used to highlight matched search values

    def __init__(self, result_data: dict, cfg: "HatConfig", use_root_folders: bool = False, parent=None) -> None:
        super().__init__(parent)
        self._result_data = result_data
        self._cfg = cfg
        self._use_root_folders = use_root_folders
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.setWindowTitle("Search Results")
        self.setMinimumSize(900, 640)
        self.resize(1160, 760)

        # Gather thumbnail folders: try all known thumbnail locations as fallbacks
        from search import find_matching_images
        idh_set = {v.strip() for v in result_data.get("idh_list", []) if v.strip()}
        thumb_folder = cfg.search_thumbnails_folder()
        fallbacks = [cfg.thumbnail_output(), cfg.thumbnail_input()]
        self._images: list[dict] = (
            find_matching_images(idh_set, thumb_folder, fallback_folders=fallbacks)
            if idh_set else []
        )

        # Map idh → image info for quick tile lookup
        self._image_by_idh: dict[str, dict] = {
            info["idh"]: info for info in self._images if info.get("idh")
        }

        self._build_ui()

    # ── UI construction ──────────────────────────────────────────────────

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 12, 16, 12)
        root.setSpacing(8)

        # ── Header bar ───────────────────────────────────────────────────
        hdr = QHBoxLayout()
        hdr.setSpacing(16)

        # Details radio group
        details_lbl = QLabel("Details:")
        details_lbl.setStyleSheet("font-weight: 700; font-size: 13px;")
        hdr.addWidget(details_lbl)

        self._radio_minimal = QRadioButton("Minimal")
        self._radio_all     = QRadioButton("All")
        self._radio_minimal.setChecked(True)
        grp_details = QButtonGroup(self)
        grp_details.addButton(self._radio_minimal)
        grp_details.addButton(self._radio_all)
        hdr.addWidget(self._radio_minimal)
        hdr.addWidget(self._radio_all)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.VLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        hdr.addWidget(sep)

        # Format radio group
        fmt_lbl = QLabel("Format:")
        fmt_lbl.setStyleSheet("font-weight: 700; font-size: 13px;")
        hdr.addWidget(fmt_lbl)

        self._radio_tiles = QRadioButton("Tiles")
        self._radio_table = QRadioButton("Table")
        self._radio_tiles.setChecked(True)
        grp_fmt = QButtonGroup(self)
        grp_fmt.addButton(self._radio_tiles)
        grp_fmt.addButton(self._radio_table)
        hdr.addWidget(self._radio_tiles)
        hdr.addWidget(self._radio_table)

        # Item count label (shown only when sample limit is "All")
        self._item_count_lbl = QLabel()
        self._item_count_lbl.setStyleSheet(
            "font-weight: 700; font-size: 12px; color: #1A5276;"
            " background: #D6EAF8; border-radius: 4px; padding: 2px 8px;"
        )
        show_count = self._result_data.get("show_item_count", False)
        self._item_count_lbl.setVisible(show_count)
        hdr.addWidget(self._item_count_lbl)

        hdr.addStretch(1)

        export_btn = QPushButton("Export search results")
        export_btn.setObjectName("collectorRunBtn")
        export_btn.setFixedHeight(36)
        export_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        hdr.addWidget(export_btn)
        root.addLayout(hdr)

        # ── "Search results for:" label ──────────────────────────────────
        summary = self._result_data.get("search_summary", "")
        if summary:
            summary_lbl = QLabel(f"<b>Search results for:</b> {summary}")
            summary_lbl.setStyleSheet(
                "font-size: 12px; color: #333333; padding: 2px 0 6px 0;"
            )
            summary_lbl.setWordWrap(True)
            root.addWidget(summary_lbl)

        # ── Scrollable body ──────────────────────────────────────────────
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.Shape.NoFrame)
        root.addWidget(self._scroll)

        # ── Wire signals ─────────────────────────────────────────────────
        self._radio_minimal.toggled.connect(self._on_options_changed)
        self._radio_tiles.toggled.connect(self._on_options_changed)
        export_btn.clicked.connect(self._export)

        self._refresh_body()

    def _on_options_changed(self) -> None:
        self._refresh_body()

    def _refresh_body(self) -> None:
        """Rebuild the scrollable content based on current Detail + Format."""
        from search import DISPLAY_COLUMN_FIELD

        old = self._scroll.takeWidget()
        if old:
            old.deleteLater()

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 4, 8, 8)
        layout.setSpacing(16)

        minimal = self._radio_minimal.isChecked()
        tiles   = self._radio_tiles.isChecked()

        rows   = self._result_data.get("rows", [])
        active = self._result_data.get("active_filters", set())
        source = self._result_data.get("source", "")
        is_rsd = source == "RSD (master)"
        groups = self._result_data.get("groups")        # list[dict] or None

        # Update item count label
        if self._result_data.get("show_item_count", False):
            total = sum(len(g.get("rows", [])) for g in groups) if groups else len(rows)
            self._item_count_lbl.setText(f"Item Count: {total}")
            self._item_count_lbl.setVisible(True)
        else:
            self._item_count_lbl.setVisible(False)

        if not rows:
            no_res = QLabel("No results found.")
            no_res.setStyleSheet("color: #888888; font-size: 13px; padding: 4px 0;")
            layout.addWidget(no_res)
        elif groups:
            # Grouped display: one section per search-cell-value
            data_cols = [c for c in (self.MINIMAL_COLS if minimal else self.ALL_COLS)
                         if not (is_rsd and c == "Pack Size")]
            if tiles:
                self._build_grouped_tiles_view(layout, groups, minimal, active, is_rsd)
            else:
                self._build_grouped_table_view(layout, groups, data_cols, active)
                self._scroll.setWidget(container)
                return
        elif tiles:
            self._build_tiles_view(layout, rows, minimal, active, is_rsd)
        else:
            cols = [c for c in (self.MINIMAL_COLS if minimal else self.ALL_COLS)
                    if not (is_rsd and c == "Pack Size")]
            self._build_table_view(layout, rows, cols, active)
            self._scroll.setWidget(container)
            return

        layout.addStretch(1)
        self._scroll.setWidget(container)

    # ── Tiles view ───────────────────────────────────────────────────────

    def _build_tiles_view(
        self,
        layout: QVBoxLayout,
        rows: list[dict],
        minimal: bool,
        active: set,
        is_rsd: bool = False,
    ) -> None:
        base_fields = self.MINIMAL_TILE_FIELDS if minimal else self.ALL_TILE_FIELDS
        tile_fields = [f for f in base_fields if not (is_rsd and f == "Pack Size")]
        active_values: list[str] = self._collect_active_values(active)

        TILE_W, IMG_H = 260, 240
        COLS = 4

        grid = QGridLayout()
        grid.setSpacing(20)
        grid.setContentsMargins(0, 0, 0, 0)

        for i, record in enumerate(rows):
            idh = record.get("IDH", "").strip()
            img_info = self._image_by_idh.get(idh)
            has_image = img_info and img_info.get("path")

            cell_w = QWidget()
            cell_w.setFixedWidth(TILE_W)
            cell_v = QVBoxLayout(cell_w)
            cell_v.setSpacing(4)
            cell_v.setContentsMargins(0, 0, 0, 0)

            # Image area
            img_lbl = QLabel()
            img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            img_lbl.setFixedSize(TILE_W, IMG_H)

            if has_image:
                px = QPixmap(img_info["path"])
                if not px.isNull():
                    px = px.scaled(
                        TILE_W - 4, IMG_H - 4,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )
                    img_lbl.setPixmap(px)
                else:
                    img_lbl.setText("Image unavailable")
                    img_lbl.setStyleSheet(
                        "background:#E0E0E0; color:#888; font-size:11px; border:1px solid #CCC;"
                    )
            else:
                img_lbl.setText("No image found")
                img_lbl.setStyleSheet(
                    "background:#D8D8D8; color:#777; font-size:12px; border:1px solid #BBBBBB;"
                )

            cell_v.addWidget(img_lbl)

            # Text description below image
            for field_name in tile_fields:
                # "Image Name" is a virtual field — resolved from the matched image path
                if field_name == "Image Name":
                    value = Path(img_info["path"]).name if has_image else ""
                else:
                    value = record.get(field_name, "").strip()
                desc_lbl = QLabel()
                desc_lbl.setWordWrap(True)
                desc_lbl.setStyleSheet("font-size: 11px;")
                desc_lbl.setTextInteractionFlags(
                    Qt.TextInteractionFlag.TextSelectableByMouse
                    | Qt.TextInteractionFlag.TextSelectableByKeyboard
                )
                display_val = value if value else "—"
                if value and active_values:
                    highlighted = self._highlight_substrings(display_val, active_values)
                    desc_lbl.setText(f"<b>{field_name}:</b> {highlighted}")
                else:
                    desc_lbl.setText(f"<b>{field_name}:</b> {display_val}")
                cell_v.addWidget(desc_lbl)

            cell_v.addStretch(1)
            grid.addWidget(cell_w, i // COLS, i % COLS)

        wrapper = QWidget()
        wrapper.setLayout(grid)
        layout.addWidget(wrapper)

    # ── Grouped tiles view ───────────────────────────────────────────────

    def _build_grouped_tiles_view(
        self,
        layout: QVBoxLayout,
        groups: list[dict],
        minimal: bool,
        active: set,
        is_rsd: bool = False,
    ) -> None:
        """Render tiles grouped by query-cell value, with a section label per group."""
        for grp in groups:
            query_val = grp.get("query_value", "")
            rows = grp.get("rows", [])

            # Section header label
            hdr = QLabel(query_val)
            hdr.setWordWrap(True)
            hdr.setStyleSheet(
                "font-family: 'Segoe UI'; font-size: 13px; font-weight: 700;"
                " color: #111F35; background: #E4E8EE;"
                " border-left: 4px solid #111F35;"
                " padding: 6px 10px; margin-top: 4px;"
            )
            layout.addWidget(hdr)

            if rows:
                self._build_tiles_view(layout, rows, minimal, active, is_rsd)
            else:
                empty_lbl = QLabel("No results for this value.")
                empty_lbl.setStyleSheet("color: #888888; font-size: 12px; padding: 4px 12px;")
                layout.addWidget(empty_lbl)

            # Divider between groups
            div = QFrame()
            div.setFrameShape(QFrame.Shape.HLine)
            div.setFrameShadow(QFrame.Shadow.Sunken)
            div.setStyleSheet("color: #CCCCCC; margin: 8px 0;")
            layout.addWidget(div)

    # ── Grouped table view ───────────────────────────────────────────────

    def _build_grouped_table_view(
        self,
        layout: QVBoxLayout,
        groups: list[dict],
        cols: list[str],
        active: set,
    ) -> None:
        """Render a single table with a leading 'Search Value' column."""
        from search import DISPLAY_COLUMN_FIELD

        # Combine all rows, tagging each with its query_value
        all_rows: list[dict] = []
        for grp in groups:
            qv = grp.get("query_value", "")
            for rec in grp.get("rows", []):
                all_rows.append({**rec, "__query__": qv})

        full_cols = ["Search Value"] + list(cols)
        active_values = self._collect_active_values(active)

        table = QTableWidget(len(all_rows), len(full_cols))
        table.horizontalHeader().setVisible(True)
        table.verticalHeader().setVisible(True)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        table.customContextMenuRequested.connect(
            lambda pos, t=table: self._show_table_context_menu(pos, t)
        )
        table.keyPressEvent = lambda event, t=table: self._table_key_press(event, t)
        table.setStyleSheet(
            "QHeaderView::section { padding: 4px 8px; font-weight: 700;"
            " font-family: 'Segoe UI'; font-size: 12px; border: 1px solid #7D8694; }"
        )

        # Header items — "Search Value" always active-style, others normal
        for ci, col_name in enumerate(full_cols):
            field_key = DISPLAY_COLUMN_FIELD.get(col_name)
            hitem = QTableWidgetItem(col_name)
            hitem.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            bold = QFont()
            bold.setBold(True)
            hitem.setFont(bold)
            if ci == 0:                              # Search Value column
                hitem.setBackground(QColor(self._ACTIVE_HEADER_BG))
                hitem.setForeground(QColor(self._ACTIVE_HEADER_FG))
            elif field_key and field_key in active:
                hitem.setBackground(QColor(self._ACTIVE_HEADER_BG))
                hitem.setForeground(QColor(self._ACTIVE_HEADER_FG))
            else:
                hitem.setBackground(QColor(self._INACTIVE_HEADER_BG))
                hitem.setForeground(QColor(self._INACTIVE_HEADER_FG))
            table.setHorizontalHeaderItem(ci, hitem)

        # Data rows — alternate background per group for visual separation
        _GROUP_BG = ["#FFFFFF", "#EFF3FA"]   # alternating group backgrounds
        current_query = None
        group_idx = -1
        for row_idx, rec in enumerate(all_rows):
            qv = rec.get("__query__", "")
            if qv != current_query:
                current_query = qv
                group_idx += 1
            bg = QColor(_GROUP_BG[group_idx % len(_GROUP_BG)])

            for ci, col_name in enumerate(full_cols):
                if col_name == "Search Value":
                    val = qv
                else:
                    val = rec.get(col_name, "")
                cell = QTableWidgetItem(val)
                cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                cell.setBackground(bg)
                if ci > 0 and active_values and val.strip() and any(
                    t in val.strip().lower() for t in active_values
                ):
                    cell.setForeground(QColor(self._HIGHLIGHT_COLOR))
                    f = QFont()
                    f.setBold(True)
                    cell.setFont(f)
                table.setItem(row_idx, ci, cell)

        row_h = 30
        table.verticalHeader().setDefaultSectionSize(row_h)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(table, 1)

    def _collect_active_values(self, active_filter_keys: set) -> list[str]:
        """Return list of lowercase search terms the user entered (longest first for replacement order)."""
        field_values: dict = self._result_data.get("field_values", {})
        multi_values: dict = self._result_data.get("multi_values", {})
        values: set[str] = set()
        for k in active_filter_keys:
            if k in multi_values:
                for v in multi_values[k]:
                    if v.strip():
                        values.add(v.strip().lower())
            elif k in field_values:
                v = field_values[k].strip()
                if v:
                    values.add(v.lower())
        # Also include ALL field_values and multi_values — so compact-mode values highlight too
        for v in field_values.values():
            if isinstance(v, str) and v.strip():
                values.add(v.strip().lower())
        for vlist in multi_values.values():
            for v in vlist:
                if v.strip():
                    values.add(v.strip().lower())
        # Longest terms first so shorter terms don't break already-replaced spans
        return sorted(values, key=len, reverse=True)

    def _highlight_substrings(self, text: str, terms: list[str]) -> str:
        """Return HTML string with each search term highlighted in red (case-insensitive)."""
        import html as _html
        escaped = _html.escape(text)
        for term in terms:
            if not term:
                continue
            escaped_term = _html.escape(term)
            # Replace case-insensitively using regex, wrapping matched text in a red span
            escaped = re.sub(
                re.escape(escaped_term),
                lambda m: f"<span style='color:{self._HIGHLIGHT_COLOR};font-weight:bold;'>{m.group(0)}</span>",
                escaped,
                flags=re.IGNORECASE,
            )
        return escaped

    # ── Table view ───────────────────────────────────────────────────────

    def _build_table_view(
        self,
        layout: QVBoxLayout,
        rows: list[dict],
        cols: list[str],
        active: set,
    ) -> None:
        from search import DISPLAY_COLUMN_FIELD

        table = QTableWidget(len(rows), len(cols))
        table.horizontalHeader().setVisible(True)
        table.verticalHeader().setVisible(True)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        table.customContextMenuRequested.connect(
            lambda pos, t=table: self._show_table_context_menu(pos, t)
        )
        table.keyPressEvent = lambda event, t=table: self._table_key_press(event, t)

        # Style the header via stylesheet so active columns show in red
        # Individual header items carry background/foreground for per-column color
        table.setStyleSheet(
            "QHeaderView::section { padding: 4px 8px; font-weight: 700;"
            " font-family: 'Segoe UI'; font-size: 12px; border: 1px solid #7D8694; }"
        )

        # Real header row — coloured per active-filter state
        for col_idx, col_name in enumerate(cols):
            field_key = DISPLAY_COLUMN_FIELD.get(col_name)
            hitem = QTableWidgetItem(col_name)
            hitem.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            bold = QFont()
            bold.setBold(True)
            hitem.setFont(bold)
            if field_key and field_key in active:
                hitem.setBackground(QColor(self._ACTIVE_HEADER_BG))
                hitem.setForeground(QColor(self._ACTIVE_HEADER_FG))
            else:
                hitem.setBackground(QColor(self._INACTIVE_HEADER_BG))
                hitem.setForeground(QColor(self._INACTIVE_HEADER_FG))
            table.setHorizontalHeaderItem(col_idx, hitem)

        # Data rows
        active_values = self._collect_active_values(active)
        for row_idx, record in enumerate(rows):
            for col_idx, col_name in enumerate(cols):
                val = record.get(col_name, "")
                cell = QTableWidgetItem(val)
                cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                # Highlight if any search term is a substring of the cell value
                if active_values and val.strip() and any(
                    t in val.strip().lower() for t in active_values
                ):
                    cell.setForeground(QColor(self._HIGHLIGHT_COLOR))
                    f = QFont()
                    f.setBold(True)
                    cell.setFont(f)
                table.setItem(row_idx, col_idx, cell)

        row_h = 30
        table.verticalHeader().setDefaultSectionSize(row_h)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(table, 1)

    # ── Table copy helpers ───────────────────────────────────────────────

    def _copy_table_selection(self, table: QTableWidget) -> None:
        """Copy selected cells to clipboard as tab-separated text."""
        selected = table.selectedIndexes()
        if not selected:
            return
        rows = sorted({idx.row() for idx in selected})
        cols = sorted({idx.column() for idx in selected})
        lines: list[str] = []
        for r in rows:
            cells: list[str] = []
            for c in cols:
                item = table.item(r, c)
                cells.append(item.text() if item else "")
            lines.append("\t".join(cells))
        QApplication.clipboard().setText("\n".join(lines))

    def _show_table_context_menu(self, pos, table: QTableWidget) -> None:
        from PySide6.QtWidgets import QMenu
        menu = QMenu(self)
        copy_action = menu.addAction("Copy")
        copy_action.setShortcut("Ctrl+C")
        action = menu.exec(table.viewport().mapToGlobal(pos))
        if action == copy_action:
            self._copy_table_selection(table)

    def _table_key_press(self, event, table: QTableWidget) -> None:
        if event.key() == Qt.Key.Key_C and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self._copy_table_selection(table)
        else:
            QTableWidget.keyPressEvent(table, event)

    # ── PDF export ───────────────────────────────────────────────────────

    def _build_pdf_html(self, minimal: bool) -> str:
        from search import DISPLAY_COLUMN_FIELD
        import base64

        rows   = self._result_data.get("rows", [])
        active = self._result_data.get("active_filters", set())
        cols   = self.MINIMAL_COLS if minimal else self.ALL_COLS

        p: list[str] = [
            "<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:11px;'>"
            "<h2 style='margin-bottom:8px;'>Search Results</h2>"
            "<h3 style='margin-bottom:4px;'>Text Data</h3>"
            "<table border='1' cellpadding='5' cellspacing='0' "
            "style='border-collapse:collapse;width:100%;'>"
            "<tr>"
        ]
        for col in cols:
            fk = DISPLAY_COLUMN_FIELD.get(col)
            if fk and fk in active:
                s = (f"background:{self._ACTIVE_HEADER_BG};"
                     f"color:{self._ACTIVE_HEADER_FG};font-weight:bold;")
            else:
                s = "background:#555;color:#aaa;font-weight:bold;"
            p.append(f"<th style='{s}'>{col}</th>")
        p.append("</tr>")

        for i, record in enumerate(rows):
            bg = "#f8f8f8" if i % 2 == 0 else "#ffffff"
            p.append(f"<tr style='background:{bg};'>")
            for col in cols:
                p.append(f"<td>{record.get(col, '')}</td>")
            p.append("</tr>")
        p.append("</table>")

        if self._images:
            p.append("<h3 style='margin-top:16px;'>Image Data</h3>")
            p.append("<table cellpadding='8' cellspacing='4'>")
            for i in range(0, len(self._images), 2):
                p.append("<tr>")
                for j in range(2):
                    idx = i + j
                    if idx < len(self._images):
                        info = self._images[idx]
                        try:
                            ext = Path(info["path"]).suffix.lower().lstrip(".")
                            if ext == "jpg":
                                ext = "jpeg"
                            with open(info["path"], "rb") as fh:
                                b64 = __import__("base64").b64encode(fh.read()).decode()
                            img_tag = (
                                f"<img src='data:image/{ext};base64,{b64}' "
                                f"width='220' style='display:block;'/>"
                            )
                        except Exception:
                            img_tag = "[unavailable]"
                        p.append(
                            f"<td style='vertical-align:top;'>{img_tag}"
                            f"<br/><b>Image Name:</b> {info['name']}"
                            f"<br/><span style='color:#666;'>"
                            f"<b>Folder:</b> {info['folder']}</span></td>"
                        )
                    else:
                        p.append("<td></td>")
                p.append("</tr>")
            p.append("</table>")

        p.append("</body></html>")
        return "".join(p)

    def _export(self) -> None:
        """Dispatch to PDF (Tiles mode) or Excel (Table mode)."""
        if self._radio_tiles.isChecked():
            self._export_pdf()
        else:
            self._export_excel()

    def _export_pdf(self) -> None:
        from datetime import datetime
        from pathlib import Path as _Path

        ts = datetime.now().strftime("%Y_%m_%d_%H_%M")
        default_name = f"search_result_{ts}"
        start_dir = ""
        if self._use_root_folders:
            root = self._cfg.root_folder()
            if root and _Path(root).is_dir():
                start_dir = root
        save_hint = str(_Path(start_dir) / default_name) if start_dir else default_name

        path, _ = QFileDialog.getSaveFileName(
            self, "Export Search Results", save_hint, "PDF Files (*.pdf)"
        )
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"

        minimal = self._radio_minimal.isChecked()
        try:
            self._write_pdf_reportlab(path, minimal)
            QMessageBox.information(self, "Export Complete", f"Saved to:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", str(exc))

    def _write_pdf_reportlab(self, out_path: str, minimal: bool) -> None:
        """Render search results to PDF using reportlab – image tiles with field labels."""
        from reportlab.lib.pagesizes import letter as _LETTER
        from reportlab.lib.colors import HexColor
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfgen.canvas import Canvas as _Canvas
        from PIL import Image as PILImage
        from io import BytesIO
        from pathlib import Path as _Path

        PW, PH   = _LETTER
        MARGIN   = 36.0
        AVAIL_W  = PW - 2 * MARGIN
        NCOLS    = 3
        COL_GAP  = 8.0
        ROW_GAP  = 12.0
        CELL_W   = (AVAIL_W - COL_GAP * (NCOLS - 1)) / NCOLS
        IMG_H    = 150.0
        PAD      = 5.0
        FS       = 6.5
        LINE_H   = FS + 2.5

        tile_fields = self.MINIMAL_TILE_FIELDS if minimal else self.ALL_TILE_FIELDS
        TEXT_H   = len(tile_fields) * LINE_H + PAD + 4.0
        CELL_H   = IMG_H + TEXT_H + PAD * 2

        FONT_B = "Helvetica-Bold"
        FONT_R = "Helvetica"
        C_HDR_BG  = HexColor("#111F35")
        C_BOX_BG  = HexColor("#EFF3F8")
        C_BOX_BD  = HexColor("#8FAABF")
        C_WHITE   = HexColor("#FFFFFF")
        C_DARK    = HexColor("#111111")
        C_GREY    = HexColor("#555555")
        C_PLHLD   = HexColor("#C8C8C8")
        C_PLHLD_T = HexColor("#666666")
        C_GRP_BG  = HexColor("#E4E8EE")
        C_GRP_FG  = HexColor("#111F35")

        rows    = self._result_data.get("rows", [])
        groups  = self._result_data.get("groups")
        summary = self._result_data.get("search_summary", "")

        entries: list[tuple[str, dict]] = []
        entries_groups = self._result_data.get("groups")
        if entries_groups:
            for grp in entries_groups:
                qv = grp.get("query_value", "")
                for rec in grp.get("rows", []):
                    entries.append((qv, rec))
        else:
            for rec in rows:
                entries.append(("", rec))

        cv = _Canvas(out_path, pagesize=_LETTER)

        def _load_img(p):
            if not p:
                return None
            try:
                pil = PILImage.open(p)
                if pil.mode == "RGBA":
                    bg = PILImage.new("RGB", pil.size, (255, 255, 255))
                    bg.paste(pil, mask=pil.split()[3])
                    pil = bg
                elif pil.mode not in ("RGB", "L"):
                    pil = pil.convert("RGB")
                buf = BytesIO()
                pil.save(buf, format="JPEG", quality=60, optimize=True)
                buf.seek(0)
                return ImageReader(buf)
            except Exception:
                return None

        def _placeholder(x, y, w, h):
            cv.setFillColor(C_PLHLD)
            cv.setStrokeColor(HexColor("#999999"))
            cv.setLineWidth(0.5)
            cv.rect(x, y, w, h, fill=1, stroke=1)
            cv.setFillColor(C_PLHLD_T)
            cv.setFont(FONT_B, 7.0)
            cv.drawCentredString(x + w / 2, y + h / 2 - 4, "No Image Found")

        y_cursor = [PH - MARGIN]

        def _draw_header():
            y = y_cursor[0]
            cv.setFillColor(C_HDR_BG)
            cv.rect(MARGIN, y - 22, AVAIL_W, 20, fill=1, stroke=0)
            cv.setFillColor(C_WHITE)
            cv.setFont(FONT_B, 13.0)
            cv.drawString(MARGIN + 6, y - 16, "Search Results")
            y -= 26
            if summary:
                cv.setFont(FONT_R, 8.0)
                cv.setFillColor(C_GREY)
                cv.drawString(MARGIN, y - 9, f"Search results for: {summary}")
                y -= 15
            if self._result_data.get("show_item_count", False):
                total = sum(len(g.get("rows", [])) for g in entries_groups) if entries_groups else len(entries)
                cv.setFont(FONT_B, 8.0)
                cv.setFillColor(C_GREY)
                cv.drawString(MARGIN, y - 9, f"Item Count: {total}")
                y -= 15
            y -= 4
            y_cursor[0] = y

        _draw_header()
        y_top = y_cursor[0]
        col   = 0
        active_group: str | None = None

        for section_lbl, record in entries:
            if groups and section_lbl != active_group:
                if col > 0:
                    y_top -= CELL_H + ROW_GAP
                    col = 0
                if y_top - 18 - CELL_H < MARGIN:
                    cv.showPage()
                    y_cursor[0] = PH - MARGIN
                    _draw_header()
                    y_top = y_cursor[0]
                    col = 0
                cv.setFillColor(C_GRP_BG)
                cv.rect(MARGIN, y_top - 14, AVAIL_W, 13, fill=1, stroke=0)
                cv.setFillColor(C_GRP_FG)
                cv.setLineWidth(2.5)
                cv.setStrokeColor(C_GRP_FG)
                cv.line(MARGIN, y_top - 14, MARGIN, y_top - 1)
                cv.setFont(FONT_B, 7.5)
                cv.drawString(MARGIN + 6, y_top - 10.5, section_lbl)
                y_top -= 20
                active_group = section_lbl

            if col == 0 and y_top - CELL_H < MARGIN:
                cv.showPage()
                y_cursor[0] = PH - MARGIN
                _draw_header()
                y_top = y_cursor[0]

            x = MARGIN + col * (CELL_W + COL_GAP)
            cell_y_bot = y_top - CELL_H

            # Cell background + border
            cv.setFillColor(C_BOX_BG)
            cv.setStrokeColor(C_BOX_BD)
            cv.setLineWidth(0.5)
            cv.rect(x, cell_y_bot, CELL_W, CELL_H, fill=1, stroke=1)

            # Image area
            img_x = x + PAD
            img_y = y_top - PAD - IMG_H
            img_w = CELL_W - 2 * PAD
            idh = record.get("IDH", "").strip()
            img_info = self._image_by_idh.get(idh)
            img_path = img_info["path"] if img_info and img_info.get("path") else None

            reader = _load_img(img_path)
            if reader:
                ow, oh = reader.getSize()
                ratio = min(img_w / ow, IMG_H / oh)
                dw, dh = ow * ratio, oh * ratio
                cv.drawImage(
                    reader,
                    img_x + (img_w - dw) / 2,
                    img_y + (IMG_H - dh) / 2,
                    dw, dh, mask="auto",
                )
            else:
                _placeholder(img_x, img_y, img_w, IMG_H)

            # Text fields
            ty = img_y - 2.0 - FS
            img_name = _Path(img_path).name if img_path else ""
            for field_name in tile_fields:
                val = img_name if field_name == "Image Name" else record.get(field_name, "").strip()
                display = val if val else "\u2014"
                lbl = f"{field_name}: "
                lbl_w = cv.stringWidth(lbl, FONT_B, FS)
                max_v_w = CELL_W - 2 * PAD - lbl_w - 2
                disp = display
                while disp and cv.stringWidth(disp, FONT_R, FS) > max_v_w:
                    disp = disp[:-1]
                if len(disp) < len(display):
                    disp = disp.rstrip() + "\u2026"
                cv.setFont(FONT_B, FS)
                cv.setFillColor(C_DARK)
                cv.drawString(x + PAD, ty, lbl)
                cv.setFont(FONT_R, FS)
                cv.setFillColor(C_GREY)
                cv.drawString(x + PAD + lbl_w, ty, disp)
                ty -= LINE_H

            col += 1
            if col >= NCOLS:
                col = 0
                y_top -= CELL_H + ROW_GAP

        cv.save()

    def _export_excel(self) -> None:
        from datetime import datetime
        from pathlib import Path as _Path

        ts = datetime.now().strftime("%Y_%m_%d_%H_%M")
        default_name = f"search_result_{ts}"
        start_dir = ""
        if self._use_root_folders:
            root = self._cfg.root_folder()
            if root and _Path(root).is_dir():
                start_dir = root
        save_hint = str(_Path(start_dir) / default_name) if start_dir else default_name

        path, _ = QFileDialog.getSaveFileName(
            self, "Export Search Results", save_hint, "Excel Files (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        minimal  = self._radio_minimal.isChecked()
        cols     = self.MINIMAL_COLS if minimal else self.ALL_COLS
        rows     = self._result_data.get("rows", [])
        groups   = self._result_data.get("groups")
        source   = self._result_data.get("source", "")
        is_rsd   = source == "RSD (master)"
        cols     = [c for c in cols if not (is_rsd and c == "Pack Size")]

        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Search Results"

            header_fill   = PatternFill(fill_type="solid", fgColor="111F35")
            header_font   = Font(color="FFFFFF", bold=True, name="Segoe UI", size=11)
            alt_fill      = PatternFill(fill_type="solid", fgColor="F3F7FC")
            body_font     = Font(name="Segoe UI", size=10)
            thin          = Side(style="thin", color="D9E1EA")
            border        = Border(left=thin, right=thin, top=thin, bottom=thin)
            center        = Alignment(horizontal="center", vertical="center")
            left_align    = Alignment(horizontal="left",   vertical="center")

            if groups:
                # Grouped export: add a leading "Search Value" column
                all_rows: list[dict] = []
                for grp in groups:
                    qv = grp.get("query_value", "")
                    for rec in grp.get("rows", []):
                        all_rows.append({"Search Value": qv, **rec})
                export_cols = ["Search Value"] + list(cols)
                export_rows = all_rows
            else:
                export_cols = list(cols)
                export_rows = rows

            # Header row
            for ci, col_name in enumerate(export_cols, start=1):
                cell = ws.cell(row=1, column=ci, value=col_name)
                cell.fill   = header_fill
                cell.font   = header_font
                cell.alignment = center
                cell.border = border

            # Data rows
            for ri, record in enumerate(export_rows, start=2):
                for ci, col_name in enumerate(export_cols, start=1):
                    cell = ws.cell(row=ri, column=ci, value=record.get(col_name, ""))
                    cell.font   = body_font
                    cell.alignment = left_align
                    cell.border = border
                    if ri % 2 == 0:
                        cell.fill = alt_fill

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = (
                f"A1:{ws.cell(row=1, column=len(export_cols)).coordinate}"
            )
            ws.row_dimensions[1].height = 22
            for ci, col_name in enumerate(export_cols, start=1):
                max_len = len(col_name)
                for rec in export_rows:
                    max_len = max(max_len, len(str(rec.get(col_name, ""))))
                ws.column_dimensions[
                    ws.cell(row=1, column=ci).column_letter
                ].width = min(max(12, max_len + 2), 50)

            wb.save(path)
            QMessageBox.information(self, "Export Complete", f"Saved to:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", str(exc))


class ArtworkCropCanvas(QLabel):
    """Preview canvas with a draggable crop rectangle in PDF-point space."""

    cropChanged = Signal(object)

    def __init__(self, image: QImage, page_bounds: object, initial_bounds: object, initial_path: tuple = (), parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._image = QPixmap.fromImage(image)
        self._page_bounds = page_bounds
        self._crop = initial_bounds
        self._drag_start: tuple[float, float] | None = None
        self._drag_current: tuple[float, float] | None = None
        self._interaction: str | None = None
        self._interaction_start: tuple[float, float] | None = None
        self._interaction_bounds: object | None = None
        self._manual_drawn = False
        self._zoom = 1.0
        self._pan_x = 0.0
        self._pan_y = 0.0
        self._pan_start: tuple[float, float] | None = None
        self._hover_bounds = None
        self._stroke_shape = "rectangle"
        self._corners = "sharp"
        self._corner_amount = 0.0
        self._stroke_weight = 1.0
        self._stroke_color = "#B8F35A"
        self._custom_path = list(initial_path)
        self._custom_segment: dict[str, tuple[float, float]] | None = None
        self._custom_closed = bool(self._custom_path and self._custom_path[-1][0] == "Z")
        self._custom_hover_point: tuple[float, float] | None = None
        self._custom_selected_anchor: int | None = None
        self._custom_hover_segment: tuple[int, float] | None = None
        self._base_canvas_width = 700
        self._base_canvas_height = 460
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumSize(self._base_canvas_width, self._base_canvas_height)
        self.resize(self._base_canvas_width, self._base_canvas_height)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self._refresh()

    def _image_rect(self):
        target = self.size()
        return self._image.scaled(target, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)

    def set_zoom(self, value: int) -> None:
        self._zoom = min(3.0, max(0.25, value / 100.0))
        content_w = max(self._base_canvas_width, round(self._base_canvas_width * self._zoom))
        content_h = max(self._base_canvas_height, round(self._base_canvas_height * self._zoom))
        self.setMinimumSize(content_w, content_h)
        self.resize(content_w, content_h)
        self._refresh()

    def wheelEvent(self, event) -> None:
        current = round(self._zoom * 100)
        step = 10 if event.angleDelta().y() > 0 else -10
        self.set_zoom(current + step)
        event.accept()

    def set_crop_style(self, color: str) -> None:
        self._crop_color = color
        self._refresh()

    def set_stroke_style(self, weight: float, color: str) -> None:
        self._stroke_weight = max(0.1, min(1.0, weight))
        self._stroke_color = color
        self._crop_color = color
        self._refresh()

    def set_hover_bounds(self, bounds: object | None) -> None:
        self._hover_bounds = bounds
        self._refresh()

    def set_shape_settings(self, stroke_shape: str, corners: str, corner_amount: float) -> None:
        self._stroke_shape = stroke_shape
        self._corners = corners
        self._corner_amount = max(0.0, min(1.0, corner_amount))
        if stroke_shape == "custom" and self._custom_closed:
            self.setCursor(Qt.CursorShape.ArrowCursor)
        elif stroke_shape == "custom":
            self._set_pen_cursor()
        else:
            self.unsetCursor()
        self._refresh()

    def _set_pen_cursor(self) -> None:
        """Use a small pen-shaped cursor while drawing a custom path."""
        cursor_pixmap = QPixmap(24, 24)
        cursor_pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(cursor_pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor("#111F35"), 1.5))
        painter.setBrush(QColor("#D02752"))
        pen_body = QPainterPath()
        pen_body.moveTo(5, 17)
        pen_body.lineTo(16, 6)
        pen_body.lineTo(19, 9)
        pen_body.lineTo(8, 20)
        pen_body.closeSubpath()
        painter.drawPath(pen_body)
        painter.setBrush(QColor("#F2F2F2"))
        pen_tip = QPainterPath()
        pen_tip.moveTo(5, 17)
        pen_tip.lineTo(8, 20)
        pen_tip.lineTo(3, 22)
        pen_tip.closeSubpath()
        painter.drawPath(pen_tip)
        painter.setPen(QPen(QColor("#FFFFFF"), 1))
        painter.drawLine(9, 13, 13, 17)
        painter.end()
        self.setCursor(QCursor(cursor_pixmap, 3, 21))

    def set_custom_path(self, commands: tuple = ()) -> None:
        self._custom_path = list(commands)
        self._custom_closed = bool(self._custom_path and self._custom_path[-1][0] == "Z")
        self._custom_segment = None
        self._custom_selected_anchor = None
        self._custom_hover_segment = None
        if self._stroke_shape == "custom" and self._custom_closed:
            self.setCursor(Qt.CursorShape.ArrowCursor)
        self._refresh()

    def _custom_path_bounds(self) -> object | None:
        points = [point for _, values in self._custom_path for point in zip(values[::2], values[1::2])]
        if not points:
            return None
        xs, ys = zip(*points)
        return Bounds(min(xs), min(ys), max(xs), max(ys))

    def _custom_qpath(self, preview: dict | None = None) -> QPainterPath:
        path = QPainterPath()
        def display_point(point: tuple[float, float]) -> tuple[float, float]:
            left, top, width, height = self._display_rect()
            return (
                left + (point[0] - self._page_bounds.left) / max(self._page_bounds.width, 1) * width,
                top + (point[1] - self._page_bounds.top) / max(self._page_bounds.height, 1) * height,
            )
        for command, values in self._custom_path:
            if command == "M":
                path.moveTo(*display_point((values[0], values[1])))
            elif command == "L":
                path.lineTo(*display_point((values[0], values[1])))
            elif command == "Q":
                control = display_point((values[0], values[1]))
                end = display_point((values[2], values[3]))
                path.quadTo(control[0], control[1], end[0], end[1])
            elif command == "Z":
                path.closeSubpath()
        if preview is not None:
            start = preview["start"]
            end = preview["end"]
            control = preview["control"]
            display_control = display_point(control)
            display_end = display_point(end)
            path.quadTo(display_control[0], display_control[1], display_end[0], display_end[1])
        return path

    def _custom_anchors(self) -> list[tuple[float, float]]:
        anchors: list[tuple[float, float]] = []
        for command, values in self._custom_path:
            if command in {"M", "L"}:
                anchors.append((values[0], values[1]))
            elif command == "Q":
                anchors.append((values[2], values[3]))
        return anchors

    def _custom_anchor_path_indices(self) -> list[int]:
        return [index for index, (command, _) in enumerate(self._custom_path) if command in {"M", "L", "Q"}]

    def _custom_anchor_hit(self, x: float, y: float) -> int | None:
        for anchor_index, anchor in enumerate(self._custom_anchors()):
            display_x, display_y = self._to_display_point(anchor)
            if (x - display_x) ** 2 + (y - display_y) ** 2 <= 12 ** 2:
                return anchor_index
        return None

    def _custom_control_hit(self, x: float, y: float) -> int | None:
        for path_index, (command, values) in enumerate(self._custom_path):
            if command != "Q":
                continue
            display_x, display_y = self._to_display_point((values[0], values[1]))
            if (x - display_x) ** 2 + (y - display_y) ** 2 <= 10 ** 2:
                return path_index
        return None

    def _custom_segment_hit(self, x: float, y: float) -> tuple[int, float] | None:
        """Return the nearest path command and parameter when close to a segment."""
        if len(self._custom_anchors()) < 2:
            return None
        best: tuple[float, int, float] | None = None
        previous: tuple[float, float] | None = None
        for path_index, (command, values) in enumerate(self._custom_path):
            if command == "M":
                previous = (values[0], values[1])
                continue
            if command == "Z":
                anchors = self._custom_anchors()
                if previous is None or not anchors:
                    continue
                end = anchors[0]
                points = [
                    (previous[0] + (end[0] - previous[0]) * t, previous[1] + (end[1] - previous[1]) * t)
                    for t in [index / 20 for index in range(21)]
                ]
                for sample_index, point in enumerate(points):
                    display_x, display_y = self._to_display_point(point)
                    distance = (x - display_x) ** 2 + (y - display_y) ** 2
                    if best is None or distance < best[0]:
                        best = (distance, path_index, sample_index / 20)
                continue
            if previous is None:
                continue
            if command == "Q":
                control = (values[0], values[1])
                end = (values[2], values[3])
                points = [
                    ((1 - t) ** 2 * previous[0] + 2 * (1 - t) * t * control[0] + t ** 2 * end[0],
                     (1 - t) ** 2 * previous[1] + 2 * (1 - t) * t * control[1] + t ** 2 * end[1])
                    for t in [index / 20 for index in range(21)]
                ]
            else:
                end = (values[0], values[1])
                points = [
                    (previous[0] + (end[0] - previous[0]) * t, previous[1] + (end[1] - previous[1]) * t)
                    for t in [index / 20 for index in range(21)]
                ]
            for sample_index, point in enumerate(points):
                display_x, display_y = self._to_display_point(point)
                distance = (x - display_x) ** 2 + (y - display_y) ** 2
                if best is None or distance < best[0]:
                    best = (distance, path_index, sample_index / 20)
            previous = end
        if best is not None and best[0] <= 14 ** 2:
            return best[1], best[2]
        return None

    def _move_custom_anchor(self, anchor_index: int, point: tuple[float, float]) -> None:
        path_indices = self._custom_anchor_path_indices()
        if anchor_index >= len(path_indices):
            return
        path_index = path_indices[anchor_index]
        command, values = self._custom_path[path_index]
        old = (values[-2], values[-1])
        delta_x = point[0] - old[0]
        delta_y = point[1] - old[1]
        if command == "Q":
            self._custom_path[path_index] = (
                command,
                (values[0] + delta_x, values[1] + delta_y, point[0], point[1]),
            )
        else:
            self._custom_path[path_index] = (command, (point[0], point[1]))

    def _move_custom_control(self, path_index: int, point: tuple[float, float]) -> None:
        command, values = self._custom_path[path_index]
        if command == "Q":
            self._custom_path[path_index] = (command, (point[0], point[1], values[2], values[3]))

    def _set_anchor_kind(self, anchor_index: int, kind: str) -> None:
        path_indices = self._custom_anchor_path_indices()
        if anchor_index >= len(path_indices):
            return
        path_index = path_indices[anchor_index]
        command, values = self._custom_path[path_index]
        if command == "M":
            return
        if kind == "corner" and command == "Q":
            self._custom_path[path_index] = ("L", (values[2], values[3]))
        elif kind == "curve" and command != "Q":
            previous_index = path_index - 1
            while previous_index >= 0 and self._custom_path[previous_index][0] not in {"M", "L", "Q"}:
                previous_index -= 1
            previous_values = self._custom_path[previous_index][1] if previous_index >= 0 else values
            previous = (previous_values[-2], previous_values[-1])
            end = (values[0], values[1])
            midpoint = ((previous[0] + end[0]) / 2, (previous[1] + end[1]) / 2)
            self._custom_path[path_index] = ("Q", (midpoint[0], midpoint[1], end[0], end[1]))
        self._refresh()
        self.cropChanged.emit(self._custom_path_bounds())

    def _insert_custom_anchor(self, path_index: int, fraction: float) -> None:
        command, values = self._custom_path[path_index]
        if command == "Z":
            anchors = self._custom_anchors()
            if not anchors:
                return
            previous_values = self._custom_path[path_index - 1][1]
            start = (previous_values[-2], previous_values[-1])
            end = anchors[0]
            middle = (start[0] + (end[0] - start[0]) * fraction, start[1] + (end[1] - start[1]) * fraction)
            self._custom_path[path_index:path_index + 1] = [
                ("L", (middle[0], middle[1])),
                ("L", (end[0], end[1])),
                ("Z", ()),
            ]
            self._custom_selected_anchor = len(self._custom_anchors()) - 1
            return
        previous_index = path_index - 1
        while previous_index >= 0 and self._custom_path[previous_index][0] not in {"M", "L", "Q"}:
            previous_index -= 1
        if previous_index < 0:
            return
        previous_values = self._custom_path[previous_index][1]
        start = (previous_values[-2], previous_values[-1])
        if command == "Q":
            control = (values[0], values[1])
            end = (values[2], values[3])
            first_control = (start[0] + (control[0] - start[0]) * fraction, start[1] + (control[1] - start[1]) * fraction)
            second_control = (control[0] + (end[0] - control[0]) * fraction, control[1] + (end[1] - control[1]) * fraction)
            middle = (first_control[0] + (second_control[0] - first_control[0]) * fraction, first_control[1] + (second_control[1] - first_control[1]) * fraction)
            self._custom_path[path_index:path_index + 1] = [
                ("Q", (first_control[0], first_control[1], middle[0], middle[1])),
                ("Q", (second_control[0], second_control[1], end[0], end[1])),
            ]
        else:
            end = (values[0], values[1])
            middle = (start[0] + (end[0] - start[0]) * fraction, start[1] + (end[1] - start[1]) * fraction)
            self._custom_path[path_index:path_index + 1] = [
                ("L", (middle[0], middle[1])),
                ("L", (end[0], end[1])),
            ]
        self._custom_selected_anchor = len(self._custom_anchors()) - 2

    def _to_display_point(self, point: tuple[float, float]) -> tuple[float, float]:
        left, top, width, height = self._display_rect()
        return (
            left + (point[0] - self._page_bounds.left) / max(self._page_bounds.width, 1) * width,
            top + (point[1] - self._page_bounds.top) / max(self._page_bounds.height, 1) * height,
        )

    def _custom_first_anchor_hit(self, x: float, y: float) -> bool:
        anchors = self._custom_anchors()
        if len(anchors) < 3:
            return False
        first_x, first_y = self._to_display_point(anchors[0])
        return (x - first_x) ** 2 + (y - first_y) ** 2 <= 12 ** 2

    def _finish_custom_path(self, close: bool = False) -> None:
        if self._custom_segment is not None:
            return
        if close and len(self._custom_anchors()) >= 3 and not self._custom_closed:
            self._custom_path.append(("Z", ()))
            self._custom_closed = True
        self._interaction = None
        self._custom_hover_point = None
        self.unsetCursor()
        self._refresh()
        self.cropChanged.emit(self._custom_path_bounds())

    def _undo_custom(self) -> None:
        if self._custom_segment is not None:
            self._custom_segment = None
        elif self._custom_path:
            if self._custom_path[-1][0] == "Z":
                self._custom_path.pop()
                self._custom_closed = False
            elif len(self._custom_path) > 1:
                self._custom_path.pop()
            else:
                self._custom_path.clear()
        self._refresh()
        self.cropChanged.emit(self._custom_path_bounds())

    def _custom_point(self, point: tuple[float, float], modifiers: Qt.KeyboardModifier) -> tuple[float, float]:
        if not self._custom_path or not modifiers & Qt.KeyboardModifier.ShiftModifier:
            return point
        last_values = self._custom_path[-1][1]
        last = (last_values[-2], last_values[-1])
        dx, dy = point[0] - last[0], point[1] - last[1]
        return (last[0] + dx, last[1]) if abs(dx) >= abs(dy) else (last[0], last[1] + dy)

    def _display_rect(self):
        scaled = self._image_rect()
        left = (self.width() - scaled.width()) / 2 + self._pan_x
        top = (self.height() - scaled.height()) / 2 + self._pan_y
        return left, top, scaled.width(), scaled.height()

    def _rect_hit_test(self, x: float, y: float) -> str | None:
        if self._crop is None:
            return None
        rect_x, rect_y, rect_w, rect_h = self._to_display_rect(self._crop)
        pad = 10
        if rect_w <= 0 or rect_h <= 0:
            if rect_x <= x <= rect_x + rect_w and rect_y <= y <= rect_y + rect_h:
                return "move"
            return None
        if x <= rect_x + pad and y <= rect_y + pad:
            return "resize_tl"
        if x >= rect_x + rect_w - pad and y <= rect_y + pad:
            return "resize_tr"
        if x <= rect_x + pad and y >= rect_y + rect_h - pad:
            return "resize_bl"
        if x >= rect_x + rect_w - pad and y >= rect_y + rect_h - pad:
            return "resize_br"
        if x <= rect_x + pad:
            return "resize_left"
        if x >= rect_x + rect_w - pad:
            return "resize_right"
        if y <= rect_y + pad:
            return "resize_top"
        if y >= rect_y + rect_h - pad:
            return "resize_bottom"
        if rect_x <= x <= rect_x + rect_w and rect_y <= y <= rect_y + rect_h:
            return "move"
        return None

    def _clamp_page_point(self, x: float, y: float) -> tuple[float, float]:
        return (
            max(self._page_bounds.left, min(self._page_bounds.right, x)),
            max(self._page_bounds.top, min(self._page_bounds.bottom, y)),
        )

    def _apply_interaction(self, end_point: tuple[float, float]) -> None:
        if self._interaction is None or self._interaction_start is None or self._interaction_bounds is None:
            return
        start_x, start_y = self._interaction_start
        end_x, end_y = end_point
        base = self._interaction_bounds
        dx = end_x - start_x
        dy = end_y - start_y
        if self._interaction == "move":
            left = base.left + dx
            top = base.top + dy
            right = base.right + dx
            bottom = base.bottom + dy
            width = base.width
            height = base.height
            if left < self._page_bounds.left:
                delta = self._page_bounds.left - left
                left += delta
                right += delta
            if top < self._page_bounds.top:
                delta = self._page_bounds.top - top
                top += delta
                bottom += delta
            if right > self._page_bounds.right:
                delta = right - self._page_bounds.right
                left -= delta
                right -= delta
            if bottom > self._page_bounds.bottom:
                delta = bottom - self._page_bounds.bottom
                top -= delta
                bottom -= delta
            self._crop = Bounds(left, top, right, bottom)
            return

        if self._interaction in {"resize_left", "resize_tl", "resize_bl"}:
            new_left = min(base.right - 1.0, max(self._page_bounds.left, base.left + dx))
            base = Bounds(new_left, base.top, base.right, base.bottom)
        if self._interaction in {"resize_right", "resize_tr", "resize_br"}:
            new_right = max(base.left + 1.0, min(self._page_bounds.right, base.right + dx))
            base = Bounds(base.left, base.top, new_right, base.bottom)
        if self._interaction in {"resize_top", "resize_tl", "resize_tr"}:
            new_top = min(base.bottom - 1.0, max(self._page_bounds.top, base.top + dy))
            base = Bounds(base.left, new_top, base.right, base.bottom)
        if self._interaction in {"resize_bottom", "resize_bl", "resize_br"}:
            new_bottom = max(base.top + 1.0, min(self._page_bounds.bottom, base.bottom + dy))
            base = Bounds(base.left, base.top, base.right, new_bottom)
        self._crop = base

    def _cursor_for_interaction(self, interaction: str | None) -> Qt.CursorShape:
        mapping = {
            "move": Qt.CursorShape.SizeAllCursor,
            "resize_left": Qt.CursorShape.SizeHorCursor,
            "resize_right": Qt.CursorShape.SizeHorCursor,
            "resize_top": Qt.CursorShape.SizeVerCursor,
            "resize_bottom": Qt.CursorShape.SizeVerCursor,
            "resize_tl": Qt.CursorShape.SizeFDiagCursor,
            "resize_tr": Qt.CursorShape.SizeBDiagCursor,
            "resize_bl": Qt.CursorShape.SizeBDiagCursor,
            "resize_br": Qt.CursorShape.SizeFDiagCursor,
        }
        return mapping.get(interaction, Qt.CursorShape.CrossCursor)

    def _to_page_point(self, x: float, y: float) -> tuple[float, float]:
        left, top, width, height = self._display_rect()
        px = self._page_bounds.left + ((x - left) / max(width, 1)) * self._page_bounds.width
        py = self._page_bounds.top + ((y - top) / max(height, 1)) * self._page_bounds.height
        return (
            max(self._page_bounds.left, min(self._page_bounds.right, px)),
            max(self._page_bounds.top, min(self._page_bounds.bottom, py)),
        )

    def _to_display_rect(self, bounds: object) -> tuple[float, float, float, float]:
        left, top, width, height = self._display_rect()
        x = left + (bounds.left - self._page_bounds.left) / max(self._page_bounds.width, 1) * width
        y = top + (bounds.top - self._page_bounds.top) / max(self._page_bounds.height, 1) * height
        w = bounds.width / max(self._page_bounds.width, 1) * width
        h = bounds.height / max(self._page_bounds.height, 1) * height
        return x, y, w, h

    def _refresh(self) -> None:
        pixmap = self._image_rect()
        canvas = QPixmap(self.size())
        canvas.fill(QColor("#F2F2F2"))
        painter = QPainter(canvas)
        painter.drawPixmap((self.width() - pixmap.width()) // 2, (self.height() - pixmap.height()) // 2, pixmap)
        x, y, width, height = self._to_display_rect(self._crop)
        stroke_color = getattr(self, "_crop_color", self._stroke_color)
        stroke_pen = QPen(QColor(stroke_color))
        stroke_pen.setWidthF(4.0 * self._stroke_weight)
        painter.setPen(stroke_pen)
        rect = QRect(round(x), round(y), max(1, round(width)), max(1, round(height)))
        if self._stroke_shape == "custom" and self._custom_path:
            painter.drawPath(self._custom_qpath({
                "start": self._custom_segment["start"],
                "end": self._custom_segment["end"],
                "control": self._custom_segment["control"],
            }) if self._custom_segment else self._custom_qpath())
            if self._custom_segment is not None:
                start_x, start_y = self._to_display_point(self._custom_segment["start"])
                control_x, control_y = self._to_display_point(self._custom_segment["control"])
                painter.setPen(QPen(QColor("#FFB000"), 2, Qt.PenStyle.DashLine))
                painter.drawLine(round(start_x), round(start_y), round(control_x), round(control_y))
                painter.setBrush(QColor("#FFFFFF"))
                painter.setPen(QPen(QColor("#FFB000"), 2))
                painter.drawEllipse(round(control_x - 5), round(control_y - 5), 10, 10)
            for index, anchor in enumerate(self._custom_anchors()):
                ax, ay = self._to_display_point(anchor)
                size = 8 if index == 0 else 6
                selected = index == self._custom_selected_anchor
                painter.setBrush(QColor("#D02752") if selected else QColor("#FFFFFF"))
                painter.setPen(QPen(QColor("#111F35") if selected else QColor(stroke_color), 2))
                if index == 0:
                    painter.drawRect(round(ax - size / 2), round(ay - size / 2), size, size)
                else:
                    painter.drawEllipse(round(ax - size / 2), round(ay - size / 2), size, size)
            if self._custom_closed:
                previous_anchor: tuple[float, float] | None = None
                for command, values in self._custom_path:
                    if command == "M":
                        previous_anchor = (values[0], values[1])
                    elif command == "Q" and previous_anchor is not None:
                        control = (values[0], values[1])
                        control_x, control_y = self._to_display_point(control)
                        start_x, start_y = self._to_display_point(previous_anchor)
                        painter.setPen(QPen(QColor("#FFB000"), 2, Qt.PenStyle.DashLine))
                        painter.drawLine(round(start_x), round(start_y), round(control_x), round(control_y))
                        painter.setBrush(QColor("#FFFFFF"))
                        painter.setPen(QPen(QColor("#FFB000"), 2))
                        painter.drawEllipse(round(control_x - 5), round(control_y - 5), 10, 10)
                        previous_anchor = (values[2], values[3])
                    elif command == "L":
                        previous_anchor = (values[0], values[1])
        elif self._stroke_shape == "ellipse":
            painter.drawEllipse(rect)
        elif self._corners == "rounded":
            radius = min(width, height) * (0.5 * self._corner_amount)
            painter.drawRoundedRect(rect, radius, radius)
        elif self._corners == "beveled" or self._stroke_shape == "custom":
            path = QPainterPath()
            bevel = min(width, height) * (0.25 * self._corner_amount)
            path.moveTo(x + bevel, y)
            path.lineTo(x + width - bevel, y)
            path.lineTo(x + width, y + bevel)
            path.lineTo(x + width, y + height - bevel)
            path.lineTo(x + width - bevel, y + height)
            path.lineTo(x + bevel, y + height)
            path.lineTo(x, y + height - bevel)
            path.lineTo(x, y + bevel)
            path.closeSubpath()
            painter.drawPath(path)
        else:
            painter.drawRect(rect)
        if self._hover_bounds is not None:
            hover_x, hover_y, hover_width, hover_height = self._to_display_rect(self._hover_bounds)
            painter.setPen(QPen(QColor("#FFFF00"), 3, Qt.PenStyle.DashLine))
            painter.drawRect(
                round(hover_x), round(hover_y),
                max(1, round(hover_width)), max(1, round(hover_height)),
            )
        painter.end()
        self.setPixmap(canvas)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._refresh()

    def mousePressEvent(self, event) -> None:
        self.setFocus()
        if event.button() == Qt.MouseButton.MiddleButton:
            self._pan_start = (event.position().x(), event.position().y())
            return
        point = self._to_page_point(event.position().x(), event.position().y())
        if self._stroke_shape == "custom":
            if self._custom_closed:
                if event.button() == Qt.MouseButton.RightButton:
                    anchor_index = self._custom_anchor_hit(event.position().x(), event.position().y())
                    if anchor_index is not None:
                        menu = QMenu(self)
                        curve_action = menu.addAction("Curve")
                        corner_action = menu.addAction("Corner")
                        chosen = menu.exec(self.mapToGlobal(event.position().toPoint()))
                        if chosen == curve_action:
                            self._set_anchor_kind(anchor_index, "curve")
                        elif chosen == corner_action:
                            self._set_anchor_kind(anchor_index, "corner")
                        return
                    return
                control_index = self._custom_control_hit(event.position().x(), event.position().y())
                if control_index is not None:
                    self._custom_selected_anchor = None
                    self._interaction = "custom_control"
                    self._interaction_start = point
                    self._interaction_bounds = control_index
                    self.setCursor(Qt.CursorShape.SizeAllCursor)
                    return
                anchor_index = self._custom_anchor_hit(event.position().x(), event.position().y())
                if anchor_index is not None:
                    self._custom_selected_anchor = anchor_index
                    self._interaction = "custom_anchor"
                    self._interaction_start = point
                    self._interaction_bounds = anchor_index
                    self.setCursor(Qt.CursorShape.SizeAllCursor)
                    self._refresh()
                    return
                segment = self._custom_segment_hit(event.position().x(), event.position().y())
                if segment is not None:
                    self._insert_custom_anchor(*segment)
                    self._manual_drawn = True
                    self._refresh()
                    self.cropChanged.emit(self._custom_path_bounds())
                    return
                self._custom_selected_anchor = None
                self._refresh()
                return
            point = self._custom_point(point, event.modifiers())
            if self._custom_first_anchor_hit(event.position().x(), event.position().y()):
                self._finish_custom_path(close=True)
                return
            if not self._custom_path:
                self._custom_path.append(("M", point))
                self._custom_segment = None
            else:
                last_values = self._custom_path[-1][1]
                # The click fixes the next anchor. Dragging moves its tangent
                # handle, matching the pen-tool interaction used by Illustrator.
                self._custom_segment = {
                    "start": (last_values[-2], last_values[-1]),
                    "end": point,
                    "control": point,
                }
            self._manual_drawn = True
            self._interaction = "custom"
            self._set_pen_cursor()
            self._refresh()
            return
        hit = self._rect_hit_test(event.position().x(), event.position().y()) if self._crop is not None else None
        if hit is not None:
            self._interaction = hit
            self._interaction_start = point
            self._interaction_bounds = self._crop
            self._manual_drawn = True
            self.setCursor(self._cursor_for_interaction(hit))
            return
        self._drag_start = point
        self._drag_current = point
        self._interaction = "draw"
        self._manual_drawn = True
        self._interaction_start = point
        self._interaction_bounds = None
        self.setCursor(Qt.CursorShape.CrossCursor)

    def mouseMoveEvent(self, event) -> None:
        if self._pan_start is not None:
            self._pan_x += event.position().x() - self._pan_start[0]
            self._pan_y += event.position().y() - self._pan_start[1]
            self._pan_start = (event.position().x(), event.position().y())
            self._refresh()
            return
        if self._interaction == "custom" and self._custom_segment is not None:
            control = self._to_page_point(event.position().x(), event.position().y())
            control = self._custom_point(control, event.modifiers())
            self._custom_segment["control"] = control
            self._refresh()
            self.cropChanged.emit(self._custom_path_bounds())
            return
        if self._stroke_shape == "custom" and self._custom_closed:
            point = self._clamp_page_point(
                *self._to_page_point(event.position().x(), event.position().y())
            )
            if self._interaction == "custom_anchor" and self._interaction_bounds is not None:
                anchor_index = self._interaction_bounds
                self._move_custom_anchor(anchor_index, point)
                self._refresh()
                self.cropChanged.emit(self._custom_path_bounds())
                return
            if self._interaction == "custom_control" and self._interaction_bounds is not None:
                self._move_custom_control(self._interaction_bounds, point)
                self._refresh()
                self.cropChanged.emit(self._custom_path_bounds())
                return
            if self._custom_anchor_hit(event.position().x(), event.position().y()) is not None or self._custom_control_hit(event.position().x(), event.position().y()) is not None:
                self.setCursor(Qt.CursorShape.SizeAllCursor)
            else:
                self.setCursor(Qt.CursorShape.ArrowCursor)
            return
        if self._stroke_shape == "custom" and self._custom_path and self._custom_segment is None:
            self._custom_hover_point = self._to_page_point(event.position().x(), event.position().y())
            self._refresh()
            self._set_pen_cursor()
            return
        if self._interaction == "draw":
            if self._drag_start is None:
                return
            self._drag_current = self._to_page_point(event.position().x(), event.position().y())
            self._set_drag_crop()
            return
        if self._interaction is not None and self._interaction_start is not None:
            point = self._clamp_page_point(*self._to_page_point(event.position().x(), event.position().y()))
            self._apply_interaction(point)
            self._refresh()
            self.cropChanged.emit(self._crop)
            return
        hit = self._rect_hit_test(event.position().x(), event.position().y()) if self._crop is not None else None
        self.setCursor(self._cursor_for_interaction(hit))

    def mouseReleaseEvent(self, event) -> None:
        if self._pan_start is not None:
            self._pan_start = None
            return
        if self._interaction == "custom" and self._custom_segment is not None:
            segment = self._custom_segment
            control = segment["control"]
            end = segment["end"]
            if abs(control[0] - segment["start"][0]) + abs(control[1] - segment["start"][1]) < 2:
                self._custom_path.append(("L", end))
            else:
                self._custom_path.append(("Q", (control[0], control[1], end[0], end[1])))
            self._custom_segment = None
            self._interaction = None
            self._set_pen_cursor()
            self._refresh()
            self.cropChanged.emit(self._custom_path_bounds())
            return
        if self._interaction in {"custom_anchor", "custom_control"}:
            self._interaction = None
            self._interaction_start = None
            self._interaction_bounds = None
            self.setCursor(Qt.CursorShape.ArrowCursor)
            self._refresh()
            self.cropChanged.emit(self._custom_path_bounds())
            return
        if self._interaction == "draw":
            if self._drag_start is not None:
                self._drag_current = self._to_page_point(event.position().x(), event.position().y())
                self._set_drag_crop()
            self._drag_start = None
            self._drag_current = None
            self._interaction = None
            self._interaction_start = None
            self._interaction_bounds = None
            self.unsetCursor()
            return
        if self._interaction is not None:
            self._interaction = None
            self._interaction_start = None
            self._interaction_bounds = None
            self._refresh()
            self.cropChanged.emit(self._crop)
            self.unsetCursor()

    def _set_drag_crop(self) -> None:
        if self._drag_start is None or self._drag_current is None:
            return
        x1, y1 = self._drag_start
        x2, y2 = self._drag_current
        self._crop = Bounds(min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
        self._refresh()
        self.cropChanged.emit(self._crop)

    def crop(self) -> object:
        return self._custom_path_bounds() if self._stroke_shape == "custom" and self._custom_path else self._crop

    def custom_path(self) -> tuple:
        return tuple(self._custom_path)

    def mouseDoubleClickEvent(self, event) -> None:
        if self._stroke_shape == "custom" and event.button() == Qt.MouseButton.LeftButton:
            self._finish_custom_path(close=False)
            return
        super().mouseDoubleClickEvent(event)

    def keyPressEvent(self, event) -> None:
        if self._stroke_shape == "custom":
            if event.key() in (Qt.Key.Key_Backspace, Qt.Key.Key_Delete):
                self._undo_custom()
                return
            if event.key() == Qt.Key.Key_Escape:
                self._finish_custom_path(close=False)
                return
        super().keyPressEvent(event)


class ArtworkCropDialog(QDialog):
    """Large single-artwork crop editor."""

    def __init__(self, inspection: ArtworkInspection, parent: QWidget | None = None, initial_bounds: object | None = None, saved_selection: CutSelection | None = None) -> None:
        super().__init__(parent)
        self.setWindowFlag(Qt.WindowType.WindowMaximizeButtonHint, True)
        self.setWindowTitle(f"Manual Cut - {inspection.input_path.name}")
        self.resize(980, 700)
        self.setMinimumSize(780, 560)
        self._inspection = inspection
        self._mode = "manual"
        self._selected_source: CutSelection | None = None
        self._manual_hide_selection: CutSelection | None = None
        self._selection_actions: dict[tuple[str, str], str] = {}
        self._action_combos: dict[tuple[str, str], QComboBox] = {}
        self._action_options: dict[tuple[str, str], CutOption] = {}
        self._initial_bounds = initial_bounds
        self._saved_selection = saved_selection
        self._hover_option: CutOption | None = None
        self._hover_enabled = False
        self._stroke_weight = 1.0
        self._stroke_color = "#B8F35A"
        self._preserved_hide_items: tuple[tuple[str, str, tuple[int, int, int] | None, int | None], ...] = ()
        self._build_ui()

    def _add_action_row(self, layout: QVBoxLayout, option: CutOption, mode: str) -> None:
        row = QHBoxLayout()
        row.setSpacing(8)
        label = QLabel(option.label)
        label.setObjectName("cropOptionCheck")
        label.installEventFilter(self)
        label.setProperty("crop_option", option)
        row.addWidget(label)
        row.addStretch(1)
        combo = QComboBox()
        combo.setObjectName("cropOptionCheck")
        combo.addItems(["show", "hide", "trim", "trim & hide"])
        combo.setCurrentText("show")
        combo.setFixedWidth(115)
        self._set_action_combo_color(combo, "show")
        row.addWidget(combo)
        key = (mode, option.label)
        self._selection_actions[key] = "show"
        self._action_combos[key] = combo
        self._action_options[key] = option

        def handle_change(action: str) -> None:
            self._selection_actions[key] = action
            self._set_action_combo_color(combo, action)
            if action == "show":
                self.canvas.set_crop_style("#B8F35A")
                self._selected_source = None
                self._manual_hide_selection = None
                self.canvas._refresh()
                return
            if action in {"trim", "trim & hide"}:
                if self._manual_crop_checkbox.isChecked():
                    self._manual_crop_checkbox.setChecked(False)
                self.canvas.set_crop_style("#00FF66")
                self.canvas._crop = option.bounds
                self._selected_source = CutSelection(option.bounds, mode, option.label, option.color_rgb, option.xref)
                self._manual_hide_selection = None if action == "trim" else CutSelection(option.bounds, mode, option.label, option.color_rgb, option.xref)
                for other_key, other_action in list(self._selection_actions.items()):
                    if other_key != key and other_action in {"trim", "trim & hide"}:
                        self._selection_actions[other_key] = "show"
                        other_combo = self._action_combos.get(other_key)
                        if other_combo is not None:
                            other_combo.blockSignals(True)
                            other_combo.setCurrentText("show")
                            other_combo.blockSignals(False)
                            self._set_action_combo_color(other_combo, "show")
            elif action == "hide":
                self.canvas.set_crop_style("#00FF66")
                self._selected_source = None
                self._manual_hide_selection = CutSelection(self.canvas.crop(), mode, option.label, option.color_rgb, option.xref)
            self.canvas._refresh()

        combo.currentTextChanged.connect(handle_change)
        layout.addLayout(row)

    @staticmethod
    def _set_action_combo_color(combo: QComboBox, action: str) -> None:
        colors = {
            "show": "#7CD5C7",
            "hide": "#D3D3D3",
            "trim": "#BD5579",
            "trim & hide": "#BD5579",
        }
        combo.setStyleSheet(
            f"QComboBox {{ background-color: {colors.get(action, '#FFFFFF')}; padding: 3px 8px; }}"
        )

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)
        title = QLabel(self._inspection.input_path.name)
        title.setObjectName("manualCutTitle")
        layout.addWidget(title)
        hint = QLabel("Select a crop method on the left. Use the mouse wheel to zoom and middle-drag to pan.")
        hint.setObjectName("manualCutDescription")
        layout.addWidget(hint)

        initial = self._initial_bounds if self._initial_bounds is not None else self._inspection.detected_bounds
        body = QHBoxLayout()
        crop_panel = QFrame()
        crop_panel.setObjectName("cropControlPanel")
        crop_panel.setFixedWidth(350)
        crop_layout = QVBoxLayout(crop_panel)
        crop_title = QLabel("Local Crop Setup")
        crop_title.setObjectName("manualCutSectionLabel")
        crop_layout.addWidget(crop_title)
        self._manual_crop_checkbox = QCheckBox("manual crop")
        self._manual_crop_checkbox.setObjectName("cropOptionCheck")
        self._manual_crop_checkbox.setStyleSheet(
            "QCheckBox#cropOptionCheck::indicator:checked "
            "{ background-color: #D02752; border: 1px solid #D02752; }"
        )
        self._manual_crop_checkbox.toggled.connect(self._manual_crop_changed)
        crop_layout.addWidget(self._manual_crop_checkbox)
        self._hover_checkbox = QCheckBox("show crop border on hover")
        self._hover_checkbox.setObjectName("cropOptionCheck")
        self._hover_checkbox.setStyleSheet(
            "QCheckBox#cropOptionCheck::indicator:checked "
            "{ background-color: #D02752; border: 1px solid #D02752; }"
        )
        self._hover_checkbox.toggled.connect(self._set_hover_enabled)
        self._hover_checkbox.setChecked(True)
        crop_layout.addWidget(self._hover_checkbox)
        tabs = QTabWidget()
        tabs.setObjectName("cropTabs")
        layers_tab = QWidget(); layers_layout = QVBoxLayout(layers_tab)
        layer_options = [option for option in self._inspection.options if option.kind == "layer"]
        if not layer_options: layers_layout.addWidget(QLabel("No PDF layers found."))
        for option in layer_options: self._add_action_row(layers_layout, option, "layer")
        layers_layout.addStretch(1)
        layers_scroll = QScrollArea()
        layers_scroll.setWidgetResizable(True)
        layers_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        layers_scroll.setWidget(layers_tab)
        separations_tab = QWidget(); separations_layout = QVBoxLayout(separations_tab)
        separation_options = [option for option in self._inspection.options if option.kind == "spot_color"]
        if not separation_options: separations_layout.addWidget(QLabel("No vector separations found."))
        for option in separation_options: self._add_action_row(separations_layout, option, "spot_color")
        separations_layout.addStretch(1)
        separations_scroll = QScrollArea()
        separations_scroll.setWidgetResizable(True)
        separations_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        separations_scroll.setWidget(separations_tab)
        tabs.addTab(layers_scroll, "By Layer")
        tabs.addTab(separations_scroll, "By Spot Color")
        shape_tab = QWidget()
        shape_layout = QVBoxLayout(shape_tab)
        shape_layout.setSpacing(8)
        stroke_shape_row = QHBoxLayout()
        stroke_shape_row.setSpacing(8)
        stroke_shape_label = QLabel("stroke shape")
        stroke_shape_label.setObjectName("cropOptionCheck")
        stroke_shape_row.addWidget(stroke_shape_label)
        stroke_shape_row.addStretch(1)
        self._stroke_shape_combo = QComboBox()
        self._stroke_shape_combo.addItems(["rectangle", "ellipse", "custom"])
        self._stroke_shape_combo.setCurrentText("rectangle")
        self._stroke_shape_combo.setObjectName("cropOptionCheck")
        self._stroke_shape_combo.setFixedWidth(115)
        self._set_action_combo_color(self._stroke_shape_combo, "hide")
        self._stroke_shape_combo.currentTextChanged.connect(self._shape_settings_changed)
        stroke_shape_row.addWidget(self._stroke_shape_combo)
        shape_layout.addLayout(stroke_shape_row)
        self._custom_help_label = QLabel(
            "Custom: click anchors, click-drag for curves, click the first anchor to close."
        )
        self._custom_help_label.setObjectName("manualCutDescription")
        self._custom_help_label.setWordWrap(True)
        shape_layout.addWidget(self._custom_help_label)
        stroke_weight_row = QHBoxLayout()
        stroke_weight_row.setSpacing(8)
        stroke_weight_label = QLabel("stroke weight")
        stroke_weight_label.setObjectName("cropOptionCheck")
        stroke_weight_row.addWidget(stroke_weight_label)
        stroke_weight_row.addStretch(1)
        self._stroke_weight_value = QLabel("1.00")
        self._stroke_weight_value.setObjectName("cropOptionCheck")
        stroke_weight_row.addWidget(self._stroke_weight_value)
        shape_layout.addLayout(stroke_weight_row)
        self._stroke_weight_slider = QSlider(Qt.Orientation.Horizontal)
        self._stroke_weight_slider.setRange(1, 10)
        self._stroke_weight_slider.setValue(10)
        self._stroke_weight_slider.valueChanged.connect(self._shape_settings_changed)
        shape_layout.addWidget(self._stroke_weight_slider)
        stroke_color_row = QHBoxLayout()
        stroke_color_row.setSpacing(8)
        stroke_color_label = QLabel("stroke color")
        stroke_color_label.setObjectName("cropOptionCheck")
        stroke_color_row.addWidget(stroke_color_label)
        stroke_color_row.addStretch(1)
        self._stroke_color_button = QPushButton()
        self._stroke_color_button.setFixedSize(115, 28)
        self._stroke_color_button.setToolTip("Choose stroke color")
        self._stroke_color_button.clicked.connect(self._choose_stroke_color)
        stroke_color_row.addWidget(self._stroke_color_button)
        shape_layout.addLayout(stroke_color_row)
        self._set_stroke_color_button()
        corners_row = QHBoxLayout()
        corners_row.setSpacing(8)
        self._corners_label = QLabel("corners")
        self._corners_label.setObjectName("cropOptionCheck")
        corners_row.addWidget(self._corners_label)
        corners_row.addStretch(1)
        self._corners_combo = QComboBox()
        self._corners_combo.addItems(["sharp", "rounded", "beveled"])
        self._corners_combo.setCurrentText("sharp")
        self._corners_combo.setObjectName("cropOptionCheck")
        self._corners_combo.setFixedWidth(115)
        self._set_action_combo_color(self._corners_combo, "hide")
        self._corners_combo.currentTextChanged.connect(self._shape_settings_changed)
        corners_row.addWidget(self._corners_combo)
        shape_layout.addLayout(corners_row)
        self._corner_slider_label = QLabel("0.00")
        shape_layout.addWidget(self._corner_slider_label)
        self._corner_slider = QSlider(Qt.Orientation.Horizontal)
        self._corner_slider.setRange(0, 100)
        self._corner_slider.setValue(0)
        self._corner_slider.valueChanged.connect(self._shape_settings_changed)
        shape_layout.addWidget(self._corner_slider)
        shape_layout.addStretch(1)
        tabs.addTab(shape_tab, "Shape Settings")
        self._update_corner_controls()
        crop_layout.addWidget(tabs, 1)
        body.addWidget(crop_panel)
        self.canvas = ArtworkCropCanvas(self._inspection_image(), self._inspection.page_bounds, initial)
        self.canvas.cropChanged.connect(self._on_canvas_crop_changed)
        canvas_scroll = QScrollArea()
        canvas_scroll.setWidgetResizable(True)
        canvas_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        canvas_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        canvas_scroll.setWidget(self.canvas)
        body.addWidget(canvas_scroll, 1)
        layout.addLayout(body, 1)
        self._restore_selection_state()

        zoom_row = QHBoxLayout(); zoom_label = QLabel("Zoom"); zoom_label.setObjectName("manualCutDescription"); zoom_row.addWidget(zoom_label)
        zoom_slider = QSlider(Qt.Orientation.Horizontal); zoom_slider.setObjectName("manualCutZoomSlider"); zoom_slider.setRange(25, 300); zoom_slider.setValue(100); zoom_slider.setTickInterval(25); zoom_slider.valueChanged.connect(self.canvas.set_zoom); zoom_row.addWidget(zoom_slider, 1)
        self.zoom_value = QLabel("100%"); self.zoom_value.setObjectName("manualCutDescription"); zoom_slider.valueChanged.connect(lambda value: self.zoom_value.setText(f"{value}%")); zoom_row.addWidget(self.zoom_value)
        layout.addLayout(zoom_row)

        footer = QHBoxLayout(); footer.addStretch(1)
        accept = QPushButton("Use this crop"); accept.setObjectName("collectorRunBtn"); accept.setFixedHeight(50); accept.setMinimumWidth(180); accept.clicked.connect(self.accept); footer.addWidget(accept); layout.addLayout(footer)

    def _manual_crop_changed(self, checked: bool) -> None:
        if not checked:
            return
        for key, action in list(self._selection_actions.items()):
            if action in {"trim", "trim & hide"}:
                self._selection_actions[key] = "show"
                combo = self._action_combos.get(key)
                if combo is not None:
                    combo.blockSignals(True)
                    combo.setCurrentText("show")
                    combo.blockSignals(False)
                    self._set_action_combo_color(combo, "show")
        self._selected_source = None
        self._manual_hide_selection = None
        self.canvas.set_crop_style("#00FF66")
        self.canvas._refresh()

    def _on_canvas_crop_changed(self, bounds: object) -> None:
        if getattr(self.canvas, "_manual_drawn", False) and not self._manual_crop_checkbox.isChecked():
            preserved = list(getattr(self._saved_selection, "hide_items", ()) or ())
            for key, option in self._action_options.items():
                if self._selection_actions.get(key) in {"hide", "trim & hide"}:
                    item = (option.kind, option.label, option.color_rgb, option.xref)
                    if item not in preserved:
                        preserved.append(item)
            self._preserved_hide_items = tuple(preserved)
            self._manual_crop_checkbox.setChecked(True)
        self.canvas._manual_drawn = False

    def _restore_selection_state(self) -> None:
        if self._saved_selection is None:
            return
        selection = self._saved_selection
        self._stroke_shape = selection.stroke_shape
        self._corners = selection.corners
        self._corner_amount = selection.corner_amount
        self._stroke_weight = selection.stroke_weight
        self._stroke_color = selection.stroke_color
        self._stroke_shape_combo.setCurrentText(self._stroke_shape)
        self._corners_combo.setCurrentText(self._corners)
        self._corner_slider.setValue(round(self._corner_amount * 100))
        self._stroke_weight_slider.setValue(round(self._stroke_weight * 10))
        self._set_stroke_color_button()
        self._shape_settings_changed()
        self.canvas.set_custom_path(selection.custom_path)
        if selection.manual_crop:
            self._manual_crop_checkbox.blockSignals(True)
            self._manual_crop_checkbox.setChecked(True)
            self._manual_crop_checkbox.blockSignals(False)
            self.canvas._crop = selection.bounds
            self.canvas._refresh()
            return
        source_key = None
        if selection.kind in {"layer", "spot_color"} and selection.name:
            source_key = (selection.kind, selection.name)
            action = "trim & hide" if any(item[:2] == source_key for item in selection.hide_items) else "trim"
            self._selection_actions[source_key] = action
            combo = self._action_combos.get(source_key)
            if combo is not None:
                combo.blockSignals(True)
                combo.setCurrentText(action)
                combo.blockSignals(False)
                self._set_action_combo_color(combo, action)
        for kind, name, _, _ in selection.hide_items:
            key = (kind, name)
            if key == source_key:
                continue
            self._selection_actions[key] = "hide"
            combo = self._action_combos.get(key)
            if combo is not None:
                combo.blockSignals(True)
                combo.setCurrentText("hide")
                combo.blockSignals(False)
                self._set_action_combo_color(combo, "hide")
        self.canvas._crop = selection.bounds
        self.canvas._refresh()

    def _set_hover_enabled(self, enabled: bool) -> None:
        self._hover_enabled = enabled
        if not hasattr(self, "canvas"):
            return
        if not enabled:
            self._hover_option = None
        self.canvas.set_hover_bounds(self._hover_option.bounds if self._hover_option else None)

    def _update_corner_controls(self) -> None:
        visible = self._stroke_shape_combo.currentText() == "rectangle" and self._corners_combo.currentText() in {"rounded", "beveled"}
        self._corners_label.setVisible(self._stroke_shape_combo.currentText() == "rectangle")
        self._corners_combo.setVisible(self._stroke_shape_combo.currentText() == "rectangle")
        self._corner_slider_label.setVisible(visible)
        self._corner_slider.setVisible(visible)
        self._custom_help_label.setVisible(self._stroke_shape_combo.currentText() == "custom")

    def _shape_settings_changed(self, _value: object = None) -> None:
        self._stroke_shape = self._stroke_shape_combo.currentText()
        self._corners = self._corners_combo.currentText()
        self._corner_amount = self._corner_slider.value() / 100.0
        self._stroke_weight = self._stroke_weight_slider.value() / 10.0
        self._corner_slider_label.setText(f"{self._corner_amount:.2f}")
        self._stroke_weight_value.setText(f"{self._stroke_weight:.2f}")
        self._update_corner_controls()
        self.canvas.set_shape_settings(self._stroke_shape, self._corners, self._corner_amount)
        self.canvas.set_stroke_style(self._stroke_weight, self._stroke_color)

    def _set_stroke_color_button(self) -> None:
        self._stroke_color_button.setStyleSheet(
            f"QPushButton {{ background-color: {self._stroke_color}; "
            "border: 1px solid #A8A8A8; border-radius: 4px; }}"
            "QPushButton:hover { border: 2px solid #555555; }"
        )

    def _choose_stroke_color(self) -> None:
        color = QColorDialog.getColor(QColor(self._stroke_color), self, "Choose stroke color")
        if not color.isValid():
            return
        self._stroke_color = color.name().upper()
        self._set_stroke_color_button()
        self.canvas.set_stroke_style(self._stroke_weight, self._stroke_color)

    def _set_source_crop(self, option: object, checked: bool, mode: str) -> None:
        if not checked:
            return
        self._mode = mode
        self.canvas._crop = option.bounds
        self._selected_source = CutSelection(
            option.bounds,
            "layer" if mode == "layer" else "spot_color",
            option.label,
            option.color_rgb,
            option.xref,
        )
        self.canvas.set_crop_style("#00FF66")
        self.canvas._refresh()

    def _set_manual_hide(self, option: object, checked: bool, mode: str) -> None:
        if not checked:
            return
        self._manual_hide_selection = CutSelection(
            self.canvas.crop(), mode, option.label, option.color_rgb, option.xref,
        )
        self.canvas.set_crop_style("#00FF66")
        self.canvas._refresh()

    def _tab_changed(self, index: int) -> None:
        self._mode = "manual" if index == 2 else "source"
        if index == 2:
            self.canvas.set_crop_style("#00FF66")
            self._selected_source = None
        else:
            self.canvas.set_crop_style("#00FF66")

    def eventFilter(self, obj, event) -> bool:
        if isinstance(obj, QLabel) and obj.property("crop_option") is not None:
            option = obj.property("crop_option")
            if event.type() == QEvent.Type.Enter:
                self._hover_option = option
                if self._hover_enabled:
                    self.canvas.set_hover_bounds(option.bounds)
            elif event.type() == QEvent.Type.Leave:
                self._hover_option = None
                self.canvas.set_hover_bounds(None)
        return super().eventFilter(obj, event)

    def _inspection_image(self) -> QImage:
        raw = self._inspection.preview
        return QImage(raw.samples, raw.width, raw.height, raw.stride, QImage.Format.Format_RGB888).copy()

    def _set_bounds(self, bounds: object) -> None:
        self.canvas._crop = bounds
        self.canvas._refresh()

    def _spin_changed(self) -> None:
        return

    def crop_bounds(self) -> object:
        return self.canvas.crop()

    def crop_selection(self) -> CutSelection:
        shape_settings = {
            "stroke_shape": self._stroke_shape,
            "corners": self._corners,
            "corner_amount": self._corner_amount,
            "custom_path": self.canvas.custom_path(),
            "stroke_weight": self._stroke_weight,
            "stroke_color": self._stroke_color,
        }
        hide_items = tuple(
            (option.kind, option.label, option.color_rgb, option.xref)
            for key, option in self._action_options.items()
            if self._selection_actions.get(key) in {"hide", "trim & hide"}
        )
        hide_items = tuple(dict.fromkeys(hide_items + self._preserved_hide_items))
        if self._manual_crop_checkbox.isChecked():
            return CutSelection(
                self.canvas.crop(),
                "manual",
                hide_items=hide_items,
                manual_crop=True,
                **shape_settings,
            )
        trim_option = next(
            (self._action_options[key] for key, action in self._selection_actions.items()
             if action in {"trim", "trim & hide"}),
            None,
        )
        if trim_option is not None:
            return CutSelection(
                trim_option.bounds,
                trim_option.kind,
                trim_option.label,
                trim_option.color_rgb,
                trim_option.xref,
                hide_items=hide_items,
                **shape_settings,
            )
        if hide_items:
            return CutSelection(self.canvas.crop(), "manual", hide_items=hide_items, **shape_settings)
        if self._manual_hide_selection is not None:
            hide = self._manual_hide_selection
            return CutSelection(
                self.canvas.crop(), "manual", "", None, None,
                hide.kind, hide.name, hide.color_rgb, hide.xref,
                hide_items=((hide.kind, hide.name, hide.color_rgb, hide.xref),),
                **shape_settings,
            )
        if self._selected_source is not None:
            return CutSelection(
                self.canvas.crop(),
                self._selected_source.kind,
                self._selected_source.name,
                self._selected_source.color_rgb,
                self._selected_source.xref,
                hide_items=(),
                **shape_settings,
            )
        return CutSelection(self._inspection.page_bounds, "manual", **shape_settings)


class ArtworkManualCutDialog(QDialog):
    """Manual cut dialog with individual and bulk review modes."""

    def __init__(self, inspections: list[ArtworkInspection], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowFlag(Qt.WindowType.WindowMaximizeButtonHint, True)
        self.setWindowTitle("Manual Cut - Artwork Preview")
        self.resize(1120, 760)
        self.setMinimumSize(860, 560)
        self._inspections = inspections
        self._selections: dict[Path, object | None] = {}
        self._preview_labels: dict[Path, QLabel] = {}
        self._copied_crop: CutSelection | None = None
        self._selected_paths: set[Path] = set()
        self._selection_anchor: Path | None = None
        self._bulk_mode = True
        self._bulk_actions: dict[tuple[str, str], str] = {}
        self._bulk_combos: dict[tuple[str, str], QComboBox] = {}
        self._bulk_rows: dict[tuple[str, str], dict[str, object]] = {}
        self._card_widgets: dict[Path, QFrame] = {}
        self._global_hover_enabled = False
        self._global_hover_option: tuple[str, str] | None = None
        self.output_folder: Path | None = None
        self._build_ui()

    @staticmethod
    def _set_action_combo_color(combo: QComboBox, action: str) -> None:
        colors = {
            "show": "#7CD5C7",
            "hide": "#D3D3D3",
            "trim": "#BD5579",
            "trim & hide": "#BD5579",
        }
        combo.setStyleSheet(
            f"QComboBox {{ background-color: {colors.get(action, '#FFFFFF')}; }}"
        )

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        heading = QLabel("Manual Cut")
        heading.setObjectName("manualCutTitle")
        layout.addWidget(heading)

        self._bulk_panel = QWidget()
        self._bulk_panel.setVisible(True)
        self._bulk_panel.setFixedWidth(350)
        self._bulk_panel.setObjectName("cropControlPanel")
        self._bulk_layout = QVBoxLayout(self._bulk_panel)
        self._bulk_layout.setContentsMargins(12, 12, 12, 12)
        crop_title = QLabel("Global Crop Setup")
        crop_title.setObjectName("manualCutSectionLabel")
        self._bulk_layout.addWidget(crop_title)
        self._global_hover_checkbox = QCheckBox("show crop border on hover")
        self._global_hover_checkbox.setObjectName("cropOptionCheck")
        self._global_hover_checkbox.setStyleSheet(
            "QCheckBox#cropOptionCheck::indicator:checked "
            "{ background-color: #D02752; border: 1px solid #D02752; }"
        )
        self._global_hover_checkbox.toggled.connect(self._set_global_hover_enabled)
        self._global_hover_checkbox.setChecked(True)
        self._bulk_layout.addWidget(self._global_hover_checkbox)
        self._bulk_tabs = QTabWidget()
        self._bulk_tabs.setObjectName("cropTabs")
        self._bulk_tabs.addTab(self._bulk_layer_tab(), "By Layer")
        self._bulk_tabs.addTab(self._bulk_spot_tab(), "By Spot Color")
        self._bulk_layout.addWidget(self._bulk_tabs, 1)

        intro = QLabel("Click an artwork thumbnail to open its manual crop editor.")
        intro.setObjectName("manualCutDescription")
        layout.addWidget(intro)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setObjectName("manualCutScroll")
        self._scroll.viewport().installEventFilter(self)
        self._cards = QWidget()
        self._cards_layout = QGridLayout(self._cards)
        self._cards_layout.setContentsMargins(2, 2, 2, 2)
        self._cards_layout.setSpacing(14)
        for index, inspection in enumerate(self._inspections):
            self._cards_layout.addWidget(self._build_card(inspection), index // 3, index % 3)
        for column in range(3):
            self._cards_layout.setColumnStretch(column, 1)
        self._scroll.setWidget(self._cards)
        content_row = QHBoxLayout()
        content_row.setSpacing(12)
        content_row.addWidget(self._bulk_panel, 0, Qt.AlignmentFlag.AlignTop)
        content_row.addWidget(self._scroll, 1)
        layout.addLayout(content_row, 1)

        footer = QHBoxLayout()
        footer.addStretch(1)
        apply_button = QPushButton("Use selected cuts")
        apply_button.setObjectName("collectorRunBtn")
        apply_button.setFixedHeight(50)
        apply_button.setMinimumWidth(180)
        apply_button.clicked.connect(self._choose_output_folder_and_accept)
        footer.addWidget(apply_button)
        layout.addLayout(footer)

    def _choose_output_folder_and_accept(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Choose folder to save artwork files")
        if not folder:
            return
        self.output_folder = Path(folder)
        self.accept()

    def _build_bulk_catalog(self, kind: str) -> list[str]:
        names: list[str] = []
        seen: set[str] = set()
        for inspection in self._inspections:
            for option in inspection.options:
                if option.kind != kind:
                    continue
                label = option.label.strip()
                if label and label not in seen:
                    seen.add(label)
                    names.append(label)
        return names

    def _bulk_layer_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        names = self._build_bulk_catalog("layer")
        if not names:
            layout.addWidget(QLabel("No PDF layers found."))
            return tab
        for name in names:
            row = QHBoxLayout()
            row.setSpacing(8)
            label = QLabel(name)
            label.setObjectName("cropOptionCheck")
            label.installEventFilter(self)
            label.setProperty("global_crop_option", ("layer", name))
            row.addWidget(label)
            row.addStretch(1)
            combo = QComboBox()
            combo.setObjectName("cropOptionCheck")
            combo.addItems(["show", "hide", "trim", "trim & hide"])
            combo.setCurrentText("show")
            combo.setFixedWidth(115)
            self._set_action_combo_color(combo, "show")
            row.addWidget(combo)
            key = ("layer", name)
            self._bulk_actions[key] = "show"
            self._bulk_combos[key] = combo
            combo.currentTextChanged.connect(
                lambda action, target=name: self._on_bulk_combo_change("layer", target, action)
            )
            layout.addLayout(row)
        layout.addStretch(1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(tab)
        return scroll

    def _bulk_spot_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        names = self._build_bulk_catalog("spot_color")
        if not names:
            layout.addWidget(QLabel("No vector separations found."))
            return tab
        for name in names:
            row = QHBoxLayout()
            row.setSpacing(8)
            label = QLabel(name)
            label.setObjectName("cropOptionCheck")
            label.installEventFilter(self)
            label.setProperty("global_crop_option", ("spot_color", name))
            row.addWidget(label)
            row.addStretch(1)
            combo = QComboBox()
            combo.setObjectName("cropOptionCheck")
            combo.addItems(["show", "hide", "trim", "trim & hide"])
            combo.setCurrentText("show")
            combo.setFixedWidth(115)
            self._set_action_combo_color(combo, "show")
            row.addWidget(combo)
            key = ("spot_color", name)
            self._bulk_actions[key] = "show"
            self._bulk_combos[key] = combo
            combo.currentTextChanged.connect(
                lambda action, target=name: self._on_bulk_combo_change("spot_color", target, action)
            )
            layout.addLayout(row)
        layout.addStretch(1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(tab)
        return scroll

    def _on_bulk_combo_change(self, kind: str, name: str, action: str) -> None:
        key = (kind, name)
        self._bulk_actions[key] = action
        combo = self._bulk_combos.get(key)
        if combo is not None:
            self._set_action_combo_color(combo, action)
        if action in {"trim", "trim & hide"}:
            for other_key, other_action in list(self._bulk_actions.items()):
                if other_key != key and other_action in {"trim", "trim & hide"}:
                    self._bulk_actions[other_key] = "show"
                    other_combo = self._bulk_combos.get(other_key)
                    if other_combo is not None:
                        other_combo.blockSignals(True)
                        other_combo.setCurrentText("show")
                        other_combo.blockSignals(False)
                        self._set_action_combo_color(other_combo, "show")
        self._refresh_bulk_previews()

    def _set_global_hover_enabled(self, enabled: bool) -> None:
        self._global_hover_enabled = enabled
        if not enabled:
            self._global_hover_option = None
        self._refresh_bulk_previews()

    def _set_global_hover_option(self, option: tuple[str, str] | None) -> None:
        self._global_hover_option = option if self._global_hover_enabled else None
        self._refresh_bulk_previews()

    def _refresh_bulk_previews(self) -> None:
        trim_choice = next(
            ((kind, name) for (kind, name), action in self._bulk_actions.items()
             if action in {"trim", "trim & hide"}),
            None,
        )
        for inspection in self._inspections:
            local_selection = self._selections.get(inspection.input_path)
            if local_selection is not None:
                local_bounds = local_selection.bounds
                self._update_preview(inspection, local_bounds, selection=local_selection)
                continue
            bounds = None
            hover_bounds = None
            if trim_choice is not None:
                trim_kind, trim_name = trim_choice
                option = next(
                    (candidate for candidate in inspection.options
                     if candidate.kind == trim_kind and candidate.label == trim_name),
                    None,
                )
                if option is not None:
                    bounds = option.bounds
            if self._global_hover_option is not None:
                hover_kind, hover_name = self._global_hover_option
                hover_option = next(
                    (candidate for candidate in inspection.options
                     if candidate.kind == hover_kind and candidate.label == hover_name),
                    None,
                )
                if hover_option is not None:
                    hover_bounds = hover_option.bounds
            self._update_preview(inspection, bounds, hover_bounds)

    def _build_card(self, inspection: ArtworkInspection) -> QFrame:
        card = QFrame()
        card.setObjectName("manualCutCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 14, 14, 14)

        preview = QLabel()
        preview.setObjectName("manualCutPreview")
        preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview.setFixedSize(300, 210)
        preview.setCursor(Qt.CursorShape.PointingHandCursor)
        preview.installEventFilter(self)
        self._preview_labels[inspection.input_path] = preview
        self._card_widgets[inspection.input_path] = card
        reset_btn = QPushButton()
        reset_btn.setObjectName("manualCutResetBtn")
        reset_btn.setToolTip("Reset crop")
        reset_btn.setFixedSize(34, 30)
        reset_btn.setStyleSheet(
            "QPushButton#manualCutResetBtn { background: #F2F4F8; color: #444; border: 1px solid #C8CDD6; border-radius: 8px; font-size: 18px; font-weight: 700; }"
            "QPushButton#manualCutResetBtn:hover { background: #E9EDF2; }"
        )
        reset_btn.setText("↺")
        reset_btn.clicked.connect(lambda checked=False, path=inspection.input_path: self._clear_saved_selection(path))

        menu_btn = QToolButton()
        menu_btn.setObjectName("manualCutMenuBtn")
        menu_btn.setToolTip("Crop actions")
        menu_btn.setFixedSize(34, 30)
        menu_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        menu_btn.setArrowType(Qt.ArrowType.NoArrow)
        menu_btn.setStyleSheet(
            "QToolButton#manualCutMenuBtn { background: #F2F4F8; color: #2D2D2D; border: 1px solid #C8CDD6; border-radius: 9px; font-size: 20px; font-weight: 800; padding: 0; }"
            "QToolButton#manualCutMenuBtn:hover { background: #E9EDF2; border-color: #B9C1CC; }"
            "QToolButton#manualCutMenuBtn:pressed { background: #E2E8F0; }"
            "QToolButton#manualCutMenuBtn::menu-indicator { image: none; }"
        )
        menu_btn.setText("⋯")
        menu = QMenu(menu_btn)
        menu.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        menu.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        menu.setStyleSheet(
            "QMenu { background: rgba(255,255,255,0.97); border: 1px solid #D9DEE5; border-radius: 10px; padding: 6px; }"
            "QMenu::item { padding: 8px 18px; color: #1F2937; font-size: 12px; border-radius: 6px; }"
            "QMenu::item:selected { background: #EEF5FF; color: #111827; }"
        )
        copy_action = menu.addAction("Copy")
        paste_action = menu.addAction("Paste")
        menu_btn.setMenu(menu)
        menu.triggered.connect(lambda action, path=inspection.input_path: self._handle_card_menu_action(path, action, copy_action, paste_action))

        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.addStretch(1)
        top_row.addWidget(reset_btn)
        top_row.addWidget(menu_btn)
        card_layout.addLayout(top_row)
        card_layout.addWidget(preview, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self._selections[inspection.input_path] = None
        self._update_preview(inspection, None)
        self._refresh_saved_tag(inspection.input_path)
        return card

    def _handle_card_menu_action(self, path: Path, action, copy_action, paste_action) -> None:
        if action == copy_action:
            self._copied_crop = self._selections.get(path)
            return
        if action == paste_action:
            source = self._copied_crop
            if source is None:
                return
            for inspection in self._inspections:
                if inspection.input_path == path:
                    self._selections[path] = source
                    self._refresh_saved_tag(path)
                    preview_bounds = source.bounds if source.kind in {"layer", "spot_color"} or source.manual_crop else None
                    self._update_preview(inspection, preview_bounds, selection=source)
                    break

    def _clear_saved_selection(self, path: Path) -> None:
        self._selections[path] = None
        self._refresh_saved_tag(path)
        inspection = next(item for item in self._inspections if item.input_path == path)
        trim_choice = next(
            ((kind, name) for (kind, name), action in self._bulk_actions.items()
             if action in {"trim", "trim & hide"}),
            None,
        )
        if trim_choice is not None:
            trim_kind, trim_name = trim_choice
            trim_option = next(
                (option for option in inspection.options
                 if option.kind == trim_kind and option.label == trim_name),
                None,
            )
            if trim_option is not None:
                self._update_preview(inspection, trim_option.bounds)
                return
        self._update_preview(inspection, None)

    def _refresh_saved_tag(self, path: Path) -> None:
        selection = self._selections.get(path)
        preview = self._preview_labels.get(path)
        has_saved = selection is not None
        if preview is not None:
            if path in self._selected_paths:
                preview.setStyleSheet(
                    "QLabel#manualCutPreview { background-color: rgba(124, 213, 199, 0.2); border: 2px solid #7CD5C7; border-radius: 8px; }"
                )
            else:
                preview.setStyleSheet(
                    "QLabel#manualCutPreview { border: 2px solid #D02752; border-radius: 8px; }"
                    if has_saved else ""
                )

    def _apply_selection(self, selected: set[Path]) -> None:
        self._selected_paths = selected
        self._selection_anchor = next(iter(selected), None) if selected else None
        for path, preview in self._preview_labels.items():
            if path in selected:
                preview.setStyleSheet("QLabel#manualCutPreview { background-color: rgba(124, 213, 199, 0.2); border: 2px solid #7CD5C7; border-radius: 8px; }")
            elif self._selections.get(path) is not None:
                preview.setStyleSheet("QLabel#manualCutPreview { border: 2px solid #D02752; border-radius: 8px; }")
            else:
                preview.setStyleSheet("")

    def eventFilter(self, obj, event) -> bool:
        if isinstance(obj, QLabel) and obj.property("global_crop_option") is not None:
            option = obj.property("global_crop_option")
            if event.type() == QEvent.Type.Enter:
                self._global_hover_option = option
                if self._global_hover_enabled:
                    self._refresh_bulk_previews()
            elif event.type() == QEvent.Type.Leave:
                self._global_hover_option = None
                self._refresh_bulk_previews()
        if event.type() == QEvent.Type.MouseButtonDblClick:
            if obj in self._preview_labels.values():
                path = next(path for path, label in self._preview_labels.items() if label is obj)
                inspection = next(item for item in self._inspections if item.input_path == path)
                local_selection = self._selections.get(path)
                initial_bounds = local_selection.bounds if local_selection is not None else None
                if self._bulk_mode and initial_bounds is None:
                    trim_choice = next(
                        ((kind, name) for (kind, name), action in self._bulk_actions.items()
                         if action in {"trim", "trim & hide"}),
                        None,
                    )
                    if trim_choice is not None:
                        trim_kind, trim_name = trim_choice
                        trim_option = next(
                            (option for option in inspection.options
                             if option.kind == trim_kind and option.label == trim_name),
                            None,
                        )
                        if trim_option is not None:
                            initial_bounds = trim_option.bounds
                editor = ArtworkCropDialog(inspection, self, initial_bounds, local_selection)
                if editor.exec() == QDialog.DialogCode.Accepted:
                    selection = editor.crop_selection()
                    self._selections[path] = selection
                    self._selected_paths = {path}
                    self._selection_anchor = path
                    self._apply_selection({path})
                    trim_bounds = (
                        selection.bounds
                        if selection.kind in {"layer", "spot_color"} or selection.manual_crop
                        else None
                    )
                    self._update_preview(inspection, trim_bounds, selection=selection)
                return False

        if event.type() == QEvent.Type.MouseButtonPress:
            if obj in self._preview_labels.values():
                path = next(path for path, label in self._preview_labels.items() if label is obj)
                mod = QApplication.keyboardModifiers()
                ctrl = bool(mod & Qt.KeyboardModifier.ControlModifier)
                shift = bool(mod & Qt.KeyboardModifier.ShiftModifier)
                if ctrl:
                    next_selection = set(self._selected_paths)
                    if path in next_selection:
                        next_selection.remove(path)
                    else:
                        next_selection.add(path)
                    self._selection_anchor = path
                    self._apply_selection(next_selection)
                    return False
                if shift:
                    if self._selection_anchor is None:
                        self._selection_anchor = path
                    anchor_index = next((idx for idx, inspection in enumerate(self._inspections) if inspection.input_path == self._selection_anchor), None)
                    current_index = next((idx for idx, inspection in enumerate(self._inspections) if inspection.input_path == path), None)
                    if anchor_index is not None and current_index is not None:
                        start, end = sorted((anchor_index, current_index))
                        next_selection = {self._inspections[idx].input_path for idx in range(start, end + 1)}
                        self._apply_selection(next_selection)
                    return False
                self._selected_paths = {path}
                self._selection_anchor = path
                self._apply_selection({path})
                return False

            if self._selected_paths and obj not in self._preview_labels.values():
                self._selected_paths.clear()
                self._selection_anchor = None
                self._apply_selection(set())
        return super().eventFilter(obj, event)

    def _update_preview(
        self,
        inspection: ArtworkInspection,
        bounds: object,
        hover_bounds: object = None,
        selection: CutSelection | None = None,
    ) -> None:
        raw = inspection.preview
        image = QImage(raw.samples, raw.width, raw.height, raw.stride, QImage.Format.Format_RGB888).copy()
        pixmap = QPixmap.fromImage(image)
        pixmap = pixmap.scaled(286, 196, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        painter = None
        if bounds is not None or hover_bounds is not None:
            painter = QPainter(pixmap)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            scale_x = pixmap.width() / max(inspection.page_bounds.width, 1)
            scale_y = pixmap.height() / max(inspection.page_bounds.height, 1)
        if bounds is not None:
            stroke_color = selection.stroke_color if selection is not None else "#00FF66"
            stroke_weight = selection.stroke_weight if selection is not None else 1.0
            pen = QPen(QColor(stroke_color))
            pen.setWidthF(4.0 * stroke_weight)
            painter.setPen(pen)
            x = bounds.left * scale_x
            y = bounds.top * scale_y
            width = max(1, bounds.width * scale_x)
            height = max(1, bounds.height * scale_y)
            rect = QRect(round(x), round(y), round(width), round(height))
            stroke_shape = selection.stroke_shape if selection is not None else "rectangle"
            corners = selection.corners if selection is not None else "sharp"
            corner_amount = selection.corner_amount if selection is not None else 0.0
            if stroke_shape == "custom" and selection is not None and selection.custom_path:
                custom_path = QPainterPath()
                for command, values in selection.custom_path:
                    if command == "M":
                        custom_path.moveTo(values[0] * scale_x, values[1] * scale_y)
                    elif command == "L":
                        custom_path.lineTo(values[0] * scale_x, values[1] * scale_y)
                    elif command == "Q":
                        custom_path.quadTo(
                            values[0] * scale_x, values[1] * scale_y,
                            values[2] * scale_x, values[3] * scale_y,
                        )
                    elif command == "Z":
                        custom_path.closeSubpath()
                painter.drawPath(custom_path)
            elif stroke_shape == "ellipse":
                painter.drawEllipse(rect)
            elif corners == "rounded":
                radius = min(width, height) * (0.5 * corner_amount)
                painter.drawRoundedRect(rect, radius, radius)
            elif corners == "beveled" or stroke_shape == "custom":
                path = QPainterPath()
                bevel = min(width, height) * (0.25 * corner_amount)
                path.moveTo(x + bevel, y)
                path.lineTo(x + width - bevel, y)
                path.lineTo(x + width, y + bevel)
                path.lineTo(x + width, y + height - bevel)
                path.lineTo(x + width - bevel, y + height)
                path.lineTo(x + bevel, y + height)
                path.lineTo(x, y + height - bevel)
                path.lineTo(x, y + bevel)
                path.closeSubpath()
                painter.drawPath(path)
            else:
                painter.drawRect(rect)
        if hover_bounds is not None:
            painter.setPen(QPen(QColor("#FFFF00"), 3, Qt.PenStyle.DashLine))
            painter.drawRect(
                round(hover_bounds.left * scale_x), round(hover_bounds.top * scale_y),
                max(1, round(hover_bounds.width * scale_x)),
                max(1, round(hover_bounds.height * scale_y)),
            )
        if painter is not None:
            painter.end()
        self._preview_labels[inspection.input_path].setPixmap(pixmap)

    def selections(self) -> dict[Path, object | None]:
        if self._bulk_mode:
            selections: dict[Path, object | None] = {}
            for inspection in self._inspections:
                local_selection = self._selections.get(inspection.input_path)
                if local_selection is not None:
                    selections[inspection.input_path] = local_selection
                    continue
                trim_option: CutOption | None = None
                hide_items: list[tuple[str, str, tuple[int, int, int] | None, int | None]] = []
                for option in inspection.options:
                    action = self._bulk_actions.get((option.kind, option.label), "")
                    if action in {"trim", "trim & hide"}:
                        trim_option = option
                    if action in {"hide", "trim & hide"}:
                        hide_items.append((option.kind, option.label, option.color_rgb, option.xref))
                if trim_option is not None:
                    selection = CutSelection(
                        bounds=trim_option.bounds,
                        kind=trim_option.kind,
                        name=trim_option.label,
                        color_rgb=trim_option.color_rgb,
                        xref=trim_option.xref,
                        hide_items=tuple(hide_items),
                    )
                elif hide_items:
                    selection = CutSelection(
                        bounds=inspection.page_bounds,
                        kind="manual",
                        hide_items=tuple(hide_items),
                    )
                else:
                    selection = CutSelection(
                        bounds=inspection.page_bounds,
                        kind="manual",
                    )
                selections[inspection.input_path] = selection
            return selections
        return dict(self._selections)


class NewUIWindow(QMainWindow):
    """UI-only dashboard shell. No business logic is connected yet."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("HAT Dashboard")
        self.resize(1047, 792)
        self.setMinimumSize(1047, 792)

        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        shell = QFrame()
        shell.setObjectName("shell")
        outer.addWidget(shell)

        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(12, 12, 12, 12)
        shell_layout.setSpacing(14)

        self.btn_home = self._make_top_button("Home")
        self.btn_master = self._make_top_button("Master")
        self.btn_project = self._make_top_button("Project")
        self.btn_search = self._make_top_button("Search")
        self.btn_artwork = self._make_top_button("Artwork")

        main_row = QHBoxLayout()
        main_row.setSpacing(14)
        shell_layout.addLayout(main_row, 1)

        icon_rail = QVBoxLayout()
        icon_rail.setSpacing(10)
        icon_rail.setAlignment(Qt.AlignTop)
        main_row.addLayout(icon_rail)

        self.btn_settings = self._make_icon_button("\u2699")  # ⚙ cog
        self.btn_settings.setToolTip("Settings")
        icon_rail.addWidget(self.btn_settings)

        self.btn_rail_toggle = ToggleSwitch(width=46, height=26)
        self.btn_rail_toggle.setToolTip("Use configured root folders")
        icon_rail.addWidget(self.btn_rail_toggle)
        icon_rail.addStretch(1)

        content = QFrame()
        content.setObjectName("contentCard")
        main_row.addWidget(content, 1)

        self._content_layout = QVBoxLayout(content)
        self._content_layout.setContentsMargins(20, 18, 20, 18)
        self._content_layout.setSpacing(0)
        content_layout = self._content_layout

        header_band = QFrame()
        header_band.setObjectName("headerBand")
        header_band.setFixedHeight(110)
        content_layout.addWidget(header_band)

        # ── Project selector row (visible only when PROJECT is active) ──────────
        self.project_selector_row = QWidget()
        self.project_selector_row.setVisible(False)
        proj_sel_outer = QHBoxLayout(self.project_selector_row)
        proj_sel_outer.setContentsMargins(0, 0, 0, 0)
        proj_sel_outer.setSpacing(0)

        self._proj_sel_band = QFrame()
        self._proj_sel_band.setObjectName("projectSelectorBand")
        self._proj_sel_band.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        proj_band_layout = QHBoxLayout(self._proj_sel_band)
        proj_band_layout.setContentsMargins(14, 6, 14, 6)
        proj_band_layout.setSpacing(10)

        lbl_proj_name = QLabel("Project Name:")
        lbl_proj_name.setObjectName("projectSelectorLabel")
        self.combo_project = QComboBox()
        self.combo_project.setObjectName("projectSelectorCombo")
        self.combo_project.setMinimumWidth(400)
        self.combo_project.setFixedHeight(28)
        proj_band_layout.addWidget(lbl_proj_name)
        proj_band_layout.addWidget(self.combo_project)
        self.combo_project.currentIndexChanged.connect(self._on_project_combo_changed)
        proj_band_layout.addStretch(1)

        proj_sel_outer.addWidget(self._proj_sel_band, 1)
        # (added to active_tools_page VBox below)

        header_row = QHBoxLayout(header_band)
        header_row.setContentsMargins(16, 14, 16, 14)
        header_row.setSpacing(0)

        left_slot = QWidget()
        left_slot.setFixedWidth(250)
        left_slot_layout = QHBoxLayout(left_slot)
        left_slot_layout.setContentsMargins(16, 0, 0, 0)
        left_slot_layout.setSpacing(0)

        hat_badge = QLabel()
        hat_badge.setObjectName("hatBadge")
        hat_badge.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        hat_badge.setMinimumWidth(231)
        _logo_path = str(_BASE_DIR / "logo" / "hat_logo.png")
        _logo_pix = QPixmap(_logo_path)
        if not _logo_pix.isNull():
            hat_badge.setPixmap(_logo_pix.scaled(
                231, 84, Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation))
        left_slot_layout.addWidget(hat_badge)
        left_slot_layout.addStretch(1)
        header_row.addWidget(left_slot)

        buttons_slot = QWidget()
        buttons_row = QHBoxLayout(buttons_slot)
        buttons_row.setContentsMargins(0, 0, 0, 0)
        buttons_row.setSpacing(8)
        buttons_row.addWidget(self.btn_home)
        buttons_row.addWidget(self.btn_master)
        buttons_row.addWidget(self.btn_project)
        buttons_row.addWidget(self.btn_search)
        buttons_row.addWidget(self.btn_artwork)
        buttons_row.addItem(QSpacerItem(20, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        header_row.addWidget(buttons_slot, 1)

        right_slot = QWidget()
        right_slot.setFixedWidth(40)
        header_row.addWidget(right_slot)

        self.section_title = QLabel("HOME")
        self.section_title.setObjectName("sectionTitle")

        # ── Active tools page (shared by MASTER and PROJECT) ─────────────────────
        self.active_tools_page = QWidget()
        self._active_page_layout = QVBoxLayout(self.active_tools_page)
        self._active_page_layout.setContentsMargins(0, 16, 0, 0)
        self._active_page_layout.setSpacing(8)
        active_page_layout = self._active_page_layout

        # Body row: left nav + right column (selector + right panel)
        active_body_widget = QWidget()
        self._active_body_layout = QHBoxLayout(active_body_widget)
        self._active_body_layout.setContentsMargins(0, 0, 0, 0)
        self._active_body_layout.setSpacing(18)
        active_page_layout.addWidget(active_body_widget, 1)
        active_body_layout = self._active_body_layout

        # Left nav stack – switches between master / project button sets
        self.left_nav_stack = QStackedWidget()
        self.left_nav_stack.setMaximumWidth(178)
        self.left_nav_blank = QWidget()  # HOME / SEARCH
        self.left_nav_stack.addWidget(self.left_nav_blank)  # index 0

        # ── MASTER left buttons: TSC, TG, SDR ────────────────────────────────
        self.master_left = QWidget()
        master_actions_layout = QVBoxLayout(self.master_left)
        master_actions_layout.setContentsMargins(0, 0, 0, 0)
        master_actions_layout.setSpacing(14)

        self.btn_action_status_collector = QPushButton("Trackers Status Collector")
        self.btn_action_thumbnail_generator = QPushButton("Thumbnail Generator")
        self.btn_action_sdr_master = QPushButton("SAP Data Reformat")

        master_action_buttons = (
            self.btn_action_status_collector,
            self.btn_action_thumbnail_generator,
            self.btn_action_sdr_master,
        )
        self.master_action_group = QButtonGroup(self)
        self.master_action_group.setExclusive(True)
        for btn in master_action_buttons:
            btn.setObjectName("basicToolActionBtn")
            btn.setCheckable(True)
            btn.setText(self._wrap_button_text(btn.text(), words_per_line=2))
            btn.setFixedSize(160, 132)
            btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            self.master_action_group.addButton(btn)
        for btn in master_action_buttons:
            master_actions_layout.addWidget(btn, 0, Qt.AlignmentFlag.AlignLeft)
        master_actions_layout.addStretch(1)
        self.left_nav_stack.addWidget(self.master_left)  # index 1

        # ── PROJECT left buttons: PNG, SDR, SDC, Project Review ───────────────────
        self.project_left = QWidget()
        project_actions_layout = QVBoxLayout(self.project_left)
        project_actions_layout.setContentsMargins(0, 10, 0, 0)
        project_actions_layout.setSpacing(14)

        self.btn_action_packshot_naming = QPushButton("Packshot Naming Generator")
        self.btn_action_sap_data_reformat = QPushButton("SAP Data Reformat")
        self.btn_action_sap_data_compare = QPushButton("SAP Data Compare")
        self.btn_action_project_viewer = QPushButton("Review Project")
        self.btn_action_other_tools = QPushButton("Other Tools")

        project_action_buttons = (
            self.btn_action_packshot_naming,
            self.btn_action_sap_data_reformat,
            self.btn_action_sap_data_compare,
            self.btn_action_project_viewer,
            self.btn_action_other_tools,
        )
        self.project_action_group = QButtonGroup(self)
        self.project_action_group.setExclusive(True)
        for btn in project_action_buttons:
            btn.setObjectName("basicToolActionBtn")
            btn.setCheckable(True)
            btn.setText(self._wrap_button_text(btn.text(), words_per_line=2))
            btn.setFixedSize(160, 83)
            btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            self.project_action_group.addButton(btn)
        for btn in project_action_buttons:
            project_actions_layout.addWidget(btn, 0, Qt.AlignmentFlag.AlignLeft)
        project_actions_layout.addStretch(1)
        self.left_nav_stack.addWidget(self.project_left)  # index 2

        active_body_layout.addWidget(self.left_nav_stack, 0)

        # ── Right column: project selector (top, hidden by default) + right panel ─
        right_col_widget = QWidget()
        right_col_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._right_col_layout = QVBoxLayout(right_col_widget)
        self._right_col_layout.setContentsMargins(0, 0, 0, 0)
        self._right_col_layout.setSpacing(6)
        self._right_col_layout.addWidget(self.project_selector_row)

        # ── Combined right panel (all tool panels in one stack) ──────────────────
        self.combined_right_panel = QFrame()
        self.combined_right_panel.setObjectName("basicToolsRightPanel")
        self.combined_right_panel.setVisible(False)
        self.combined_right_panel.setMinimumSize(382, 437)
        self.combined_right_panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._right_col_layout.addWidget(self.combined_right_panel, 1)
        active_body_layout.addWidget(right_col_widget, 1)
        self._build_combined_right_panel()

        # MASTER button connections
        self.btn_action_status_collector.clicked.connect(self._on_status_panel_clicked)
        self.btn_action_thumbnail_generator.clicked.connect(self._on_thumbnail_panel_clicked)
        self.btn_action_sdr_master.clicked.connect(self._on_mapper_reformat_clicked)
        # PROJECT button connections
        self.btn_action_packshot_naming.clicked.connect(self._on_packshot_panel_clicked)
        self.btn_action_sap_data_reformat.clicked.connect(self._on_mapper_reformat_clicked)
        self.btn_action_sap_data_compare.clicked.connect(self._on_mapper_compare_clicked)
        self.btn_action_project_viewer.clicked.connect(self._on_project_viewer_clicked)
        self.btn_action_other_tools.clicked.connect(self._on_other_tools_clicked)

        self.empty_page = QWidget()
        _home_outer = QVBoxLayout(self.empty_page)
        _home_outer.setContentsMargins(0, 0, 0, 0)
        _home_outer.setSpacing(0)

        # White card panel matching other pages
        _home_card = QFrame()
        _home_card.setObjectName("homeCard")
        _home_card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        _home_card_layout = QHBoxLayout(_home_card)
        _home_card_layout.setContentsMargins(40, 0, 40, 0)
        _home_card_layout.setSpacing(24)
        _home_card_layout.addStretch(1)

        # Text graphic: home_page_text.png
        _home_img_label = QLabel()
        _home_img_label.setAlignment(Qt.AlignTop | Qt.AlignRight)
        _home_img_label.setScaledContents(False)
        _home_img_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        _home_img_path = str(_BASE_DIR / "logo" / "home_page_text.png")
        _home_pixmap = QPixmap(_home_img_path)
        if not _home_pixmap.isNull():
            _home_img_label.setPixmap(_home_pixmap)
        # Wrapper so we can shift the text PNG via top margin
        _text_wrapper = QWidget()
        _text_wrapper.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        _text_wrapper_layout = QVBoxLayout(_text_wrapper)
        _text_wrapper_layout.setContentsMargins(0, 0, 0, 0)
        _text_wrapper_layout.setSpacing(0)
        _text_wrapper_layout.addWidget(_home_img_label, 0, Qt.AlignTop | Qt.AlignRight)
        _text_wrapper_layout.addStretch(1)
        _home_card_layout.addWidget(_text_wrapper, 0, Qt.AlignRight)

        # Animated GIF: 360_Bottle 1.gif
        _gif_label = QLabel()
        _gif_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        _gif_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        _gif_path = str(_BASE_DIR / "logo" / "360_Bottle 1.gif")
        _gif_movie = QMovie(_gif_path)
        if _gif_movie.isValid():
            _gif_label.setMovie(_gif_movie)
            _gif_movie.start()
        # Wrapper so we can shift the GIF upward via top margin
        _gif_wrapper = QWidget()
        _gif_wrapper.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        _gif_wrapper_layout = QVBoxLayout(_gif_wrapper)
        _gif_wrapper_layout.setContentsMargins(0, 0, 0, 0)
        _gif_wrapper_layout.setSpacing(0)
        _gif_wrapper_layout.addWidget(_gif_label, 0, Qt.AlignTop | Qt.AlignLeft)
        _gif_wrapper_layout.addStretch(1)
        _home_card_layout.addWidget(_gif_wrapper, 0, Qt.AlignLeft)

        _home_card_layout.addStretch(1)
        _home_outer.addWidget(_home_card, 1)

        self._home_img_label = _home_img_label
        self._home_pixmap = _home_pixmap
        self._text_wrapper_layout = _text_wrapper_layout
        self._gif_movie = _gif_movie  # keep reference alive
        self._gif_label = _gif_label
        self._gif_wrapper_layout = _gif_wrapper_layout
        self._home_card = _home_card
        self.content_stack = QStackedWidget()
        self.content_stack.addWidget(self.empty_page)
        self.content_stack.addWidget(self.active_tools_page)
        self.search_page = self._build_search_page()
        self.content_stack.addWidget(self.search_page)
        self.artwork_page = self._build_artwork_page()
        self.content_stack.addWidget(self.artwork_page)
        content_layout.addWidget(self.content_stack, 1)
        content_layout.addWidget(self.section_title, 0, Qt.AlignLeft | Qt.AlignBottom)

        self.top_group = QButtonGroup(self)
        self.top_group.setExclusive(True)
        self._register_top_button(self.btn_home, "HOME")
        self._register_top_button(self.btn_master, "MASTER")
        self._register_top_button(self.btn_project, "PROJECT")
        self._register_top_button(self.btn_search, "SEARCH")
        self._register_top_button(self.btn_artwork, "ARTWORK")
        self.btn_home.setChecked(True)
        self._set_section_title("HOME")

        self._config = HatConfig()
        self._use_root_folders: bool = False
        self.btn_settings.clicked.connect(self._open_settings_dialog)
        self.btn_rail_toggle.toggled.connect(self._on_root_folders_toggle)

        self._init_modules()
        self.setStyleSheet(self._stylesheet())
        QTimer.singleShot(0, self._update_home_image)
        QTimer.singleShot(0, lambda: self.btn_rail_toggle.setChecked(True))

    def _make_top_button(self, text: str) -> QPushButton:
        btn = QPushButton(text.upper())
        btn.setObjectName("topHeaderBtn")
        btn.setCheckable(True)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setMinimumHeight(50)
        width_map = {
            "HOME": 90,
            "MASTER": 120,
            "PROJECT": 120,
            "SEARCH": 110,
            "ARTWORK": 120,
        }
        btn.setMinimumWidth(width_map.get(text.upper(), 120))
        return btn

    def _make_icon_button(self, symbol: str) -> QPushButton:
        btn = QPushButton(symbol)
        btn.setObjectName("iconBtn")
        btn.setCursor(Qt.PointingHandCursor)
        btn.setFixedSize(38, 38)
        return btn

    def _register_top_button(self, button: QPushButton, title: str) -> None:
        self.top_group.addButton(button)
        button.clicked.connect(lambda _checked=False, t=title: self._set_section_title(t))

    def _set_section_title(self, title: str) -> None:
        self.section_title.setText(title)
        self.section_title.setVisible(title != "SEARCH")
        if title == "MASTER":
            self.project_selector_row.setVisible(False)
            self._content_layout.setSpacing(0)
            self._active_page_layout.setContentsMargins(0, 16, 0, 0)
            self.content_stack.setCurrentWidget(self.active_tools_page)
            self.left_nav_stack.setCurrentWidget(self.master_left)
            self.btn_action_status_collector.setChecked(True)
            self._show_tracker_status_panel()
        elif title == "PROJECT":
            self._populate_project_combo()
            self.project_selector_row.setVisible(True)
            self._content_layout.setSpacing(0)
            self._active_page_layout.setContentsMargins(0, 4, 0, 0)
            self.content_stack.setCurrentWidget(self.active_tools_page)
            self.left_nav_stack.setCurrentWidget(self.project_left)
            self.btn_action_packshot_naming.setChecked(True)
            self._show_packshot_panel()
        elif title == "SEARCH":
            self.project_selector_row.setVisible(False)
            self._content_layout.setSpacing(0)
            self._active_page_layout.setContentsMargins(0, 16, 0, 0)
            self.content_stack.setCurrentWidget(self.search_page)
            self.left_nav_stack.setCurrentWidget(self.left_nav_blank)
            self.combined_right_panel.setVisible(False)
        elif title == "ARTWORK":
            self.project_selector_row.setVisible(False)
            self._content_layout.setSpacing(0)
            self.content_stack.setCurrentWidget(self.artwork_page)
            self.left_nav_stack.setCurrentWidget(self.left_nav_blank)
            self.combined_right_panel.setVisible(False)
        else:  # HOME
            self.project_selector_row.setVisible(False)
            self._content_layout.setSpacing(16)
            self.content_stack.setCurrentWidget(self.empty_page)
            self.combined_right_panel.setVisible(False)

    def _build_artwork_page(self) -> QWidget:
        """Build the Artwork processing controls and results panel."""
        outer = QWidget()
        outer_layout = QHBoxLayout(outer)
        outer_layout.setContentsMargins(0, 16, 0, 0)
        outer_layout.setSpacing(18)

        left_spacer = QWidget()
        left_spacer.setMaximumWidth(178)
        outer_layout.addWidget(left_spacer)

        panel = QFrame()
        panel.setObjectName("artworkPanel")
        panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(24, 24, 24, 24)
        panel_layout.setSpacing(12)

        title = QLabel("Artwork Processing")
        title.setObjectName("artworkTitle")
        panel_layout.addWidget(title)

        description = QLabel("Detect a dieline, crop the artwork, and export TIFF files.")
        description.setObjectName("artworkDescription")
        panel_layout.addWidget(description)

        source_row = QHBoxLayout()
        source_row.setSpacing(8)
        source_label = QLabel("PDF artwork")
        source_label.setObjectName("artworkFieldLabel")
        self.artwork_input = QLineEdit()
        self.artwork_input.setObjectName("artworkInput")
        self.artwork_input.setPlaceholderText("Select one or more PDF files")
        source_button = QPushButton("Browse")
        source_button.setObjectName("artworkBrowseButton")
        source_button.clicked.connect(self._choose_artwork_input)
        source_row.addWidget(source_label)
        source_row.addWidget(self.artwork_input, 1)
        source_row.addWidget(source_button)
        panel_layout.addLayout(source_row)

        settings_row = QHBoxLayout()
        settings_row.setSpacing(10)
        settings_row.addStretch(1)
        self.artwork_dpi = QLineEdit("300")
        self.artwork_dpi.setObjectName("artworkSmallInput")
        self.artwork_dpi.setFixedWidth(70)
        dpi_unit = QLabel("dpi")
        dpi_unit.setObjectName("artworkMutedLabel")
        settings_row.addWidget(self.artwork_dpi)
        settings_row.addWidget(dpi_unit)
        settings_row.addSpacing(16)
        self.artwork_include_diecut = QCheckBox("include w/ diecut version")
        self.artwork_include_diecut.setObjectName("artworkCheckBox")
        self.artwork_include_diecut.setChecked(True)
        settings_row.addWidget(self.artwork_include_diecut)
        settings_row.addStretch(1)
        panel_layout.addLayout(settings_row)
        panel_layout.addSpacing(12)

        manual_cut_row = QHBoxLayout()
        manual_cut_row.addStretch(1)
        self.artwork_manual_cut_button = QPushButton("Manual Cut")
        self.artwork_manual_cut_button.setObjectName("collectorRunBtn")
        self.artwork_manual_cut_button.setFixedHeight(50)
        self.artwork_manual_cut_button.setMinimumWidth(180)
        self.artwork_manual_cut_button.clicked.connect(self._open_manual_cut)
        manual_cut_row.addWidget(self.artwork_manual_cut_button)
        manual_cut_row.addStretch(1)
        panel_layout.addLayout(manual_cut_row)

        self.artwork_log = QTextEdit()
        self.artwork_log.setObjectName("artworkLog")
        self.artwork_log.setReadOnly(True)
        self.artwork_log.setPlaceholderText("Detection details and output paths will appear here.")
        panel_layout.addWidget(self.artwork_log, 1)
        outer_layout.addWidget(panel, 1)
        return outer

    def _choose_artwork_input(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "Choose artwork PDFs", "", "PDF files (*.pdf)")
        if paths:
            self._artwork_paths = [Path(path) for path in paths]
            self.artwork_input.setText(" | ".join(paths))

    def _process_artwork(self) -> None:
        input_paths = [path for path in getattr(self, "_artwork_paths", []) if path.is_file()]
        if not input_paths:
            input_path = Path(self.artwork_input.text().strip())
            if input_path.is_file() and input_path.suffix.lower() == ".pdf":
                input_paths = [input_path]
        if not input_paths:
            QMessageBox.warning(self, "Missing artwork PDF", "Choose one or more existing PDF files first.")
            return
        try:
            dpi = int(self.artwork_dpi.text().strip())
            if dpi < 72:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "Invalid resolution", "Resolution must be a whole number of at least 72 DPI.")
            return

        output_dir = input_paths[0].parent / "artwork_output"
        self.artwork_process_button.setEnabled(False)
        QApplication.processEvents()
        try:
            options = ProcessOptions(dpi=dpi, output_format="tiff", include_diecut=self.artwork_include_diecut.isChecked())
            lines: list[str] = []
            for input_path in input_paths:
                result = ArtworkProcessor(options).process_file(input_path, output_dir)
                lines.extend([
                    f"{input_path.name} | confidence {result.detection.confidence:.2f} | {result.detection.source}",
                    f"Crop: {result.detection.bounds.width:.1f} x {result.detection.bounds.height:.1f} pt",
                ])
                if result.diecut_path:
                    lines.append(f"Die-cut TIFF: {result.diecut_path}")
                lines.append(f"Clean TIFF: {result.clean_path}")
                lines.extend(f"Warning: {warning}" for warning in result.warnings)
            self.artwork_log.setPlainText("\n".join(lines))
        except (FileNotFoundError, RuntimeError, ValueError, OSError) as error:
            self.artwork_log.setPlainText(f"Processing failed: {error}")
            QMessageBox.critical(self, "Artwork processing failed", str(error))
        finally:
            self.artwork_process_button.setEnabled(True)

    def _open_manual_cut(self) -> None:
        input_paths = [path for path in getattr(self, "_artwork_paths", []) if path.is_file()]
        if not input_paths:
            input_path = Path(self.artwork_input.text().strip())
            if input_path.is_file() and input_path.suffix.lower() == ".pdf":
                input_paths = [input_path]
        if not input_paths:
            QMessageBox.warning(self, "Missing artwork PDF", "Choose one or more PDF files before opening Manual Cut.")
            return
        try:
            inspections = [inspect_artwork(path, dpi=72) for path in input_paths]
        except (FileNotFoundError, RuntimeError, ValueError, OSError) as error:
            QMessageBox.critical(self, "Artwork preview failed", str(error))
            return
        dialog = ArtworkManualCutDialog(inspections, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        output_dir = dialog.output_folder
        if output_dir is None:
            return
        try:
            dpi = int(self.artwork_dpi.text().strip())
            if dpi < 72:
                raise ValueError
            options = ProcessOptions(dpi=dpi, output_format="tiff", include_diecut=self.artwork_include_diecut.isChecked())
            lines: list[str] = []
            for input_path in input_paths:
                choice = dialog.selections().get(input_path)
                result = ArtworkProcessor(options).process_file(
                    input_path, output_dir, choice
                )
                lines.append(f"{input_path.name} | {result.detection.source} | confidence {result.detection.confidence:.2f}")
                if result.diecut_path:
                    lines.append(f"  Die-cut TIFF: {result.diecut_path}")
                lines.append(f"  Clean TIFF: {result.clean_path}")
                lines.extend(f"  Warning: {warning}" for warning in result.warnings)
            self.artwork_log.setPlainText("\n".join(lines))
        except (FileNotFoundError, RuntimeError, ValueError, OSError) as error:
            self.artwork_log.setPlainText(f"Manual processing failed: {error}")
            QMessageBox.critical(self, "Manual processing failed", str(error))

    def _update_home_image(self) -> None:
        """Scale home_page_text.png and the GIF to fit the card height."""
        from PySide6.QtCore import QSize
        card_h = self._home_card.height()
        if card_h <= 0:
            return

        # Text PNG: ~19.97% of card height (+20%), shifted up from centre
        text_target_h = max(20, int(card_h * 0.1997))
        text_top_margin = max(0, (card_h - text_target_h) // 2 - 20)
        self._text_wrapper_layout.setContentsMargins(0, text_top_margin, 0, 0)
        if not self._home_pixmap.isNull():
            scaled_pix = self._home_pixmap.scaledToHeight(
                text_target_h, Qt.TransformationMode.SmoothTransformation
            )
            self._home_img_label.setPixmap(scaled_pix)

        # GIF: ~46.59% of card height (+20%), fixed position
        gif_target_h = max(60, int(card_h * 0.4659))
        if self._gif_movie.isValid():
            nat = self._gif_movie.currentImage().size()
            if nat.height() > 0:
                gif_w = int(nat.width() * gif_target_h / nat.height())
                self._gif_movie.setScaledSize(QSize(gif_w, gif_target_h))
        gif_top_margin = max(0, (card_h - gif_target_h) // 2 - 67)
        self._gif_wrapper_layout.setContentsMargins(0, gif_top_margin, 0, 0)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._update_home_image()

    def _populate_project_combo(self) -> None:
        """Fill combo_project with folder names inside HAT DASHBOARD ROOT that are not MASTER."""
        from pathlib import Path
        self.combo_project.blockSignals(True)
        self.combo_project.clear()
        root = self._config.root_folder()
        if not root:
            self.combo_project.addItem("(no root configured)")
            self.combo_project.blockSignals(False)
            return
        hat_root = Path(root) / "HAT DASHBOARD ROOT"
        if not hat_root.is_dir():
            self.combo_project.addItem("(root not found)")
            self.combo_project.blockSignals(False)
            return
        projects = sorted(
            d.name for d in hat_root.iterdir()
            if d.is_dir() and d.name.upper() != "MASTER"
        )
        if projects:
            for name in projects:
                self.combo_project.addItem(name)
        else:
            self.combo_project.addItem("(no projects found)")
        self.combo_project.blockSignals(False)
        if self._use_root_folders:
            self._autofill_project_page()

    def _on_project_combo_changed(self) -> None:
        """Re-autofill PROJECT fields when the selected project changes."""
        if self._use_root_folders:
            self._autofill_project_page()

    def _show_basic_tools_right_panel(self) -> None:
        self.combined_panel_stack.setCurrentWidget(self.right_panel_blank)
        self.combined_right_panel.setVisible(True)

    def _show_tracker_status_panel(self) -> None:
        self.combined_panel_stack.setCurrentWidget(self.right_panel_tracker)
        self.combined_right_panel.setVisible(True)

    def _show_thumbnail_panel(self) -> None:
        self.combined_panel_stack.setCurrentWidget(self.right_panel_thumbnail)
        self.combined_right_panel.setVisible(True)

    def _show_packshot_panel(self) -> None:
        self.combined_panel_stack.setCurrentWidget(self.right_panel_packshot)
        self.combined_right_panel.setVisible(True)

    def _on_status_panel_clicked(self) -> None:
        clear_other_panel_inputs(self, active_panel="status")
        if self._use_root_folders:
            self._apply_root_folder_inputs()
        self._show_tracker_status_panel()

    def _on_thumbnail_panel_clicked(self) -> None:
        clear_other_panel_inputs(self, active_panel="thumbnail")
        if self._use_root_folders:
            self._apply_root_folder_inputs()
        self._show_thumbnail_panel()

    def _on_packshot_panel_clicked(self) -> None:
        clear_other_panel_inputs(self, active_panel="packshot")
        if self._use_root_folders:
            self._apply_root_folder_inputs()
        self._show_packshot_panel()

    def _on_mapper_reformat_clicked(self) -> None:
        if self._use_root_folders:
            self._apply_root_folder_inputs()
        self._show_mapper_reformat_panel()

    def _on_mapper_compare_clicked(self) -> None:
        self._reset_mapper_reformat_inputs()
        if self._use_root_folders:
            self._apply_root_folder_inputs()
        self._show_mapper_compare_panel()

    def _on_project_viewer_clicked(self) -> None:
        self.combined_panel_stack.setCurrentWidget(self.right_panel_project_viewer)
        self.combined_right_panel.setVisible(True)
        if self._use_root_folders:
            self._autofill_project_page()

    def _on_other_tools_clicked(self) -> None:
        self.combined_panel_stack.setCurrentWidget(self.right_panel_other_tools)
        self.combined_right_panel.setVisible(True)

    def _reset_mapper_reformat_inputs(self) -> None:
        """Return the SAP Data Reformat panel to its default state."""
        self.radio_mapper_option_1.setChecked(True)  # also triggers _sync_mapper_reformat_mode_ui
        self.input_mapper_sap_data_files.clear()
        self.input_mapper_output_location.clear()
        self._mapper_sap_file_paths = []
        self._mapper_output_location = ""
        self.check_mapper_cleanup_1.setChecked(True)
        self.check_mapper_cleanup_2.setChecked(True)
        self.check_mapper_cleanup_3.setChecked(True)
        self.check_mapper_cleanup_4.setChecked(False)
        self.checkbox_mapper_include_grouping_report.setChecked(True)
        self.mapper_reformat_progress_bar.setValue(0)
        self.mapper_reformat_progress_bar.setVisible(False)
        self.btn_run_process_mapper_reformat.setEnabled(False)

    # ── SEARCH page ────────────────────────────────────────────────────────

    def _build_search_page(self) -> QWidget:
        """Build and return the full-width SEARCH page panel."""
        outer = QWidget()
        outer_layout = QVBoxLayout(outer)
        outer_layout.setContentsMargins(0, 16, 0, 0)
        outer_layout.setSpacing(0)

        # White card panel
        panel = QFrame()
        panel.setObjectName("searchPanel")
        panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(32, 28, 32, 28)
        panel_layout.setSpacing(16)

        # ── "Expand search" checkbox ──────────────────────────────────────
        self.chk_expand_search = QCheckBox("Expand search")
        self.chk_expand_search.setObjectName("searchExpandCheckbox")
        self.chk_expand_search.setChecked(False)
        panel_layout.addWidget(self.chk_expand_search, 0, Qt.AlignmentFlag.AlignLeft)

        # ── Compact search bar: [FieldDropdown][Input][MultiBtn][SearchBtn] ─
        self.search_bar_compact = QWidget()
        self.search_bar_compact.setObjectName("searchBarCompact")
        compact_layout = QHBoxLayout(self.search_bar_compact)
        compact_layout.setContentsMargins(0, 0, 0, 0)
        compact_layout.setSpacing(0)

        # Field-selector label (dropdown styled as pill label)
        self.combo_search_field = QComboBox()
        self.combo_search_field.setObjectName("searchFieldLabel")
        self.combo_search_field.addItems([
            "IDH", "Basic", "Pack Name", "Pack Type", "Pack Size",
            "Project Name", "Label Size", "SBU", "Custom",
        ])
        self.combo_search_field.setCurrentText("IDH")
        self.combo_search_field.setFixedWidth(120)
        compact_layout.addWidget(self.combo_search_field, 0)

        # Text input
        self.input_search = QLineEdit()
        self.input_search.setObjectName("searchFieldInput")
        self.input_search.setPlaceholderText("Search…")
        self.input_search.setClearButtonEnabled(True)
        compact_layout.addWidget(self.input_search, 1)

        # Multi-value add button (SAP-style, inline at end of input)
        self.btn_compact_multi = QPushButton("⊞")
        self.btn_compact_multi.setObjectName("searchMultiAddBtn")
        self.btn_compact_multi.setFixedSize(36, 40)
        self.btn_compact_multi.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_compact_multi.setToolTip("Add multiple values for this field")
        compact_layout.addWidget(self.btn_compact_multi, 0)

        # Search button (magnifier)
        self.btn_search_go = QPushButton("⌕")
        self.btn_search_go.setObjectName("searchIconBtn")
        self.btn_search_go.setFixedSize(46, 40)
        self.btn_search_go.setCursor(Qt.CursorShape.PointingHandCursor)
        compact_layout.addWidget(self.btn_search_go, 0)

        # ── Shared filter row: Source / Asset Type / Sample Limit ────────
        # Always visible regardless of compact vs expanded state, placed at top
        shared_filters = QWidget()
        shared_filters_layout = QHBoxLayout(shared_filters)
        shared_filters_layout.setContentsMargins(0, 0, 0, 8)
        shared_filters_layout.setSpacing(16)

        def _make_mini_combo(label: str, items: list, default: str, attr: str) -> QWidget:
            col = QVBoxLayout()
            col.setSpacing(3)
            col.setContentsMargins(0, 0, 0, 0)
            lbl = QLabel(label)
            lbl.setObjectName("searchDropdownLabel")
            col.addWidget(lbl)
            cb = QComboBox()
            cb.setObjectName("searchDropdownMini")
            cb.addItems(items)
            cb.setCurrentText(default)
            setattr(self, attr, cb)
            col.addWidget(cb)
            w = QWidget()
            w.setLayout(col)
            return w

        shared_filters_layout.addWidget(
            _make_mini_combo("Source", ["TSC", "RSD (master)", "Library"], "TSC",
                             "combo_search_source"))
        shared_filters_layout.addWidget(
            _make_mini_combo("Asset Type", ["All", "2D", "3D"], "All",
                             "combo_search_asset_type"))
        shared_filters_layout.addWidget(
            _make_mini_combo("Sample Limit", ["5", "10", "All"], "5",
                             "combo_search_sample_limit"))

        # "Only with matching image" checkbox — always visible, aligned with dropdowns
        chk_col = QVBoxLayout()
        chk_col.setSpacing(3)
        chk_col.setContentsMargins(12, 0, 0, 0)
        _chk_spacer_lbl = QLabel("")
        _chk_spacer_lbl.setObjectName("searchDropdownLabel")
        chk_col.addWidget(_chk_spacer_lbl)
        self.chk_search_matching_image = QCheckBox("Only with matching image")
        self.chk_search_matching_image.setObjectName("searchExpandCheckbox")
        self.chk_search_matching_image.setChecked(False)
        chk_col.addWidget(self.chk_search_matching_image)
        chk_w = QWidget()
        chk_w.setLayout(chk_col)
        shared_filters_layout.addWidget(chk_w)

        shared_filters_layout.addStretch(1)
        panel_layout.addWidget(shared_filters)

        panel_layout.addWidget(self.search_bar_compact)

        # ── Expanded multi-field search (hidden by default) ───────────────
        self.search_bar_expanded = QWidget()
        self.search_bar_expanded.setVisible(False)
        expanded_layout = QVBoxLayout(self.search_bar_expanded)
        expanded_layout.setContentsMargins(0, 0, 0, 0)
        expanded_layout.setSpacing(10)

        # 7 text fields arranged in 2 columns
        _expand_fields = [
            ("IDH",          "input_search_idh",          "000000"),
            ("Pack Name",    "input_search_pack_name",    "loctite 401, Teroson MS 949 FR"),
            ("Basic",        "input_search_basic",        "000000"),
            ("Pack Type",    "input_search_pack_type",    "bottle"),
            ("Pack Size",    "input_search_pack_size",    "500ml"),
            ("Label Size",   "input_search_label_size",   "8 x 12"),
            ("Project Name", "input_search_project_name", "2025A002"),
            ("Custom",       "input_search_custom",       "Anaerobic, right-angle"),
        ]

        def _make_field_pill(label_text: str, attr_name: str, placeholder: str) -> QHBoxLayout:
            pill = QHBoxLayout()
            pill.setSpacing(0)
            lbl = QLineEdit()
            lbl.setObjectName("searchFieldLabel")
            lbl.setText(label_text)
            lbl.setReadOnly(True)
            lbl.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            lbl.setFixedWidth(115)
            pill.addWidget(lbl, 0)
            field = QLineEdit()
            field.setObjectName("searchFieldInput")
            field.setPlaceholderText(placeholder)
            field.setClearButtonEnabled(True)
            setattr(self, attr_name, field)
            pill.addWidget(field, 1)
            return pill

        # Build rows of 2 columns
        grid_layout = QGridLayout()
        grid_layout.setHorizontalSpacing(10)
        grid_layout.setVerticalSpacing(2)
        grid_layout.setContentsMargins(0, 0, 0, 0)
        for i, (label_text, attr_name, placeholder) in enumerate(_expand_fields):
            col = i % 2
            row_idx = i // 2
            pill_widget = QWidget()
            pill_widget.setLayout(_make_field_pill(label_text, attr_name, placeholder))
            grid_layout.addWidget(pill_widget, row_idx, col)
        expanded_layout.addLayout(grid_layout)

        expanded_layout.addSpacing(6)

        # ── Search button centered below all fields ───────────────────────
        btn_row = QHBoxLayout()
        btn_row.setContentsMargins(0, 8, 0, 0)
        self.btn_search_go_expanded = QPushButton("Search")
        self.btn_search_go_expanded.setObjectName("collectorRunBtn")
        self.btn_search_go_expanded.setFixedHeight(50)
        self.btn_search_go_expanded.setMinimumWidth(130)
        self.btn_search_go_expanded.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_search_go_expanded)
        btn_row.addStretch(1)
        expanded_layout.addLayout(btn_row)

        panel_layout.addWidget(self.search_bar_expanded)

        panel_layout.addStretch(1)

        outer_layout.addWidget(panel, 1)

        # ── wire controls ────────────────────────────────────────────────
        self.chk_expand_search.toggled.connect(self._on_expand_search_toggled)
        self.combo_search_source.currentTextChanged.connect(self._on_search_source_changed)
        self.combo_search_field.currentTextChanged.connect(self._on_compact_field_changed)
        self.btn_compact_multi.clicked.connect(self._on_compact_multi_clicked)
        self.input_search_custom.textChanged.connect(self._on_search_custom_changed)
        self.input_search_custom.installEventFilter(self)
        self.btn_search_go.clicked.connect(self._on_search_go_clicked)
        self.btn_search_go_expanded.clicked.connect(self._on_search_go_clicked)
        self._search_multi_values: dict[str, list[str]] = {}
        self._search_multi_active_attr: str | None = None
        self._search_compact_multi_values: dict[str, list[str]] = {}

        # Wire all text inputs to the button state validator
        self.input_search.textChanged.connect(self._update_search_button_state)
        for attr in self._FIELD_ATTRS.values():
            w = getattr(self, attr, None)
            if w:
                w.textChanged.connect(self._update_search_button_state)
        self.chk_expand_search.toggled.connect(self._update_search_button_state)
        self._update_search_button_state()

        return outer

    def _on_expand_search_toggled(self, checked: bool) -> None:
        self.search_bar_compact.setVisible(not checked)
        self.search_bar_expanded.setVisible(checked)

    def _update_search_button_state(self, *_) -> None:
        """Enable/disable search buttons depending on whether any input has a value."""
        is_expanded = self.chk_expand_search.isChecked()
        if is_expanded:
            has_value = False
            for attr in self._FIELD_ATTRS.values():
                w = getattr(self, attr, None)
                if w and w.text().strip():
                    has_value = True
                    break
            if not has_value:
                has_value = any(bool(v) for v in self._search_multi_values.values())
        else:
            label = self.combo_search_field.currentText()
            compact_vals = self._search_compact_multi_values.get(label, [])
            has_value = bool(compact_vals) or bool(self.input_search.text().strip())
        self.btn_search_go.setEnabled(has_value)
        self.btn_search_go_expanded.setEnabled(has_value)

    # Maps compact dropdown label → field key used by run_search
    _COMPACT_FIELD_MAP: dict[str, str] = {
        "IDH":          "idh",
        "Basic":        "basic",
        "Pack Name":    "pack_name",
        "Pack Type":    "pack_type",
        "Pack Size":    "pack_size",
        "Project Name": "project_name",
        "Label Size":   "label_size",
        "SBU":          "sbu",
        "Custom":       "custom",
    }

    def _on_compact_field_changed(self, label: str) -> None:
        """Called when the compact field dropdown changes — reset display."""
        self._refresh_compact_input_display()
        self._update_search_button_state()

    def _on_compact_multi_clicked(self) -> None:
        """Open multi-value editor for the currently selected compact field."""
        label = self.combo_search_field.currentText()
        initial = self._search_compact_multi_values.get(label, [])
        dlg = _SearchMultiValueDialog(label, initial, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._search_compact_multi_values[label] = dlg.get_values()
            self._refresh_compact_input_display()
            self._update_search_button_state()

    def _refresh_compact_input_display(self) -> None:
        """Update the compact input to show multi-value summary or be editable."""
        label = self.combo_search_field.currentText()
        vals = self._search_compact_multi_values.get(label, [])
        if vals:
            count = len(vals)
            preview = ", ".join(vals[:3])
            if count > 3:
                preview += f", … (+{count - 3} more)"
            self.input_search.setText(
                f"{count} value{'s' if count != 1 else ''}: {preview}"
            )
            self.input_search.setReadOnly(True)
            self.input_search.setClearButtonEnabled(False)
        else:
            self.input_search.clear()
            self.input_search.setReadOnly(False)
            self.input_search.setClearButtonEnabled(True)
            self.input_search.setPlaceholderText("Search…")

    def _on_search_source_changed(self, source: str) -> None:
        """Lock/unlock Asset Type dropdown based on Source selection."""
        is_library = source == "Library"
        is_rsd = source == "RSD (master)"
        # Library → lock to "3D"; RSD → lock to "All"; otherwise free
        if is_library:
            self.combo_search_asset_type.setCurrentText("3D")
            self.combo_search_asset_type.setEnabled(False)
        elif is_rsd:
            self.combo_search_asset_type.setCurrentText("All")
            self.combo_search_asset_type.setEnabled(False)
        else:
            self.combo_search_asset_type.setEnabled(True)

    # Fields that Custom freezes (all expanded fields except Custom itself)
    _CUSTOM_FREEZES: tuple[str, ...] = (
        "input_search_idh",
        "input_search_pack_name",
        "input_search_basic",
        "input_search_pack_type",
        "input_search_pack_size",
        "input_search_label_size",
        "input_search_project_name",
    )

    def _on_search_custom_changed(self, text: str) -> None:
        """Freeze/unfreeze other search fields based on whether Custom has a value."""
        self._set_custom_freeze_state(bool(text.strip()))

    def _set_custom_freeze_state(self, freeze: bool) -> None:
        """Enable or disable all non-Custom search input fields."""
        for attr in self._CUSTOM_FREEZES:
            field = getattr(self, attr, None)
            if field is None:
                continue
            field.setReadOnly(freeze)
            field.setEnabled(not freeze)

    # ── Search execution ─────────────────────────────────────────────────

    #  Maps field_key → attr name on self
    _FIELD_ATTRS: dict[str, str] = {
        "idh":          "input_search_idh",
        "pack_name":    "input_search_pack_name",
        "basic":        "input_search_basic",
        "pack_type":    "input_search_pack_type",
        "pack_size":    "input_search_pack_size",
        "label_size":   "input_search_label_size",
        "project_name": "input_search_project_name",
        "custom":       "input_search_custom",
    }

    def _on_search_go_clicked(self) -> None:
        """Collect inputs, run search, open the Search Result window."""
        from search import run_search

        is_expanded = self.chk_expand_search.isChecked()

        # Collect field values
        field_values: dict[str, str] = {}
        if is_expanded:
            for field_key, attr in self._FIELD_ATTRS.items():
                w = getattr(self, attr, None)
                field_values[field_key] = w.text() if w else ""
        else:
            # Compact mode: use the selected field from the dropdown
            label = self.combo_search_field.currentText()
            field_key = self._COMPACT_FIELD_MAP.get(label, "custom")
            compact_vals = self._search_compact_multi_values.get(label, [])
            if not compact_vals:
                field_values[field_key] = self.input_search.text()

        # Convert multi-select values (attr_name → field_key) for expanded mode
        _attr_to_field = {v: k for k, v in self._FIELD_ATTRS.items()}
        multi_values: dict[str, list[str]] = {
            _attr_to_field[attr]: vals
            for attr, vals in self._search_multi_values.items()
            if attr in _attr_to_field and vals
        }

        # For compact mode, inject compact multi-values if set
        if not is_expanded:
            label = self.combo_search_field.currentText()
            compact_vals = self._search_compact_multi_values.get(label, [])
            if compact_vals:
                field_key = self._COMPACT_FIELD_MAP.get(label, "custom")
                multi_values[field_key] = compact_vals

        # Dropdown values
        source     = self.combo_search_source.currentText()
        asset_type = self.combo_search_asset_type.currentText()
        limit_txt  = self.combo_search_sample_limit.currentText()
        sample_limit: int | None = None if limit_txt == "All" else int(limit_txt)

        # Build human-readable query summary for display in results window
        search_summary_parts: list[str] = []
        if is_expanded:
            for field_key, attr in self._FIELD_ATTRS.items():
                label = field_key.replace("_", " ").title()
                multi = self._search_multi_values.get(attr, [])
                if multi:
                    search_summary_parts.append(f"{label}: {', '.join(multi)}")
                else:
                    w = getattr(self, attr, None)
                    val = w.text().strip() if w else ""
                    if val:
                        search_summary_parts.append(f"{label}: {val}")
        else:
            label = self.combo_search_field.currentText()
            compact_vals = self._search_compact_multi_values.get(label, [])
            if compact_vals:
                search_summary_parts.append(f"{label}: {', '.join(compact_vals)}")
            else:
                val = self.input_search.text().strip()
                if val:
                    search_summary_parts.append(f"{label}: {val}")
        search_summary = "  |  ".join(search_summary_parts) if search_summary_parts else "(no filter)"

        # ── Determine if a grouped (per-cell-value) search is needed ─────
        # A grouped search runs one query per cell value whenever any single
        # field has more than one multi-value entry.  Each cell value is kept
        # intact (no splitting); results are tagged with their originating
        # query value for grouped display in the results window.
        grouped_field: str | None = None
        grouped_vals: list[str] = []
        for fk, vals in multi_values.items():
            if len(vals) > 1:
                grouped_field = fk
                grouped_vals = vals
                break

        if grouped_field:
            # Build summary: field label + cell values separated by " | "
            if is_expanded:
                field_label = grouped_field.replace("_", " ").title()
            else:
                field_label = self.combo_search_field.currentText()
            search_summary = field_label + ":  " + "  |  ".join(grouped_vals)

            base_mv = {k: v for k, v in multi_values.items() if k != grouped_field}
            all_rows: list[dict] = []
            all_idh: list[str] = []
            groups: list[dict] = []
            first_error: str | None = None
            first_source_key = "tsc"

            for cell_val in grouped_vals:
                g_result = run_search(
                    field_values=field_values,
                    multi_values={**base_mv, grouped_field: [cell_val]},
                    source=source,
                    asset_type=asset_type,
                    status="All",
                    sample_limit=sample_limit,
                    cfg=self._config,
                )
                if g_result["error"] and not first_error:
                    first_error = g_result["error"]
                    continue
                first_source_key = g_result["source_key"]
                all_rows.extend(g_result["rows"])
                all_idh.extend(g_result["idh_list"])
                groups.append({
                    "query_value":   cell_val,
                    "rows":          g_result["rows"],
                    "idh_list":      g_result["idh_list"],
                    "active_filters": g_result["active_filters"],
                })

            if first_error and not groups:
                QMessageBox.warning(self, "Search Error", first_error)
                return

            result = {
                "rows":           all_rows,
                "active_filters": groups[0]["active_filters"] if groups else set(),
                "source_key":     first_source_key,
                "idh_list":       list(dict.fromkeys(all_idh)),
                "error":          None,
                "groups":         groups,
            }
        else:
            result = run_search(
                field_values=field_values,
                multi_values=multi_values,
                source=source,
                asset_type=asset_type,
                status="All",
                sample_limit=sample_limit,
                cfg=self._config,
            )

            if result["error"]:
                QMessageBox.warning(self, "Search Error", result["error"])
                return

        result["search_summary"] = search_summary
        result["field_values"]   = field_values
        result["multi_values"]   = multi_values
        result["source"]         = source
        result["show_item_count"] = (sample_limit is None)  # True when limit is "All"

        # Filter to only rows that have a matching image if checkbox is ticked
        if self.chk_search_matching_image.isChecked():
            from search import find_matching_images
            thumb_folder = self._config.search_thumbnails_folder()
            fallbacks = [self._config.thumbnail_output(), self._config.thumbnail_input()]
            idh_set = {v.strip() for v in result["idh_list"] if v.strip()}
            matched = find_matching_images(idh_set, thumb_folder, fallback_folders=fallbacks)
            matched_idhs = {info["idh"] for info in matched if info.get("path")}
            result["rows"] = [r for r in result["rows"] if r.get("IDH", "").strip() in matched_idhs]
            result["idh_list"] = [idh for idh in result["idh_list"] if idh.strip() in matched_idhs]
            if result.get("groups"):
                for grp in result["groups"]:
                    grp["rows"] = [r for r in grp["rows"] if r.get("IDH", "").strip() in matched_idhs]
                    grp["idh_list"] = [idh for idh in grp.get("idh_list", []) if idh.strip() in matched_idhs]

        dlg = _SearchResultWindow(result, self._config, use_root_folders=self._use_root_folders, parent=self)
        dlg.setModal(False)
        dlg.show()

    # ── "Multiple Selection" multi-value helpers ──────────────────────────

    _MULTI_SELECT_FIELD_MAP: dict[str, str] = {
        "IDH":          "input_search_idh",
        "Pack Name":    "input_search_pack_name",
        "Basic":        "input_search_basic",
        "Pack Type":    "input_search_pack_type",
        "Pack Size":    "input_search_pack_size",
        "Label Size":   "input_search_label_size",
        "Project Name": "input_search_project_name",
        "Custom":       "input_search_custom",
    }

    _MULTI_SELECT_LABEL_MAP: dict[str, str] = {
        v: k for k, v in _MULTI_SELECT_FIELD_MAP.items()
    }

    def _freeze_non_multi_fields(self, active_attr: str | None) -> None:
        """Disable all search inputs except the active multi-select field.

        Pass *None* to unfreeze everything.
        """
        for attr in self._MULTI_SELECT_FIELD_MAP.values():
            if attr == active_attr:
                continue
            field = getattr(self, attr, None)
            if field is None:
                continue
            if active_attr is None:
                field.setReadOnly(False)
                field.setEnabled(True)
            else:
                field.setReadOnly(True)
                field.setEnabled(False)

    def _on_search_multi_select_changed(self, text: str) -> None:
        """Activate / deactivate multi-value mode for a search field."""
        # Restore previous active field
        prev = self._search_multi_active_attr
        if prev:
            field = getattr(self, prev, None)
            if field:
                field.removeEventFilter(self)
                field.setReadOnly(False)
                field.setClearButtonEnabled(True)
                field.setCursor(Qt.CursorShape.IBeamCursor)
                field.setStyleSheet("")
                vals = self._search_multi_values.get(prev, [])
                field.setText(", ".join(vals) if vals else "")
            self._search_multi_active_attr = None

        if text == "No":
            self._freeze_non_multi_fields(None)
            return

        attr_name = self._MULTI_SELECT_FIELD_MAP.get(text)
        if not attr_name:
            return

        self._search_multi_active_attr = attr_name
        self._freeze_non_multi_fields(attr_name)
        field = getattr(self, attr_name, None)
        if not field:
            return

        # Capture any existing plain text as first multi-value if not yet set
        if attr_name not in self._search_multi_values:
            existing = field.text().strip()
            self._search_multi_values[attr_name] = [existing] if existing else []

        field.setReadOnly(True)
        field.setClearButtonEnabled(False)
        field.setCursor(Qt.CursorShape.PointingHandCursor)
        field.setStyleSheet("""
            QLineEdit {
                background-color: #E8EEF8;
                color: #111F35;
                border: none;
                border-top-left-radius: 0px;
                border-bottom-left-radius: 0px;
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
                min-height: 40px;
                padding: 0 12px;
                font-family: "Segoe UI";
                font-size: 13px;
                font-style: italic;
            }
            QLineEdit:hover { background-color: #D6E2F5; }
        """)
        self._refresh_search_multi_field_display(attr_name)
        field.installEventFilter(self)

    def _refresh_search_multi_field_display(self, attr_name: str) -> None:
        """Update the read-only display text for the active multi-value field."""
        field = getattr(self, attr_name, None)
        if not field:
            return
        vals = self._search_multi_values.get(attr_name, [])
        if vals:
            count = len(vals)
            preview = ", ".join(vals[:3])
            if count > 3:
                preview += f", … (+{count - 3} more)"
            field.setText(f"{count} value{'s' if count != 1 else ''}: {preview}")
        else:
            field.setText("")
            field.setPlaceholderText("Click to add values…")

    def _open_search_multi_value_dialog(self, attr_name: str) -> None:
        """Open the multi-value editor for the given field attribute."""
        label_text = self._MULTI_SELECT_LABEL_MAP.get(attr_name, attr_name)
        initial = self._search_multi_values.get(attr_name, [])
        dlg = _SearchMultiValueDialog(label_text, initial, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._search_multi_values[attr_name] = dlg.get_values()
            self._refresh_search_multi_field_display(attr_name)
            self._update_search_button_state()

    def eventFilter(self, obj, event) -> bool:  # type: ignore[override]
        # Custom field: gray others on focus-in; revert on focus-out if empty
        custom = getattr(self, "input_search_custom", None)
        if custom is not None and obj is custom:
            if event.type() == QEvent.Type.FocusIn:
                self._set_custom_freeze_state(True)
            elif event.type() == QEvent.Type.FocusOut:
                if not custom.text().strip():
                    self._set_custom_freeze_state(False)

        attr = getattr(self, "_search_multi_active_attr", None)
        if attr:
            field = getattr(self, attr, None)
            if obj is field and event.type() == QEvent.Type.MouseButtonPress:
                self._open_search_multi_value_dialog(attr)
                return True
        return super().eventFilter(obj, event)

    def _build_combined_right_panel(self) -> None:
        panel_layout = QVBoxLayout(self.combined_right_panel)
        panel_layout.setContentsMargins(16, 16, 16, 16)
        panel_layout.setSpacing(0)

        self.combined_panel_stack = QStackedWidget()
        self.right_panel_blank = QWidget()
        self.right_panel_tracker = self._create_tracker_status_collector_page()
        self.right_panel_thumbnail = self._create_thumbnail_generator_page()
        self.right_panel_packshot = self._create_packshot_naming_page()
        self.mapper_right_panel_reformat = self._create_mapper_reformat_page()
        self.mapper_right_panel_compare = self._create_mapper_compare_page()
        self.right_panel_project_viewer = self._create_project_review_page()
        self.right_panel_other_tools = self._create_other_tools_page()
        self.combined_panel_stack.addWidget(self.right_panel_blank)
        self.combined_panel_stack.addWidget(self.right_panel_tracker)
        self.combined_panel_stack.addWidget(self.right_panel_thumbnail)
        self.combined_panel_stack.addWidget(self.right_panel_packshot)
        self.combined_panel_stack.addWidget(self.mapper_right_panel_reformat)
        self.combined_panel_stack.addWidget(self.mapper_right_panel_compare)
        self.combined_panel_stack.addWidget(self.right_panel_project_viewer)
        self.combined_panel_stack.addWidget(self.right_panel_other_tools)
        panel_layout.addWidget(self.combined_panel_stack)

    def _create_tracker_status_collector_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        title = QLabel("Trackers Status Collector")
        title.setObjectName("collectorTitle")
        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.addWidget(title, 1)
        layout.addLayout(title_row)

        # --- Option 1 / Option 2 radios ---
        self.radio_tsc_option_1 = QRadioButton("Option 1")
        self.radio_tsc_option_1.setObjectName("collectorModeRadio")
        self.radio_tsc_option_2 = QRadioButton("Option 2")
        self.radio_tsc_option_2.setObjectName("collectorModeRadio")
        self.radio_tsc_option_1.setChecked(True)

        self.tsc_option_group = QButtonGroup(page)
        self.tsc_option_group.setExclusive(True)
        self.tsc_option_group.addButton(self.radio_tsc_option_1)
        self.tsc_option_group.addButton(self.radio_tsc_option_2)

        layout.addWidget(self.radio_tsc_option_1)

        # Option 1 – Open Window button
        opt1_row = QHBoxLayout()
        opt1_row.setSpacing(12)
        self.btn_tsc_open_window = QPushButton("Open window")
        self.btn_tsc_open_window.setObjectName("collectorGrayBtn")
        opt1_row.addWidget(self.btn_tsc_open_window, 0)
        opt1_row.addStretch(1)
        layout.addLayout(opt1_row)

        layout.addSpacing(8)

        tsc_opt2_header = QHBoxLayout()
        tsc_opt2_header.setSpacing(0)
        tsc_opt2_header.addWidget(self.radio_tsc_option_2)
        tsc_opt2_header.addSpacing(80)
        self.lbl_tsc_input_count = QLabel("")
        self.lbl_tsc_input_count.setObjectName("inputCountLabel")
        tsc_opt2_header.addWidget(self.lbl_tsc_input_count, 0)
        tsc_opt2_header.addStretch(1)
        layout.addLayout(tsc_opt2_header)

        # Option 2 – existing elements
        trackers_row = QHBoxLayout()
        trackers_row.setSpacing(12)
        self.btn_sc_select_trackers = QPushButton("Select Trackers")
        self.btn_sc_select_trackers.setObjectName("collectorGrayBtn")
        trackers_row.addWidget(self.btn_sc_select_trackers, 0)
        self.input_trackers = QLineEdit()
        self.input_trackers.setObjectName("collectorLineEdit")
        trackers_row.addWidget(self.input_trackers, 1)
        layout.addLayout(trackers_row)

        output_row = QHBoxLayout()
        output_row.setSpacing(12)
        self.btn_sc_output_location = QPushButton("Output Report Location")
        self.btn_sc_output_location.setObjectName("collectorGrayBtn")
        output_row.addWidget(self.btn_sc_output_location, 0)
        self.input_output = QLineEdit()
        self.input_output.setObjectName("collectorLineEdit")
        output_row.addWidget(self.input_output, 1)
        layout.addLayout(output_row)

        # Select Status / Input Status sub-radios (styled like Cleanup radios)
        self.radioButton_status_select = QRadioButton("Select Status")
        self.radioButton_status_select.setObjectName("collectorCleanupRadio")
        self.radioButton_status_input = QRadioButton("Input Status")
        self.radioButton_status_input.setObjectName("collectorCleanupRadio")
        self.radioButton_status_select.setChecked(True)
        self.collector_mode_group = QButtonGroup(page)
        self.collector_mode_group.setExclusive(True)
        self.collector_mode_group.addButton(self.radioButton_status_select)
        self.collector_mode_group.addButton(self.radioButton_status_input)
        layout.addWidget(self.radioButton_status_select)

        statuses_box = QVBoxLayout()
        statuses_box.setSpacing(6)
        # Row 1 – the 5 status checkboxes + Apply Cleanup
        statuses_row1 = QHBoxLayout()
        statuses_row1.setSpacing(8)
        self.checkbox_status_collector_to_do = QCheckBox("To Do")
        self.checkbox_status_collector_cancelled = QCheckBox("Cancelled")
        self.checkbox_status_collector_completed = QCheckBox("Completed")
        self.checkbox_status_collector_completed.setChecked(True)
        self.checkbox_status_collector_in_progress = QCheckBox("In Progress")
        self.checkbox_status_collector_on_hold = QCheckBox("On Hold")
        status_checkboxes = [
            self.checkbox_status_collector_to_do,
            self.checkbox_status_collector_in_progress,
            self.checkbox_status_collector_completed,
            self.checkbox_status_collector_on_hold,
            self.checkbox_status_collector_cancelled,
        ]
        for cb in status_checkboxes:
            cb.setObjectName("collectorCheckSmall")
            statuses_row1.addWidget(cb)
        # Apply Cleanup on the same row, after Cancelled
        self.checkbox_sc_apply_cleanup = QCheckBox("Apply Cleanup")
        self.checkbox_sc_apply_cleanup.setObjectName("collectorCheckSmall")
        self.checkbox_sc_apply_cleanup.setChecked(True)
        statuses_row1.addWidget(self.checkbox_sc_apply_cleanup)
        statuses_row1.addStretch(1)
        statuses_box.addLayout(statuses_row1)
        statuses_indent_row = QHBoxLayout()
        statuses_indent_row.setContentsMargins(24, 0, 0, 0)
        statuses_indent_row.addLayout(statuses_box)
        layout.addLayout(statuses_indent_row)
        layout.addSpacing(4)
        layout.addWidget(self.radioButton_status_input)

        self._input_status_line = QLineEdit()
        self._input_status_line.setObjectName("collectorLineEdit")
        self._input_status_line.setPlaceholderText("completed")
        self._input_status_line.setText("completed")
        layout.addWidget(self._input_status_line)

        layout.addStretch(1)
        run_row = QHBoxLayout()
        run_row.addStretch(1)
        self.btn_run_process_all_trackers_status_collector = QPushButton("Run Process")
        self.btn_run_process_all_trackers_status_collector.setObjectName("collectorRunBtn")
        self.btn_run_process_all_trackers_status_collector.setFixedHeight(50)
        self.btn_run_process_all_trackers_status_collector.setMinimumWidth(180)
        run_row.addWidget(self.btn_run_process_all_trackers_status_collector)
        run_row.addStretch(1)
        layout.addLayout(run_row)
        layout.addSpacing(8)

        progress_row = QHBoxLayout()
        progress_row.addStretch(1)
        self.collector_progress_bar = QProgressBar()
        self.collector_progress_bar.setObjectName("collectorProgressBar")
        self.collector_progress_bar.setMinimumWidth(260)
        self.collector_progress_bar.setMaximumWidth(360)
        self.collector_progress_bar.setFixedHeight(14)
        self.collector_progress_bar.setTextVisible(False)
        self.collector_progress_bar.setRange(0, 100)
        self.collector_progress_bar.setValue(0)
        self.collector_progress_bar.setVisible(False)
        progress_row.addWidget(self.collector_progress_bar)
        progress_row.addStretch(1)
        layout.addLayout(progress_row)

        self.radio_tsc_option_1.toggled.connect(self._sync_tsc_mode_ui)
        self.radio_tsc_option_2.toggled.connect(self._sync_tsc_mode_ui)
        self.btn_tsc_open_window.clicked.connect(self._open_tsc_option1_table)
        self.input_trackers.textChanged.connect(lambda: self._update_tsc_input_count())
        self._sync_tsc_mode_ui()

        return page

    def _sync_tsc_mode_ui(self) -> None:
        is_opt1 = self.radio_tsc_option_1.isChecked()

        self.btn_tsc_open_window.setEnabled(is_opt1)

        opt2_widgets = [
            self.btn_sc_select_trackers,
            self.input_trackers,
            self.btn_sc_output_location,
            self.input_output,
            self.radioButton_status_select,
            self.radioButton_status_input,
            self.checkbox_status_collector_to_do,
            self.checkbox_status_collector_cancelled,
            self.checkbox_status_collector_completed,
            self.checkbox_status_collector_in_progress,
            self.checkbox_status_collector_on_hold,
            self._input_status_line,
            self.checkbox_sc_apply_cleanup,
            self.btn_run_process_all_trackers_status_collector,
        ]
        for w in opt2_widgets:
            w.setEnabled(not is_opt1)

        if is_opt1:
            self.input_trackers.clear()
            self.input_output.clear()
            self.radioButton_status_select.setChecked(True)
            self.checkbox_status_collector_completed.setChecked(True)
            self.checkbox_status_collector_to_do.setChecked(False)
            self.checkbox_status_collector_cancelled.setChecked(False)
            self.checkbox_status_collector_in_progress.setChecked(False)
            self.checkbox_status_collector_on_hold.setChecked(False)
            self._input_status_line.setText("completed")
        else:
            self.checkbox_sc_apply_cleanup.setChecked(True)

    def _create_thumbnail_generator_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(16)

        title = QLabel("Thumbnail Generator")
        title.setObjectName("thumbnailTitle")
        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.addWidget(title, 1)
        layout.addLayout(title_row)

        description = QLabel(
            "NOTE: Thumbnail sized packshots cropped edge to edge used for library\n"
            "and reference collector."
        )
        description.setObjectName("thumbnailDescription")
        layout.addWidget(description)
        layout.addSpacing(12)

        images_row = QHBoxLayout()
        images_row.setSpacing(12)
        self.btn_pg7_images_folder = QPushButton("Images Folder")
        self.btn_pg7_images_folder.setObjectName("collectorGrayBtn")
        images_row.addWidget(self.btn_pg7_images_folder, 0)
        self.input_pg7_images_folder = QLineEdit()
        self.input_pg7_images_folder.setObjectName("collectorLineEdit")
        images_row.addWidget(self.input_pg7_images_folder, 1)
        layout.addLayout(images_row)

        output_row = QHBoxLayout()
        output_row.setSpacing(12)
        self.btn_pg7_output = QPushButton("Output Location")
        self.btn_pg7_output.setObjectName("collectorGrayBtn")
        output_row.addWidget(self.btn_pg7_output, 0)
        self.input_pg7_output = QLineEdit()
        self.input_pg7_output.setObjectName("collectorLineEdit")
        output_row.addWidget(self.input_pg7_output, 1)
        layout.addLayout(output_row)

        layout.addStretch(1)
        run_row = QHBoxLayout()
        run_row.addStretch(1)
        self.btn_run_process_generate_thumbnails = QPushButton("Run Process")
        self.btn_run_process_generate_thumbnails.setObjectName("collectorRunBtn")
        self.btn_run_process_generate_thumbnails.setFixedHeight(50)
        self.btn_run_process_generate_thumbnails.setMinimumWidth(180)
        run_row.addWidget(self.btn_run_process_generate_thumbnails)
        run_row.addStretch(1)
        layout.addLayout(run_row)
        layout.addSpacing(8)

        progress_row = QHBoxLayout()
        progress_row.addStretch(1)
        self.thumbnail_progress_bar = QProgressBar()
        self.thumbnail_progress_bar.setObjectName("collectorProgressBar")
        self.thumbnail_progress_bar.setMinimumWidth(260)
        self.thumbnail_progress_bar.setMaximumWidth(360)
        self.thumbnail_progress_bar.setFixedHeight(14)
        self.thumbnail_progress_bar.setTextVisible(False)
        self.thumbnail_progress_bar.setRange(0, 100)
        self.thumbnail_progress_bar.setValue(0)
        self.thumbnail_progress_bar.setVisible(False)
        progress_row.addWidget(self.thumbnail_progress_bar)
        progress_row.addStretch(1)
        layout.addLayout(progress_row)

        return page

    def _show_mapper_reformat_panel(self) -> None:
        self.combined_panel_stack.setCurrentWidget(self.mapper_right_panel_reformat)
        self.combined_right_panel.setVisible(True)

    def _show_mapper_compare_panel(self) -> None:
        self.combined_panel_stack.setCurrentWidget(self.mapper_right_panel_compare)
        self.combined_right_panel.setVisible(True)

    def _create_mapper_reformat_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        title = QLabel("SAP Data Reformat")
        title.setObjectName("mapperTitle")
        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.addWidget(title, 1)
        layout.addLayout(title_row)

        note = QLabel("NOTE: Do not alter SAP data file. This tool is heavily based on the default state.")
        note.setObjectName("mapperDescription")
        note.setWordWrap(True)
        layout.addWidget(note)
        layout.addSpacing(8)

        self.radio_mapper_option_1 = QRadioButton("Option 1")
        self.radio_mapper_option_1.setObjectName("collectorModeRadio")
        self.radio_mapper_option_2 = QRadioButton("Option 2")
        self.radio_mapper_option_2.setObjectName("collectorModeRadio")
        self.radio_mapper_option_1.setChecked(True)

        self.mapper_option_group = QButtonGroup(page)
        self.mapper_option_group.setExclusive(True)
        self.mapper_option_group.addButton(self.radio_mapper_option_1)
        self.mapper_option_group.addButton(self.radio_mapper_option_2)

        layout.addWidget(self.radio_mapper_option_1)

        option1_row = QHBoxLayout()
        option1_row.setSpacing(12)
        self.btn_mapper_open_window = QPushButton("Open window")
        self.btn_mapper_open_window.setObjectName("collectorGrayBtn")
        option1_row.addWidget(self.btn_mapper_open_window, 0)
        option1_row.addStretch(1)
        layout.addLayout(option1_row)

        layout.addSpacing(8)

        sap_opt2_header = QHBoxLayout()
        sap_opt2_header.setSpacing(0)
        sap_opt2_header.addWidget(self.radio_mapper_option_2)
        sap_opt2_header.addSpacing(100)
        self.lbl_sap_input_count = QLabel("")
        self.lbl_sap_input_count.setObjectName("inputCountLabel")
        sap_opt2_header.addWidget(self.lbl_sap_input_count, 0)
        sap_opt2_header.addStretch(1)
        layout.addLayout(sap_opt2_header)

        sap_data_row = QHBoxLayout()
        sap_data_row.setSpacing(12)
        self.btn_mapper_sap_data_files = QPushButton("SAP Data Files")
        self.btn_mapper_sap_data_files.setObjectName("collectorGrayBtn")
        sap_data_row.addWidget(self.btn_mapper_sap_data_files, 0)
        self.input_mapper_sap_data_files = QLineEdit()
        self.input_mapper_sap_data_files.setObjectName("collectorLineEdit")
        self.input_mapper_sap_data_files.setReadOnly(True)
        sap_data_row.addWidget(self.input_mapper_sap_data_files, 1)
        layout.addLayout(sap_data_row)

        output_row = QHBoxLayout()
        output_row.setSpacing(12)
        self.btn_mapper_output_location = QPushButton("Output Location")
        self.btn_mapper_output_location.setObjectName("collectorGrayBtn")
        output_row.addWidget(self.btn_mapper_output_location, 0)
        self.input_mapper_output_location = QLineEdit()
        self.input_mapper_output_location.setObjectName("collectorLineEdit")
        self.input_mapper_output_location.setReadOnly(True)
        output_row.addWidget(self.input_mapper_output_location, 1)
        layout.addLayout(output_row)

        cleanup_radio_row = QHBoxLayout()
        cleanup_radio_row.setSpacing(16)

        self.check_mapper_cleanup_1 = QCheckBox("Cleanup 1")
        self.check_mapper_cleanup_1.setChecked(True)
        self.check_mapper_cleanup_1.setObjectName("mapperCleanupCheck")
        self.check_mapper_cleanup_1.setToolTip(
            "remove non-SMU and unneeded packaging\n"
            "(retains accl and flex SMU rows):\n"
            "sal, pal, film, ship, wgl, sheet, shee,\n"
            "pl, t-secur, saco, rbosac, acco_pe, tear, bulk"
        )
        cleanup_radio_row.addWidget(self.check_mapper_cleanup_1, 0)

        self.check_mapper_cleanup_2 = QCheckBox("Cleanup 2")
        self.check_mapper_cleanup_2.setChecked(True)
        self.check_mapper_cleanup_2.setObjectName("mapperCleanupCheck")
        self.check_mapper_cleanup_2.setToolTip(
            "remove non-SMU and unneeded packaging\n"
            "(including accl and flex):\n"
            "sal, flex, pall, accl, film, ship, wgl,\n"
            "sheet, shee, pl, t-secur, saco, acco_pe,\n"
            "tear, bulk"
        )
        cleanup_radio_row.addWidget(self.check_mapper_cleanup_2, 0)

        self.check_mapper_cleanup_3 = QCheckBox("Cleanup 3")
        self.check_mapper_cleanup_3.setChecked(True)
        self.check_mapper_cleanup_3.setObjectName("mapperCleanupCheck")
        self.check_mapper_cleanup_3.setToolTip(
            "remove non-SMU, un-identifiable basic name\n"
            "and unneeded packaging:\n"
            "sal, flex, pall, accl, film, ship, wgl,\n"
            "sheet, shee, pl, t-secur, saco, bag,\n"
            "rbosac, leaflet, acco, paco, tear, bulk"
        )
        cleanup_radio_row.addWidget(self.check_mapper_cleanup_3, 0)

        self.check_mapper_cleanup_4 = QCheckBox("Cleanup 4")
        self.check_mapper_cleanup_4.setChecked(False)
        self.check_mapper_cleanup_4.setObjectName("mapperCleanupCheck")
        self.check_mapper_cleanup_4.setToolTip("no cleanup, all info retained")
        cleanup_radio_row.addWidget(self.check_mapper_cleanup_4, 0)

        cleanup_radio_row.addStretch(1)
        layout.addLayout(cleanup_radio_row)

        self.checkbox_mapper_include_grouping_report = QCheckBox("Include Grouping Report")
        self.checkbox_mapper_include_grouping_report.setObjectName("collectorCheck")
        self.checkbox_mapper_include_grouping_report.setChecked(True)
        layout.addWidget(self.checkbox_mapper_include_grouping_report)

        layout.addStretch(1)
        run_row = QHBoxLayout()
        run_row.addStretch(1)
        self.btn_run_process_mapper_reformat = QPushButton("Run Process")
        self.btn_run_process_mapper_reformat.setObjectName("collectorRunBtn")
        self.btn_run_process_mapper_reformat.setFixedHeight(50)
        self.btn_run_process_mapper_reformat.setMinimumWidth(180)
        run_row.addWidget(self.btn_run_process_mapper_reformat)
        run_row.addStretch(1)
        layout.addLayout(run_row)
        layout.addSpacing(8)

        progress_row = QHBoxLayout()
        progress_row.addStretch(1)
        self.mapper_reformat_progress_bar = QProgressBar()
        self.mapper_reformat_progress_bar.setObjectName("collectorProgressBar")
        self.mapper_reformat_progress_bar.setMinimumWidth(260)
        self.mapper_reformat_progress_bar.setMaximumWidth(360)
        self.mapper_reformat_progress_bar.setFixedHeight(14)
        self.mapper_reformat_progress_bar.setTextVisible(False)
        self.mapper_reformat_progress_bar.setRange(0, 100)
        self.mapper_reformat_progress_bar.setValue(0)
        self.mapper_reformat_progress_bar.setVisible(False)
        progress_row.addWidget(self.mapper_reformat_progress_bar)
        progress_row.addStretch(1)
        layout.addLayout(progress_row)

        self.radio_mapper_option_1.toggled.connect(self._sync_mapper_reformat_mode_ui)
        self.radio_mapper_option_2.toggled.connect(self._sync_mapper_reformat_mode_ui)
        self.btn_mapper_open_window.clicked.connect(self._open_mapper_option1_table)
        self.btn_mapper_sap_data_files.clicked.connect(self._on_mapper_select_sap_files)
        self.btn_mapper_output_location.clicked.connect(self._on_mapper_select_output_location)
        self.btn_run_process_mapper_reformat.clicked.connect(self._on_mapper_run_process)
        self.input_mapper_sap_data_files.textChanged.connect(lambda: self._update_sap_input_count())
        self._sync_mapper_reformat_mode_ui()

        return page

    # ------------------------------------------------------------------
    # SAP Data Compare page
    # ------------------------------------------------------------------

    def _create_mapper_compare_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(7)

        title = QLabel("SAP Data Compare")
        title.setObjectName("mapperTitle")
        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.addWidget(title, 0)
        title_row.addSpacing(140)
        self.lbl_sdc_input_count = QLabel("")
        self.lbl_sdc_input_count.setObjectName("inputCountLabel")
        title_row.addWidget(self.lbl_sdc_input_count, 0)
        title_row.addStretch(1)
        layout.addLayout(title_row)

        # ── Set vs Individual Files mode ──────────────────────────────────
        sdc_mode_row = QHBoxLayout()
        sdc_mode_row.setSpacing(16)
        self.radio_sdc_set = QRadioButton("Set")
        self.radio_sdc_set.setObjectName("compareModeRadioSmall")
        self.radio_sdc_set.setChecked(True)
        self.radio_sdc_set.setToolTip(
            "Compare matching cu-level pairs (cu1 vs cu1, cu2 vs cu2, …)\n"
            "and merge results by best hit type."
        )
        self.radio_sdc_individual = QRadioButton("Individual Files")
        self.radio_sdc_individual.setObjectName("compareModeRadioSmall")
        self.radio_sdc_individual.setToolTip("Compare one RSD: Target file against one RSD: Master file.")
        self.sdc_input_mode_group = QButtonGroup(page)
        self.sdc_input_mode_group.setExclusive(True)
        self.sdc_input_mode_group.addButton(self.radio_sdc_set)
        self.sdc_input_mode_group.addButton(self.radio_sdc_individual)
        sdc_mode_row.addWidget(self.radio_sdc_set, 0)
        sdc_mode_row.addWidget(self.radio_sdc_individual, 0)
        sdc_mode_row.addStretch(1)
        layout.addLayout(sdc_mode_row)

        def _make_pair(btn_label: str):
            row = QHBoxLayout()
            row.setSpacing(8)
            btn = QPushButton(btn_label)
            btn.setObjectName("compareGrayBtn")
            btn.setFixedWidth(160)
            row.addWidget(btn, 0)
            lbl = QLineEdit()
            lbl.setObjectName("comparePathLabel")
            lbl.setFixedHeight(34)
            row.addWidget(lbl, 1)
            layout.addLayout(row)
            return btn, lbl

        self.btn_compare_rsd_target, self.lbl_compare_rsd_target = _make_pair("RSD: Target")

        layout.addSpacing(12)

        master_row = QHBoxLayout()
        master_row.setSpacing(0)
        self.checkbox_compare_with_master = QCheckBox("Compare with Master Data")
        self.checkbox_compare_with_master.setObjectName("collectorCheckSmall")
        self.checkbox_compare_with_master.setChecked(True)
        master_row.addWidget(self.checkbox_compare_with_master, 0)
        master_row.addSpacing(40)
        self.radio_compare_with_rsd_master = QRadioButton("Compare with RSD: Master")
        self.radio_compare_with_rsd_master.setObjectName("compareModeRadioSmall")
        self.radio_compare_with_rsd_master.setChecked(True)
        master_row.addWidget(self.radio_compare_with_rsd_master, 0)
        master_row.addSpacing(16)
        self.radio_compare_only_tsc = QRadioButton("Compare only with TSC")
        self.radio_compare_only_tsc.setObjectName("compareModeRadioSmall")
        master_row.addWidget(self.radio_compare_only_tsc, 0)
        master_row.addStretch(1)
        self.compare_master_mode_group = QButtonGroup(page)
        self.compare_master_mode_group.setExclusive(True)
        self.compare_master_mode_group.addButton(self.radio_compare_with_rsd_master)
        self.compare_master_mode_group.addButton(self.radio_compare_only_tsc)
        layout.addLayout(master_row)

        self.btn_compare_rsd_master, self.lbl_compare_rsd_master = _make_pair("RSD: Master")
        self.btn_compare_tsc_data, self.lbl_compare_tsc_data = _make_pair("TSC Data")

        # ── Compare with Library ──────────────────────────────────────────
        layout.addSpacing(10)
        self.checkbox_compare_with_library = QCheckBox("Compare with Library")
        self.checkbox_compare_with_library.setObjectName("collectorCheckSmall")
        self.checkbox_compare_with_library.setChecked(True)
        layout.addWidget(self.checkbox_compare_with_library)
        self.btn_compare_library, self.lbl_compare_library = _make_pair("Excel Library")

        # ── Run reference collector ───────────────────────────────────────
        layout.addSpacing(10)
        self.checkbox_compare_run_ref_collector = QCheckBox("Run reference collector")
        self.checkbox_compare_run_ref_collector.setObjectName("collectorCheckSmall")
        self.checkbox_compare_run_ref_collector.setChecked(True)
        layout.addWidget(self.checkbox_compare_run_ref_collector)
        self.btn_compare_packshot_location, self.lbl_compare_packshot_location = _make_pair("Packshot Image Library")

        max_packshot_row = QHBoxLayout()
        max_packshot_row.setSpacing(8)
        self.lbl_max_packshot = QLabel("Maximum count:")
        self.lbl_max_packshot.setObjectName("compareSmallLabel")
        max_packshot_row.addWidget(self.lbl_max_packshot, 0)
        self.input_compare_max_packshot = QLineEdit("5")
        self.input_compare_max_packshot.setObjectName("compareSmallInput")
        self.input_compare_max_packshot.setFixedWidth(50)
        self.input_compare_max_packshot.setFixedHeight(28)
        max_packshot_row.addWidget(self.input_compare_max_packshot, 0)
        max_packshot_row.addStretch(1)
        layout.addLayout(max_packshot_row)

        # ── Output Location ───────────────────────────────────────────────
        self.btn_compare_output_location, self.lbl_compare_output_location = _make_pair("Output Location")

        layout.addStretch(1)

        run_row = QHBoxLayout()
        run_row.addStretch(1)
        self.btn_run_process_mapper_compare = QPushButton("Run Process")
        self.btn_run_process_mapper_compare.setObjectName("compareRunBtn")
        self.btn_run_process_mapper_compare.setFixedHeight(35)
        self.btn_run_process_mapper_compare.setMinimumWidth(126)
        run_row.addWidget(self.btn_run_process_mapper_compare)
        run_row.addStretch(1)
        layout.addLayout(run_row)
        layout.addSpacing(4)

        progress_row = QHBoxLayout()
        progress_row.addStretch(1)
        self.mapper_compare_progress_bar = QProgressBar()
        self.mapper_compare_progress_bar.setObjectName("compareProgressBar")
        self.mapper_compare_progress_bar.setMinimumWidth(180)
        self.mapper_compare_progress_bar.setMaximumWidth(260)
        self.mapper_compare_progress_bar.setFixedHeight(10)
        self.mapper_compare_progress_bar.setTextVisible(False)
        self.mapper_compare_progress_bar.setRange(0, 100)
        self.mapper_compare_progress_bar.setValue(0)
        self.mapper_compare_progress_bar.setVisible(False)
        progress_row.addWidget(self.mapper_compare_progress_bar)
        progress_row.addStretch(1)
        layout.addLayout(progress_row)

        self.radio_sdc_set.toggled.connect(self._sync_mapper_compare_mode_ui)
        self.radio_sdc_individual.toggled.connect(self._sync_mapper_compare_mode_ui)
        self.checkbox_compare_with_master.stateChanged.connect(self._sync_mapper_compare_mode_ui)
        self.radio_compare_with_rsd_master.toggled.connect(self._sync_mapper_compare_mode_ui)
        self.radio_compare_only_tsc.toggled.connect(self._sync_mapper_compare_mode_ui)
        self.checkbox_compare_with_library.stateChanged.connect(self._sync_mapper_compare_mode_ui)
        self.checkbox_compare_run_ref_collector.stateChanged.connect(self._sync_mapper_compare_mode_ui)
        self._sync_mapper_compare_mode_ui()

        self.lbl_compare_rsd_target.textChanged.connect(lambda: self._update_sdc_input_count())
        self.btn_compare_rsd_target.clicked.connect(self._on_compare_browse_rsd_target)
        self.btn_compare_rsd_master.clicked.connect(self._on_compare_browse_rsd_master)
        self.btn_compare_tsc_data.clicked.connect(self._on_compare_browse_tsc_data)
        self.btn_compare_library.clicked.connect(self._on_compare_browse_library)
        self.btn_compare_packshot_location.clicked.connect(self._on_compare_browse_packshot_location)
        self.btn_compare_output_location.clicked.connect(self._on_compare_browse_output_location)
        self.btn_run_process_mapper_compare.clicked.connect(self._on_mapper_compare_run)

        return page

    def _on_mapper_compare_run(self) -> None:
        """Dispatch to set-mode or individual-mode comparison."""
        if hasattr(self, "radio_sdc_set") and self.radio_sdc_set.isChecked():
            self._on_mapper_compare_run_set()
        else:
            self._on_mapper_compare_run_individual()

    def _extract_cu_level(self, file_path: str) -> "int | None":
        """Return the cu level (1–4) from an RSD filename, or None if not found."""
        import re as _re
        m = _re.search(r"[_\-]cu(\d+)[_\-]", Path(file_path).stem, _re.IGNORECASE)
        return int(m.group(1)) if m else None

    def _on_mapper_compare_run_set(self) -> None:
        """Set-mode: pair files by cu level, merge results by best hit type."""
        from pathlib import Path as _Path

        raw_target  = self.lbl_compare_rsd_target.text().strip()
        raw_master  = self.lbl_compare_rsd_master.text().strip()
        tsc_data_path = self.lbl_compare_tsc_data.text().strip()
        output_dir  = self.lbl_compare_output_location.text().strip()

        # ── basic validation ─────────────────────────────────────────────
        missing: list[str] = []
        compare_mode = "tsc_only" if self.radio_compare_only_tsc.isChecked() else "rsd_master"
        if not raw_target:
            missing.append("RSD: Target")
        if compare_mode == "rsd_master" and not raw_master:
            missing.append("RSD: Master")
        if not tsc_data_path:
            missing.append("TSC Data")
        if not output_dir:
            missing.append("Output Location")
        if missing:
            QMessageBox.warning(
                self, "Missing Input",
                "Please fill in the following fields:\n" + "\n".join(f"  • {m}" for m in missing),
            )
            return

        # ── parse cu levels ───────────────────────────────────────────────
        target_paths = [p.strip() for p in raw_target.split(",") if p.strip()]
        master_paths = [p.strip() for p in raw_master.split(",") if p.strip()] if raw_master else []

        target_cu: dict[int, str] = {}
        for p in target_paths:
            lvl = self._extract_cu_level(p)
            if lvl is not None:
                target_cu[lvl] = p

        master_cu: dict[int, str] = {}
        if compare_mode == "rsd_master":
            for p in master_paths:
                lvl = self._extract_cu_level(p)
                if lvl is not None:
                    master_cu[lvl] = p

        if not target_cu:
            QMessageBox.warning(
                self, "No cu-tagged Files",
                "RSD: Target files must include a cu-level tag in their filename (e.g. rsd_name_cu1_...).\n"
                "Please select the correct files.",
            )
            return

        if compare_mode == "rsd_master" and not master_cu:
            QMessageBox.warning(
                self, "No cu-tagged Files",
                "RSD: Master files must include a cu-level tag in their filename (e.g. rsd_name_cu1_...).\n"
                "Please select the correct files.",
            )
            return

        common_levels = sorted(
            set(target_cu.keys()) & (set(master_cu.keys()) if compare_mode == "rsd_master" else set(target_cu.keys()))
        )
        if not common_levels:
            QMessageBox.warning(
                self, "No Matching Levels",
                "No matching cu levels found between RSD: Target and RSD: Master.\n"
                "Ensure both sides have files tagged with the same cu level (cu1, cu2, etc.).",
            )
            return

        # ── mismatch dialog ───────────────────────────────────────────────
        target_levels = sorted(target_cu.keys())
        master_levels = sorted(master_cu.keys()) if compare_mode == "rsd_master" else target_levels
        if target_levels != master_levels:
            dlg = _SetMismatchDialog(target_levels, master_levels, common_levels, self)
            if dlg.exec() != QDialog.Accepted:
                return

        # ── library ───────────────────────────────────────────────────────
        _LIB_ERROR_STYLE = (
            "background-color: rgba(208, 39, 82, 128); color: #000000; "
            "border: 1px solid #6F6F6F; border-radius: 11px; "
            "padding: 0 8px; font-family: 'Segoe UI'; font-size: 10px;"
        )
        excel_library_path = ""
        if self.checkbox_compare_with_library.isChecked():
            lib_text = self.lbl_compare_library.text().strip()
            lib_paths = [p.strip() for p in lib_text.split(",") if p.strip()]
            lib_is_single_file = (
                len(lib_paths) == 1
                and lib_paths[0].lower().endswith(".xlsx")
                and _Path(lib_paths[0]).is_file()
            )
            if lib_is_single_file:
                excel_library_path = lib_paths[0]

        # ── derive base name from first target file ───────────────────────
        import re as _re
        _DATE_SUFFIX = _re.compile(
            r"[_-]\d{4}[_-]\d{2}[_-]\d{2}[_-]\d{2}[_-]\d{2}$"
            r"|[_-]\d{2}[_-]\d{2}[_-]\d{2}[_-]\d{2}$"
        )
        _CU_TAG = _re.compile(r"[_\-]cu\d+[_\-]?", _re.IGNORECASE)

        def _set_base_name(path: str) -> str:
            stem = _Path(path).stem
            if stem.lower().startswith("rsd_"):
                stem = stem[4:]
            stem = _CU_TAG.sub("_", stem)
            return _DATE_SUFFIX.sub("", stem).strip("_")

        base_name = _set_base_name(list(target_cu.values())[0])

        # ── run ───────────────────────────────────────────────────────────
        self.mapper_compare_progress_bar.setValue(0)
        self.mapper_compare_progress_bar.setVisible(True)
        self.btn_run_process_mapper_compare.setEnabled(False)
        QApplication.processEvents()

        compare_result = run_set_comparison(
            target_cu_paths=target_cu,
            master_cu_paths=master_cu if compare_mode == "rsd_master" else {},
            common_levels=common_levels,
            tsc_data_path=tsc_data_path,
            output_dir=output_dir,
            base_name=base_name,
            excel_library_path=excel_library_path,
            compare_mode=compare_mode,
        )

        self.mapper_compare_progress_bar.setValue(80)
        QApplication.processEvents()

        # ── reference collector (optional) ────────────────────────────────
        rc_result = None
        if (
            self.checkbox_compare_run_ref_collector.isChecked()
            and compare_result.output_paths
        ):
            packshot_folder = self.lbl_compare_packshot_location.text().strip()
            if not packshot_folder or not Path(packshot_folder).is_dir():
                compare_result.warnings.append(
                    "Reference collector skipped: Packshot Image Library folder is missing or invalid."
                )
            else:
                try:
                    max_imgs = int(self.input_compare_max_packshot.text().strip())
                except ValueError:
                    max_imgs = 5
                rc_params = RefCollectorParams(
                    sdc_output_paths   = compare_result.output_paths,
                    image_library_path = packshot_folder,
                    output_dir         = output_dir,
                    tsc_data_path      = tsc_data_path,
                    max_images         = max(1, max_imgs),
                )
                rc_result = run_reference_collector(rc_params)
                compare_result.warnings.extend(rc_result.warnings)

        self.mapper_compare_progress_bar.setValue(100)
        QApplication.processEvents()
        self.mapper_compare_progress_bar.setVisible(False)
        self.btn_run_process_mapper_compare.setEnabled(True)

        if compare_result.warnings:
            warn_text = "\n".join(compare_result.warnings)
            if not compare_result.output_paths:
                QMessageBox.critical(self, "SDC Set Error", warn_text)
                return
            QMessageBox.warning(self, "SDC Set completed with warnings", warn_text)

        if compare_result.output_paths:
            paths_text = "\n".join(compare_result.output_paths)
            rc_info = ""
            if rc_result and rc_result.pdf_paths:
                rc_info = (
                    f"\n\nReference Collector PDFs ({len(rc_result.pdf_paths)}) "
                    f"saved to:\n{output_dir}"
                )
            QMessageBox.information(self, "SDC Set Complete", f"Output saved to:\n{paths_text}{rc_info}")

    def _on_mapper_compare_run_individual(self) -> None:
        """Individual-mode: existing single-target comparison logic."""

        _ERROR_STYLE = (
            "background-color: rgba(208, 39, 82, 128); color: #000000; "
            "border: 1px solid #6F6F6F; border-radius: 11px; "
            "padding: 0 8px; font-family: 'Segoe UI'; font-size: 10px;"
        )

        # ── gather RSD: Target paths ─────────────────────────────────────
        raw_target = self.lbl_compare_rsd_target.text().strip()
        rsd_target_paths = [p.strip() for p in raw_target.split(",") if p.strip()] if raw_target else []

        raw_master = self.lbl_compare_rsd_master.text().strip()
        tsc_data_path = self.lbl_compare_tsc_data.text().strip()
        output_dir    = self.lbl_compare_output_location.text().strip()

        # ── validate single master file ──────────────────────────────────
        compare_mode = "tsc_only" if self.radio_compare_only_tsc.isChecked() else "rsd_master"
        if compare_mode == "rsd_master":
            master_parts = [p.strip() for p in raw_master.split(",") if p.strip()]
            if len(master_parts) > 1:
                self.lbl_compare_rsd_master.setStyleSheet(_ERROR_STYLE)
                self.lbl_compare_rsd_master.setText("Only 1 file allowed in Individual mode")
                return
            # Ensure any stale error style is cleared when there is exactly 1 file
            if master_parts:
                self.lbl_compare_rsd_master.setStyleSheet("")
        rsd_master_path = raw_master

        # ── library validation ───────────────────────────────────────────
        excel_library_path = ""
        if self.checkbox_compare_with_library.isChecked():
            lib_text = self.lbl_compare_library.text().strip()
            lib_paths = [p.strip() for p in lib_text.split(",") if p.strip()]
            lib_is_single_file = (
                len(lib_paths) == 1
                and lib_paths[0].lower().endswith(".xlsx")
                and Path(lib_paths[0]).is_file()
            )
            if len(lib_paths) > 1 or (lib_text and not lib_is_single_file and "multiple" in lib_text.lower()):
                self.lbl_compare_library.setStyleSheet(_ERROR_STYLE)
                self.lbl_compare_library.setText("multiple excel files in library folder")
                return
            elif not lib_is_single_file:
                # no valid library file – warn but continue without library
                pass
            else:
                excel_library_path = lib_paths[0]

        # ── basic validation ─────────────────────────────────────────────
        missing: list[str] = []
        if not rsd_target_paths:
            missing.append("RSD: Target")
        if compare_mode == "rsd_master" and not rsd_master_path:
            missing.append("RSD: Master")
        if not tsc_data_path:
            missing.append("TSC Data")
        if not output_dir:
            missing.append("Output Location")
        if missing:
            QMessageBox.warning(
                self, "Missing Input",
                "Please fill in the following fields:\n" + "\n".join(f"  • {m}" for m in missing)
            )
            return

        # ── run ──────────────────────────────────────────────────────────
        self.mapper_compare_progress_bar.setValue(0)
        self.mapper_compare_progress_bar.setVisible(True)
        self.btn_run_process_mapper_compare.setEnabled(False)
        QApplication.processEvents()

        params = CompareParams(
            rsd_target_paths=rsd_target_paths,
            rsd_master_path=rsd_master_path,
            tsc_data_path=tsc_data_path,
            output_dir=output_dir,
            excel_library_path=excel_library_path,
            compare_mode=compare_mode,
        )

        self.mapper_compare_progress_bar.setValue(30)
        QApplication.processEvents()

        compare_result = run_comparison(params)

        # ── reference collector (optional) ──────────────────────────────────
        rc_result = None
        if (
            self.checkbox_compare_run_ref_collector.isChecked()
            and compare_result.output_paths
        ):
            packshot_folder = self.lbl_compare_packshot_location.text().strip()
            if not packshot_folder or not Path(packshot_folder).is_dir():
                compare_result.warnings.append(
                    "Reference collector skipped: Packshot Image Library folder "
                    "is missing or invalid."
                )
            else:
                try:
                    max_imgs = int(self.input_compare_max_packshot.text().strip())
                except ValueError:
                    max_imgs = 5

                # Determine which SDC files to process
                sdc_paths_for_rc: list[str] = []
                if len(compare_result.output_paths) == 1:
                    sdc_paths_for_rc = compare_result.output_paths
                else:
                    dlg = _SdcSelectionDialog(compare_result.output_paths, parent=self)
                    if dlg.exec() == QDialog.Accepted and dlg.selected_paths:
                        sdc_paths_for_rc = dlg.selected_paths
                    # If cancelled or nothing checked → skip RC silently

                if sdc_paths_for_rc:
                    rc_params = RefCollectorParams(
                        sdc_output_paths   = sdc_paths_for_rc,
                        image_library_path = packshot_folder,
                        output_dir         = output_dir,
                        tsc_data_path      = tsc_data_path,
                        max_images         = max(1, max_imgs),
                    )
                    rc_result = run_reference_collector(rc_params)
                    compare_result.warnings.extend(rc_result.warnings)

        self.mapper_compare_progress_bar.setValue(100)
        QApplication.processEvents()
        self.mapper_compare_progress_bar.setVisible(False)
        self.btn_run_process_mapper_compare.setEnabled(True)

        if compare_result.warnings:
            warn_text = "\n".join(compare_result.warnings)
            if not compare_result.output_paths:
                QMessageBox.critical(self, "SDC Error", warn_text)
                return
            QMessageBox.warning(self, "SDC completed with warnings", warn_text)

        if compare_result.output_paths:
            paths_text = "\n".join(compare_result.output_paths)
            rc_info = ""
            if rc_result and rc_result.pdf_paths:
                folders_text = "\n".join(rc_result.output_folders)
                rc_info = (
                    f"\n\nReference Collector PDFs ({len(rc_result.pdf_paths)}) "
                    f"in {len(rc_result.output_folders)} folder(s):"
                    f"\n{folders_text}"
                )
            elif rc_result and not rc_result.pdf_paths:
                rc_info = "\n\nReference Collector: no PDFs generated."
            QMessageBox.information(
                self, "SDC Complete",
                f"Output saved to:\n{paths_text}{rc_info}"
            )

    def _on_compare_browse_rsd_target(self) -> None:
        start = self._get_browse_dir("sap_compare")
        if self._use_root_folders:
            text = self.lbl_compare_rsd_target.text().strip()
            if text:
                from pathlib import Path as _P
                _first = _P(text.split(",")[0].strip())
                if _first.is_absolute() and _first.parent.is_dir():
                    start = str(_first.parent)
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select RSD: Target File(s)", start, "Excel Files (*.xlsx)")
        if paths:
            self.lbl_compare_rsd_target.setText(", ".join(paths))

    def _on_compare_browse_rsd_master(self) -> None:
        start = self._get_browse_dir("sap_compare_master")
        if self._use_root_folders:
            text = self.lbl_compare_rsd_master.text().strip()
            if text:
                from pathlib import Path as _P
                _p = _P(text.split(",")[0].strip())
                if _p.is_absolute() and _p.parent.is_dir():
                    start = str(_p.parent)
        # In "Set" mode allow multi-file selection; in "Individual" use single file picker
        if hasattr(self, "radio_sdc_set") and self.radio_sdc_set.isChecked():
            paths, _ = QFileDialog.getOpenFileNames(
                self, "Select RSD: Master File(s)", start, "Excel Files (*.xlsx *.xls)")
            if paths:
                self.lbl_compare_rsd_master.setStyleSheet("")
                self.lbl_compare_rsd_master.setText(", ".join(paths))
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "Select RSD: Master File", start, "Excel Files (*.xlsx *.xls)")
            if path:
                self.lbl_compare_rsd_master.setStyleSheet("")
                self.lbl_compare_rsd_master.setText(path)

    def _on_compare_browse_tsc_data(self) -> None:
        start = self._get_browse_dir("sap_compare")
        if self._use_root_folders:
            text = self.lbl_compare_tsc_data.text().strip()
            if text:
                from pathlib import Path as _P
                _p = _P(text)
                if _p.is_absolute() and _p.parent.is_dir():
                    start = str(_p.parent)
        path, _ = QFileDialog.getOpenFileName(
            self, "Select TSC Data File", start, "Excel Files (*.xlsx *.xls)")
        if path:
            self.lbl_compare_tsc_data.setText(path)

    def _on_compare_browse_library(self) -> None:
        start = self._get_browse_dir("excel_library")
        if self._use_root_folders:
            text = self.lbl_compare_library.text().strip()
            if text:
                from pathlib import Path as _P
                _p = _P(text)
                if _p.is_absolute() and _p.parent.is_dir():
                    start = str(_p.parent)
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel Library File", start, "Excel Files (*.xlsx)")
        if path:
            self.lbl_compare_library.setText(path)

    def _on_compare_browse_packshot_location(self) -> None:
        start = self._get_browse_dir("sap_compare")
        if self._use_root_folders:
            text = self.lbl_compare_packshot_location.text().strip()
            if text:
                from pathlib import Path as _P
                _p = _P(text)
                if _p.is_absolute() and _p.is_dir():
                    start = text
        folder = QFileDialog.getExistingDirectory(
            self, "Select Packshot Location", start)
        if folder:
            self.lbl_compare_packshot_location.setText(folder)

    def _on_compare_browse_output_location(self) -> None:
        start = self._get_browse_dir("sap_compare")
        if self._use_root_folders:
            text = self.lbl_compare_output_location.text().strip()
            if text:
                from pathlib import Path as _P
                _p = _P(text)
                if _p.is_absolute() and _p.is_dir():
                    start = text
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Location", start)
        if folder:
            self.lbl_compare_output_location.setText(folder)

    def _on_root_folders_toggle(self, state: bool) -> None:
        """Validate config, update shared flag, then auto-fill or clear inputs."""
        if state and not self._config.is_root_configured():
            self._use_root_folders = False
            if self.btn_rail_toggle.isChecked():
                self.btn_rail_toggle.setChecked(False)
            QMessageBox.warning(
                self,
                "Root folders not configured",
                "No root folders configured.\n\nGo to Settings (\u2699) and set up root folders first."
            )
            return

        self._use_root_folders = state

        if not state:
            self._clear_root_folder_inputs()
            return

        self._apply_root_folder_inputs()

    def _clear_root_folder_inputs(self) -> None:
        """Clear all fields that were auto-filled by the root folder toggle."""
        for field in (
            self.input_trackers,
            self.input_output,
            self.input_pg7_images_folder,
            self.input_pg7_output,
            self.input_pg6_excel_tracker,
            self.input_pg6_output,
            self.input_mapper_sap_data_files,
            self.input_mapper_output_location,
        ):
            field.clear()
        if hasattr(self, "input_pr_source_bma"):
            self.input_pr_source_bma.clear()
        if hasattr(self, "input_pr_project_tracker"):
            self.input_pr_project_tracker.clear()
        self._update_tsc_input_count()
        self._update_png_input_count()
        self._update_sap_input_count()
        for lbl in (
            self.lbl_compare_rsd_target,
            self.lbl_compare_rsd_master,
            self.lbl_compare_tsc_data,
            self.lbl_compare_library,
            self.lbl_compare_packshot_location,
            self.lbl_compare_output_location,
        ):
            lbl.setText("")
            lbl.setStyleSheet("")

    def _apply_root_folder_inputs(self) -> None:
        """Route autofill to MASTER or PROJECT depending on the active nav page."""
        is_project = (self.left_nav_stack.currentWidget() is self.project_left)
        if is_project:
            self._autofill_project_page()
        else:
            self._autofill_master_page()

    def _autofill_master_page(self) -> None:
        """Fill MASTER tool input fields from configured root folder paths."""
        from pathlib import Path

        # ── TSC ───────────────────────────────────────────────────────────
        bt_folder = self._config.briefing_tracker_input()
        if bt_folder:
            p = Path(bt_folder)
            if p.is_dir():
                files = sorted(f for f in p.glob("*.xlsx"))
                if files:
                    file_strs = [str(f) for f in files]
                    self.input_trackers.setText(", ".join(file_strs))
                    if hasattr(self, "status_collector"):
                        sc = self.status_collector
                        sc.excel_trackers = file_strs
                        sc.excel_tracker_names = [f.name for f in files]
                        sc.editText_selected_trackers.setPlainText(
                            ",\n".join(f.name for f in files)
                        )
        tsc_out = self._config.tsc_output()
        if tsc_out:
            self.input_output.setText(tsc_out)
            if hasattr(self, "status_collector"):
                self.status_collector.output_location = tsc_out
                self.status_collector.editText_output_location.setPlainText(tsc_out)
        self._update_tsc_input_count()

        # ── Thumbnail Generator ───────────────────────────────────────────
        tg_folder = self._config.thumbnail_input()
        if tg_folder:
            self.input_pg7_images_folder.setText(tg_folder)
            self.input_pg7_output.setText(tg_folder)

        # ── SAP Data Reformat (MASTER) ────────────────────────────────────
        sdr_folder = self._config.sap_reformat_input()
        if sdr_folder:
            p = Path(sdr_folder)
            if p.is_dir():
                files = sorted(
                    f for f in p.iterdir()
                    if f.suffix.lower() == ".xls"
                    and "rsd" not in f.stem.lower()
                )
                if files:
                    file_strs = [str(f) for f in files]
                    self.input_mapper_sap_data_files.setText(", ".join(file_strs))
                    self._mapper_sap_file_paths = file_strs
            self.input_mapper_output_location.setText(sdr_folder)
            self._mapper_output_location = sdr_folder
        self._update_sap_input_count()

    def _autofill_project_page(self) -> None:
        """Fill PROJECT tool input fields based on the selected project folder."""
        from pathlib import Path
        project_folder = self._get_active_project_folder()
        if not project_folder:
            return

        # ── Packshot Naming Generator ──────────────────────────────────────
        png_folder = str(Path(project_folder) / "Packshot Naming Generator")
        p_png = Path(png_folder)
        if p_png.is_dir():
            files = sorted(
                f for f in p_png.glob("*.xlsx")
                if "packshot_naming" not in f.stem.lower()
            )
            if files:
                first = files[0]
                self.input_pg6_excel_tracker.setText(str(first))
                if hasattr(self, "packshot_naming_generator"):
                    png = self.packshot_naming_generator
                    png.excel_values.out_excel_file = str(first)
                    png.excel_values.out_excel_file_name = first.name
                    png.excel_values.check_tracker_loaded = True
            self.input_pg6_output.setText(png_folder)
        self._update_png_input_count()

        # ── SAP Data Reformat (PROJECT) ───────────────────────────────────
        sdr_folder = str(Path(project_folder) / "SAP Data Reformat")
        p_sdr = Path(sdr_folder)
        if p_sdr.is_dir():
            files = sorted(
                f for f in p_sdr.iterdir()
                if f.suffix.lower() == ".xls"
                and "rsd" not in f.stem.lower()
            )
            if files:
                file_strs = [str(f) for f in files]
                self.input_mapper_sap_data_files.setText(", ".join(file_strs))
                self._mapper_sap_file_paths = file_strs
            self.input_mapper_output_location.setText(sdr_folder)
            self._mapper_output_location = sdr_folder
        self._update_sap_input_count()

        # ── SAP Data Compare (PROJECT) ────────────────────────────────────
        # RSD: Target — project SAP Data Reformat folder, all .xlsx files with "rsd" in name
        p_rsd_target = Path(sdr_folder)
        if p_rsd_target.is_dir():
            rsd_target_files = sorted(
                f for f in p_rsd_target.iterdir()
                if f.suffix.lower() == ".xlsx"
                and "rsd" in f.stem.lower()
            )
            if rsd_target_files:
                self.lbl_compare_rsd_target.setText(", ".join(str(f) for f in rsd_target_files))

        # RSD: Master — [sap_compare] rsd_master_folder → sap_reformat_input() → Desktop default
        master_rsd_folder = self._config.rsd_master_folder()
        if not master_rsd_folder:
            master_rsd_folder = str(Path.home() / "Desktop" / "HAT DASHBOARD ROOT" / "MASTER" / "SAP Data Reformat")
        if master_rsd_folder:
            p_master_rsd = Path(master_rsd_folder)
            if p_master_rsd.is_dir():
                master_rsd_files = sorted(
                    f for f in p_master_rsd.iterdir()
                    if f.suffix.lower() == ".xlsx"
                    and "rsd" in f.stem.lower()
                )
                _is_set = hasattr(self, "radio_sdc_set") and self.radio_sdc_set.isChecked()
                if _is_set:
                    # Set mode: fill all matching rsd_cu* files
                    if master_rsd_files:
                        self.lbl_compare_rsd_master.setStyleSheet("")
                        self.lbl_compare_rsd_master.setText(", ".join(str(f) for f in master_rsd_files))
                else:
                    # Individual mode: expect exactly one file
                    if len(master_rsd_files) == 1:
                        self.lbl_compare_rsd_master.setStyleSheet("")
                        self.lbl_compare_rsd_master.setText(str(master_rsd_files[0]))
                    elif len(master_rsd_files) > 1:
                        self.lbl_compare_rsd_master.setStyleSheet(
                            "background-color: rgba(208, 39, 82, 128); color: #111111; "
                            "border: 1px solid #6F6F6F; border-radius: 11px; "
                            "padding: 0 8px; font-family: 'Segoe UI'; font-size: 10px;"
                        )
                        self.lbl_compare_rsd_master.setText("Multiple rsd files found in master")

        # TSC Data — MASTER Tracker Status Collector folder, single .xlsx file with "tsc" in name
        tsc_folder = self._config.tsc_output()
        if tsc_folder:
            p_tsc = Path(tsc_folder)
            if p_tsc.is_dir():
                tsc_files = sorted(
                    f for f in p_tsc.iterdir()
                    if f.suffix.lower() == ".xlsx"
                    and "tsc" in f.stem.lower()
                )
                if len(tsc_files) == 1:
                    self.lbl_compare_tsc_data.setStyleSheet("")
                    self.lbl_compare_tsc_data.setText(str(tsc_files[0]))
                elif len(tsc_files) > 1:
                    self.lbl_compare_tsc_data.setStyleSheet("color: rgba(208, 39, 82, 128);")
                    self.lbl_compare_tsc_data.setText("Multiple tsc files found in Master")

        # Excel Library — INI override or MASTER/Excel Library folder
        excel_lib_folder = self._config.excel_library_folder()
        if excel_lib_folder:
            p_lib = Path(excel_lib_folder)
            if p_lib.is_dir():
                lib_files = sorted(f for f in p_lib.glob("*.xlsx"))
                if len(lib_files) == 1:
                    self.lbl_compare_library.setStyleSheet("")
                    self.lbl_compare_library.setText(str(lib_files[0]))
                elif len(lib_files) > 1:
                    self.lbl_compare_library.setStyleSheet(
                        "background-color: rgba(208, 39, 82, 128); color: #000000; "
                        "border: 1px solid #6F6F6F; border-radius: 11px; "
                        "padding: 0 8px; font-family: 'Segoe UI'; font-size: 10px;"
                    )
                    self.lbl_compare_library.setText("multiple excel files in library folder")

        # Packshot Image Library — MASTER Thumbnail Generator folder
        tg_folder = self._config.thumbnail_input()
        if tg_folder and Path(tg_folder).is_dir():
            self.lbl_compare_packshot_location.setText(tg_folder)

        # Output Location — project SAP Data Compare folder
        sdc_out = str(Path(project_folder) / "SAP Data Compare")
        if Path(sdc_out).is_dir():
            self.lbl_compare_output_location.setText(sdc_out)

        # ── Project Review — Source BMA ────────────────────────────────────
        if hasattr(self, "input_pr_source_bma"):
            # New Review mode — link from SAP Data Compare ("sdc" in filename)
            # Continue from Existing mode — link from Project Review ("review" in filename)
            if self.radio_pr_continue.isChecked():
                pr_folder = Path(project_folder) / "Project Review"
                if pr_folder.is_dir():
                    review_files = sorted(
                        f for f in pr_folder.iterdir()
                        if f.suffix.lower() in (".xlsx", ".xls", ".xlsm")
                        and "review" in f.stem.lower()
                    )
                    if review_files:
                        chosen = review_files[0] if len(review_files) == 1 else self._pick_latest_review_file(review_files)
                        self.input_pr_source_bma.setText(str(chosen))
                    else:
                        self.input_pr_source_bma.clear()
                else:
                    self.input_pr_source_bma.clear()
            else:  # New Review
                sdc_folder = Path(project_folder) / "SAP Data Compare"
                if sdc_folder.is_dir():
                    sdc_files = sorted(
                        f for f in sdc_folder.iterdir()
                        if f.suffix.lower() in (".xlsx", ".xls", ".xlsm")
                        and "sdc" in f.stem.lower()
                    )
                    if sdc_files:
                        self.input_pr_source_bma.setText(", ".join(str(f) for f in sdc_files))
                    else:
                        self.input_pr_source_bma.clear()
                else:
                    self.input_pr_source_bma.clear()

        # ── Project Review — Project Tracker ───────────────────────────────
        if hasattr(self, "input_pr_project_tracker"):
            self._autofill_pr_tracker(project_folder)

    def _autofill_pr_tracker(self, project_folder: str) -> None:
        """Scan the Project Review folder for Briefing Tracker file(s) and fill the input."""
        if not hasattr(self, "input_pr_project_tracker"):
            return
        tracker_files = review_project.scan_tracker_files(project_folder)
        if tracker_files:
            self.input_pr_project_tracker.setText(", ".join(str(f) for f in tracker_files))
        else:
            self.input_pr_project_tracker.clear()

    def _get_active_project_folder(self) -> str:
        """Return full path to the currently selected project folder, or '' if unavailable."""
        from pathlib import Path
        root = self._config.root_folder()
        if not root:
            return ""
        project_name = self.combo_project.currentText().strip()
        if not project_name or project_name.startswith("("):
            return ""
        p = Path(root) / "HAT DASHBOARD ROOT" / project_name
        return str(p) if p.is_dir() else ""

    def _update_sdc_input_count(self) -> None:
        text = self.lbl_compare_rsd_target.text().strip()
        if text:
            count = len([p for p in text.split(",") if p.strip()])
        else:
            count = 0
        lbl = getattr(self, "lbl_sdc_input_count", None)
        if lbl is not None:
            lbl.setText(f"Input files count: {count}" if count > 0 else "")

    def _update_tsc_input_count(self) -> None:
        text = self.input_trackers.text().strip()
        if text:
            count = len([p for p in text.split(",") if p.strip()])
        else:
            count = 0
        lbl = getattr(self, "lbl_tsc_input_count", None)
        if lbl is not None:
            lbl.setText(f"Input files count: {count}" if count > 0 else "")

    def _update_png_input_count(self) -> None:
        text = self.input_pg6_excel_tracker.text().strip()
        count = 1 if text else 0
        lbl = getattr(self, "lbl_png_input_count", None)
        if lbl is not None:
            lbl.setText(f"Input files count: {count}" if count > 0 else "")

    def _update_sap_input_count(self) -> None:
        text = self.input_mapper_sap_data_files.text().strip()
        if text:
            count = len([p for p in text.split(",") if p.strip()])
        else:
            count = 0
        lbl = getattr(self, "lbl_sap_input_count", None)
        if lbl is not None:
            lbl.setText(f"Input files count: {count}" if count > 0 else "")

    def _sync_mapper_compare_mode_ui(self) -> None:
        on = self.checkbox_compare_with_master.isChecked()
        rsd_mode = self.radio_compare_with_rsd_master.isChecked()
        self.radio_compare_with_rsd_master.setEnabled(on)
        self.radio_compare_only_tsc.setEnabled(on)
        self.btn_compare_rsd_master.setEnabled(on and rsd_mode)
        self.lbl_compare_rsd_master.setEnabled(on and rsd_mode)
        self.btn_compare_tsc_data.setEnabled(on)
        self.lbl_compare_tsc_data.setEnabled(on)

        lib_on = self.checkbox_compare_with_library.isChecked()
        self.btn_compare_library.setEnabled(lib_on)
        self.lbl_compare_library.setEnabled(lib_on)

        ref_on = self.checkbox_compare_run_ref_collector.isChecked()
        self.btn_compare_packshot_location.setEnabled(ref_on)
        self.lbl_compare_packshot_location.setEnabled(ref_on)
        self.lbl_max_packshot.setEnabled(ref_on)
        self.input_compare_max_packshot.setEnabled(ref_on)

        # When switching to Individual mode, flag multi-file master field immediately
        _ERROR_STYLE = (
            "background-color: rgba(208, 39, 82, 128); color: #000000; "
            "border: 1px solid #6F6F6F; border-radius: 11px; "
            "padding: 0 8px; font-family: 'Segoe UI'; font-size: 10px;"
        )
        is_individual = hasattr(self, "radio_sdc_individual") and self.radio_sdc_individual.isChecked()
        if is_individual and rsd_mode:
            master_text = self.lbl_compare_rsd_master.text().strip()
            parts = [p for p in master_text.split(",") if p.strip()]
            if len(parts) > 1:
                self.lbl_compare_rsd_master.setStyleSheet(_ERROR_STYLE)
                self.lbl_compare_rsd_master.setText("Only 1 file allowed in Individual mode")
        elif not is_individual:
            # Switching back to Set mode: clear any individual-mode error
            current = self.lbl_compare_rsd_master.text().strip()
            if current == "Only 1 file allowed in Individual mode":
                self.lbl_compare_rsd_master.setStyleSheet("")
                self.lbl_compare_rsd_master.setText("")

    def _sync_mapper_reformat_mode_ui(self) -> None:
        is_option_1 = self.radio_mapper_option_1.isChecked()

        self.btn_mapper_open_window.setEnabled(is_option_1)

        self.btn_mapper_sap_data_files.setEnabled(not is_option_1)
        self.input_mapper_sap_data_files.setEnabled(not is_option_1)
        self.btn_mapper_output_location.setEnabled(not is_option_1)
        self.input_mapper_output_location.setEnabled(not is_option_1)

        self.check_mapper_cleanup_1.setEnabled(not is_option_1)
        self.check_mapper_cleanup_2.setEnabled(not is_option_1)
        self.check_mapper_cleanup_3.setEnabled(not is_option_1)
        self.check_mapper_cleanup_4.setEnabled(not is_option_1)
        self.checkbox_mapper_include_grouping_report.setEnabled(not is_option_1)

        self.btn_run_process_mapper_reformat.setEnabled(not is_option_1)

        if is_option_1:
            self.input_mapper_sap_data_files.clear()
            self.input_mapper_output_location.clear()
            self.check_mapper_cleanup_1.setChecked(True)
            self.check_mapper_cleanup_2.setChecked(True)
            self.check_mapper_cleanup_3.setChecked(True)
            self.check_mapper_cleanup_4.setChecked(False)
            self.checkbox_mapper_include_grouping_report.setChecked(True)

    # ------------------------------------------------------------------
    # Option 2 handlers
    # ------------------------------------------------------------------

    def _on_mapper_select_sap_files(self) -> None:
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select SAP Data Files",
            self._get_browse_dir("sap_reformat"),
            "Excel Files (*.xls)",
        )
        if not file_paths:
            return
        self._mapper_sap_file_paths = file_paths
        names = [Path(p).name for p in file_paths]
        self.input_mapper_sap_data_files.setText(", ".join(names))

    def _on_mapper_select_output_location(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Folder", self._get_browse_dir("sap_reformat"))
        if not folder:
            return
        self._mapper_output_location = folder
        self.input_mapper_output_location.setText(folder)

    # ---- helpers reused from the SAP-table dialog (static-compatible) ----

    @staticmethod
    def _opt2_normalize_header(value: object) -> str:
        if value is None:
            return ""
        text = str(value).strip().lower()
        text = text.replace("\n", " ").replace("\r", " ")
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return " ".join(text.split())

    @staticmethod
    def _opt2_cell_to_text(value: object) -> str:
        if pd.isna(value):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def _opt2_read_excel_raw(self, file_path: str) -> pd.DataFrame:
        """Read a raw SAP .xls / .xlsx file into a headerless DataFrame."""
        lower = file_path.lower()

        if lower.endswith(".xls"):
            # Try openpyxl first (in case the .xls is actually xlsx)
            try:
                ef = pd.ExcelFile(file_path, engine="openpyxl")
                if ef.sheet_names:
                    return pd.read_excel(ef, sheet_name=ef.sheet_names[0], header=None, dtype=object)
            except Exception:
                pass

            # Try xlrd (genuine BIFF .xls)
            try:
                ef = pd.ExcelFile(file_path, engine="xlrd")
                if ef.sheet_names:
                    return pd.read_excel(ef, sheet_name=ef.sheet_names[0], header=None, dtype=object)
            except Exception:
                pass

            # Fallback – text-style SAP export
            text_df = self._opt2_read_text_style_xls(file_path)
            if text_df is not None:
                return text_df

            raise ValueError(f"Unable to read .xls file: {Path(file_path).name}")
        else:
            for engine in ("openpyxl", None):
                try:
                    ef = pd.ExcelFile(file_path, engine=engine)
                    if ef.sheet_names:
                        return pd.read_excel(ef, sheet_name=ef.sheet_names[0], header=None, dtype=object)
                except Exception:
                    pass
            raise ValueError(f"Unable to read file: {Path(file_path).name}")

    def _opt2_read_text_style_xls(self, file_path: str) -> pd.DataFrame | None:
        import csv as _csv
        parse_configs = [
            {"encoding": "utf-16", "sep": "\t", "skiprows": 3},
            {"encoding": "utf-16", "sep": r"\s{2,}", "skiprows": 3},
            {"encoding": "utf-16le", "sep": "\t", "skiprows": 3},
            {"encoding": "utf-8-sig", "sep": "\t", "skiprows": 3},
            {"encoding": "cp1252", "sep": "\t", "skiprows": 3},
            {"encoding": "utf-16", "sep": "\t", "skiprows": 0},
        ]
        for cfg in parse_configs:
            try:
                df = pd.read_csv(
                    file_path, header=None, dtype=object, engine="python",
                    on_bad_lines="skip", quotechar='"', quoting=_csv.QUOTE_MINIMAL, **cfg,
                )
            except Exception:
                continue
            if df.empty:
                continue
            df = df.dropna(how="all").reset_index(drop=True)
            if df.empty or df.shape[1] < 4:
                continue
            return df
        return None

    def _opt2_detect_headers(self, raw_df: pd.DataFrame) -> tuple[int, dict[str, int]]:
        target_columns: dict[str, list[str]] = {
            "Head Bom Mat": ["head bom mat", "head bom", "headbommat"],
            "BOM COMPONENT": ["bom component", "bom comp", "bomcomponent"],
            "Sort String": ["sort string", "sortstring"],
            "Component Desc": ["component desc", "component description", "component"],
            "Basic Number": ["basic number", "basic num", "basic no", "basic"],
            "Basic Name": ["basic name", "basicname"],
        }
        normalized_targets = {
            t: {self._opt2_normalize_header(a) for a in aliases}
            for t, aliases in target_columns.items()
        }

        max_scan = min(120, raw_df.shape[0])
        for row_idx in range(max_scan):
            row_norm = [self._opt2_normalize_header(v) for v in raw_df.iloc[row_idx].tolist()]
            found: dict[str, int] = {}
            for tname, aliases in normalized_targets.items():
                for ci, cv in enumerate(row_norm):
                    if cv in aliases:
                        found[tname] = ci
                        break
            if len(found) == len(target_columns):
                return row_idx, found

        raise ValueError("Unable to detect required SAP headers in file.")

    def _opt2_extract_rows(self, raw_df: pd.DataFrame, header_row: int, col_map: dict[str, int]) -> list[list[str]]:
        ordered = ["Head Bom Mat", "BOM COMPONENT", "Sort String", "Component Desc", "Basic Number", "Basic Name"]
        out: list[list[str]] = []
        for ri in range(header_row + 1, raw_df.shape[0]):
            vals: list[str] = []
            has = False
            for h in ordered:
                ci = col_map[h]
                v = raw_df.iat[ri, ci] if ci < raw_df.shape[1] else ""
                t = self._opt2_cell_to_text(v)
                if t:
                    has = True
                vals.append(t)
            if has:
                out.append(vals)
        return out

    @staticmethod
    def _opt2_build_grouping_data(reformatted_rows: list[list[str]]) -> list[tuple[str, int, str, str]]:
        """Build grouping counts from reformatted rows (same columns as the dialog table)."""
        # Reformatted row layout: [Head Bom Mat, HSI, BOM COMPONENT, Component Desc,
        #                          Basic Number, BC, Basic Name]
        combination_col = 5  # BC
        basic_number_col = 4
        basic_name_col = 6

        counts: dict[str, int] = {}
        bn_map: dict[str, dict[str, int]] = {}
        bname_map: dict[str, dict[str, int]] = {}

        for row in reformatted_rows:
            comb = row[combination_col].strip() if len(row) > combination_col else ""
            if not comb:
                continue
            counts[comb] = counts.get(comb, 0) + 1

            bn = row[basic_number_col].strip() if len(row) > basic_number_col else ""
            bn_map.setdefault(comb, {})
            bn_map[comb][bn] = bn_map[comb].get(bn, 0) + 1

            bname = row[basic_name_col].strip() if len(row) > basic_name_col else ""
            bname_map.setdefault(comb, {})
            bname_map[comb][bname] = bname_map[comb].get(bname, 0) + 1

        grouped: list[tuple[str, int, str, str]] = []
        for comb, cnt in counts.items():
            best_bn = ""
            if comb in bn_map and bn_map[comb]:
                best_bn = max(bn_map[comb].items(), key=lambda p: (p[1], p[0]))[0]
            best_name = ""
            if comb in bname_map and bname_map[comb]:
                best_name = max(bname_map[comb].items(), key=lambda p: (p[1], p[0]))[0]
            grouped.append((comb, cnt, best_bn, best_name))

        match = re.search  # local ref
        def sort_key(r):
            m = re.search(r"(\d+)", r[0])
            return (-r[1], int(m.group(1)) if m else 10**9, r[0])

        return sorted(grouped, key=sort_key)

    @staticmethod
    def _opt2_style_worksheet(ws, accent_color: str = "8A244B") -> None:
        """Apply the standard modern styling to a worksheet."""
        accent_fill = PatternFill(fill_type="solid", fgColor=accent_color)
        white_font = Font(color="FFFFFF", bold=True)
        normal_font = Font(color="111111")
        border = Border(
            left=Side(style="thin", color="B8B8B8"),
            right=Side(style="thin", color="B8B8B8"),
            top=Side(style="thin", color="B8B8B8"),
            bottom=Side(style="thin", color="B8B8B8"),
        )
        max_row = ws.max_row
        max_col = ws.max_column

        for col in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = accent_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        for row in range(2, max_row + 1):
            for col in range(1, max_col + 1):
                c = ws.cell(row=row, column=col)
                c.font = normal_font
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = border

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=max_col).coordinate}"

        for col_cells in ws.columns:
            letter = col_cells[0].column_letter
            mx = 0
            for cell in col_cells:
                v = "" if cell.value is None else str(cell.value)
                if len(v) > mx:
                    mx = len(v)
            ws.column_dimensions[letter].width = min(max(mx + 2, 12), 56)

    def _on_mapper_run_process(self) -> None:
        # Validate inputs
        file_paths = getattr(self, "_mapper_sap_file_paths", None) or []
        if not file_paths:
            QMessageBox.warning(self, "Missing Input", "Please select SAP data files first.")
            return
        output_dir = getattr(self, "_mapper_output_location", None) or ""
        if not output_dir:
            QMessageBox.warning(self, "Missing Input", "Please select an output location first.")
            return

        # Determine cleanup modes (all checked boxes)
        cleanup_modes = [
            i for i, cb in enumerate(
                [self.check_mapper_cleanup_1, self.check_mapper_cleanup_2,
                 self.check_mapper_cleanup_3, self.check_mapper_cleanup_4],
                start=1,
            )
            if cb.isChecked()
        ]
        if not cleanup_modes:
            QMessageBox.warning(self, "No Cleanup Selected", "Please select at least one cleanup mode.")
            return

        include_grouping = self.checkbox_mapper_include_grouping_report.isChecked()
        reformatter = SapTableReformatter()
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        errors: list[str] = []
        success_count = 0
        total_jobs = len(file_paths) * len(cleanup_modes)
        job_idx = 0

        # Show and reset progress bar
        self.mapper_reformat_progress_bar.setValue(0)
        self.mapper_reformat_progress_bar.setVisible(True)
        self.btn_run_process_mapper_reformat.setEnabled(False)
        QApplication.processEvents()

        for cleanup_mode in cleanup_modes:
            for file_path in file_paths:
                fname = Path(file_path).stem
                try:
                    self.mapper_reformat_progress_bar.setValue(int(job_idx / total_jobs * 30))
                    QApplication.processEvents()
                    raw_df = self._opt2_read_excel_raw(file_path)

                    self.mapper_reformat_progress_bar.setValue(int(job_idx / total_jobs * 60))
                    QApplication.processEvents()
                    header_row_idx, col_map = self._opt2_detect_headers(raw_df)
                    extracted_rows = self._opt2_extract_rows(raw_df, header_row_idx, col_map)
                    if not extracted_rows:
                        errors.append(f"{Path(file_path).name} (cu{cleanup_mode}): No data rows found.")
                        job_idx += 1
                        self.mapper_reformat_progress_bar.setValue(int(job_idx / total_jobs * 100))
                        QApplication.processEvents()
                        continue

                    self.mapper_reformat_progress_bar.setValue(int(job_idx / total_jobs * 80))
                    QApplication.processEvents()
                    reformatted_rows, basic_comb_count = reformatter.reformat_from_rows(extracted_rows, cleanup_mode)

                    wb = Workbook()
                    ws_reformat = wb.active
                    ws_reformat.title = "reformatted_sap"

                    reformat_header = [
                        "Head Bom Mat", "HSI", "BOM COMPONENT", "Component Desc",
                        "Basic Number", "BC", "Basic Name",
                    ]
                    ws_reformat.append(reformat_header)
                    for row_vals in reformatted_rows:
                        ws_reformat.append(row_vals)

                    self._opt2_style_worksheet(ws_reformat)

                    if include_grouping:
                        grouping_data = self._opt2_build_grouping_data(reformatted_rows)
                        ws_grouping = wb.create_sheet("grouping_data")
                        ws_grouping.append(["BC", "Count", "Basic Number", "Basic Name"])
                        for comb_name, count, bn, bname in grouping_data:
                            ws_grouping.append([comb_name, count, bn, bname])
                        self._opt2_style_worksheet(ws_grouping)

                    cu_tag = f"cu{cleanup_mode}"
                    out_name = f"rsd_{fname}_{cu_tag}_{timestamp}.xlsx"
                    out_path = out_dir / out_name
                    wb.save(str(out_path))
                    success_count += 1

                except Exception as exc:
                    errors.append(f"{Path(file_path).name} (cu{cleanup_mode}): {exc}")

                job_idx += 1
                self.mapper_reformat_progress_bar.setValue(int(job_idx / total_jobs * 100))
                QApplication.processEvents()

        self.mapper_reformat_progress_bar.setVisible(False)
        self.btn_run_process_mapper_reformat.setEnabled(True)

        # Summary message
        parts: list[str] = []
        if success_count:
            parts.append(f"Successfully processed {success_count} file(s).")
        if errors:
            parts.append("Errors:\n" + "\n".join(errors))

        msg = QMessageBox(self)
        msg.setWindowTitle("Process Complete" if success_count else "Error")
        msg.setIcon(QMessageBox.Icon.Information if success_count else QMessageBox.Icon.Warning)
        msg.setText("\n\n".join(parts))
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def _create_project_review_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        title = QLabel("Review Project")
        title.setObjectName("collectorTitle")
        layout.addWidget(title)

        # ── Mode radio buttons (one row) ───────────────────────────────────
        self.radio_pr_new_review = QRadioButton("New Review")
        self.radio_pr_new_review.setObjectName("collectorModeRadio")
        self.radio_pr_new_review.setChecked(True)
        self.radio_pr_continue = QRadioButton("Continue from Existing")
        self.radio_pr_continue.setObjectName("collectorModeRadio")

        self.pr_mode_group = QButtonGroup(page)
        self.pr_mode_group.setExclusive(True)
        self.pr_mode_group.addButton(self.radio_pr_new_review)
        self.pr_mode_group.addButton(self.radio_pr_continue)

        # Row 1: both radios on the same line
        # Row 2: count label aligned under "Continue from Existing" (col 1)
        self.lbl_pr_source_bma_count = QLabel("")
        self.lbl_pr_source_bma_count.setObjectName("inputCountLabel")

        mode_grid = QGridLayout()
        mode_grid.setContentsMargins(0, 0, 10, 0)
        mode_grid.setHorizontalSpacing(24)
        mode_grid.setVerticalSpacing(20)
        mode_grid.addWidget(self.radio_pr_new_review,       0, 0)
        mode_grid.addWidget(self.radio_pr_continue,         0, 1)
        mode_grid.addWidget(self.lbl_pr_source_bma_count,   1, 2, Qt.AlignmentFlag.AlignRight)
        mode_grid.setColumnStretch(1, 1)  # col 1 stretches; col 2 (label) hugs right edge
        layout.addLayout(mode_grid)

        # ── Source BMA (browse+edit row) ───────────────────────────────────
        bma_row = QHBoxLayout()
        bma_row.setSpacing(12)
        btn_bma_browse = QPushButton("Source BMA")
        btn_bma_browse.setObjectName("collectorGrayBtn")
        btn_bma_browse.clicked.connect(self._on_pr_source_bma_browse)
        bma_row.addWidget(btn_bma_browse, 0)
        self.input_pr_source_bma = QLineEdit()
        self.input_pr_source_bma.setObjectName("collectorLineEdit")
        self.input_pr_source_bma.setPlaceholderText("Select SDC file(s)…")
        bma_row.addWidget(self.input_pr_source_bma, 1)
        layout.addLayout(bma_row)

        # Wire up dynamic count + continue-mode autofill
        self.input_pr_source_bma.textChanged.connect(self._update_pr_source_bma_count)
        self.radio_pr_continue.toggled.connect(self._on_pr_mode_toggled)

        # ── Project Tracker (browse+edit row) ──────────────────────────────
        tracker_row = QHBoxLayout()
        tracker_row.setSpacing(12)
        btn_tracker_browse = QPushButton("Project Tracker")
        btn_tracker_browse.setObjectName("collectorGrayBtn")
        btn_tracker_browse.clicked.connect(self._on_pr_tracker_browse)
        tracker_row.addWidget(btn_tracker_browse, 0)
        self.input_pr_project_tracker = QLineEdit()
        self.input_pr_project_tracker.setObjectName("collectorLineEdit")
        self.input_pr_project_tracker.setPlaceholderText("Select Briefing Tracker file(s)…")
        tracker_row.addWidget(self.input_pr_project_tracker, 1)
        layout.addLayout(tracker_row)

        layout.addStretch(1)

        # ── Open Review Table button ───────────────────────────────────────
        run_row = QHBoxLayout()
        run_row.addStretch(1)
        self.btn_pr_open_review_table = QPushButton("Open Review Table")
        self.btn_pr_open_review_table.setObjectName("collectorRunBtn")
        self.btn_pr_open_review_table.setFixedHeight(50)
        self.btn_pr_open_review_table.setMinimumWidth(200)
        self.btn_pr_open_review_table.clicked.connect(self._on_pr_open_review_table_clicked)
        run_row.addWidget(self.btn_pr_open_review_table)
        run_row.addStretch(1)
        layout.addLayout(run_row)
        layout.addSpacing(8)

        return page

    @staticmethod
    def _pick_latest_review_file(files: list):
        """Delegates to review_project.pick_latest_review_file."""
        return review_project.pick_latest_review_file(files)

    def _on_pr_mode_toggled(self, checked: bool) -> None:
        """Autofill Source BMA when the mode radio changes.

        ``checked=True``  → 'Continue from Existing' active → load latest review file
        ``checked=False`` → 'New Review' active             → load SDC files
        """
        if not self._use_root_folders:
            return
        project_folder = self._get_active_project_folder()
        if not project_folder:
            return
        if checked:
            # Continue from Existing — latest review file from Project Review folder
            review_files = review_project.scan_review_files(project_folder)
            if review_files:
                chosen = review_files[0] if len(review_files) == 1 else review_project.pick_latest_review_file(review_files)
                self.input_pr_source_bma.setText(str(chosen))
            else:
                self.input_pr_source_bma.clear()
            self._autofill_pr_tracker(project_folder)
        else:
            # New Review — all SDC files from SAP Data Compare folder
            sdc_files = review_project.scan_sdc_files(project_folder)
            if sdc_files:
                self.input_pr_source_bma.setText(", ".join(str(f) for f in sdc_files))
            else:
                self.input_pr_source_bma.clear()

    def _update_pr_source_bma_count(self) -> None:
        """Update the Input files count label for Source BMA."""
        text = self.input_pr_source_bma.text().strip()
        count = len([p for p in text.split(",") if p.strip()]) if text else 0
        lbl = getattr(self, "lbl_pr_source_bma_count", None)
        if lbl is not None:
            lbl.setText(f"Input files count: {count}" if count > 0 else "")

    def _on_pr_tracker_browse(self) -> None:
        """Open file dialog to select one or more Briefing Tracker files."""
        from pathlib import Path
        start = ""
        project_folder = self._get_active_project_folder()
        if project_folder:
            pr = Path(project_folder) / "Project Review"
            start = str(pr) if pr.is_dir() else str(project_folder)
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Briefing Tracker file(s)", start,
            "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        if files:
            self.input_pr_project_tracker.setText(", ".join(files))

    def _on_pr_source_bma_browse(self) -> None:
        """Open file dialog to select Source BMA file(s).
        New Review: multiple files allowed. Continue from Existing: single file only."""
        from pathlib import Path
        start = ""
        project_folder = self._get_active_project_folder()
        continue_mode = getattr(self, "radio_pr_continue", None) and self.radio_pr_continue.isChecked()
        if project_folder:
            if continue_mode:
                pr = Path(project_folder) / "Project Review"
                start = str(pr) if pr.is_dir() else str(project_folder)
            else:
                sdc = Path(project_folder) / "SAP Data Compare"
                start = str(sdc) if sdc.is_dir() else str(project_folder)
        if continue_mode:
            file, _ = QFileDialog.getOpenFileName(
                self, "Select Source BMA file", start,
                "Excel Files (*.xlsx *.xls *.xlsm)"
            )
            if file:
                self.input_pr_source_bma.setText(file)
        else:
            files, _ = QFileDialog.getOpenFileNames(
                self, "Select Source BMA file(s)", start,
                "Excel Files (*.xlsx *.xls *.xlsm)"
            )
            if files:
                self.input_pr_source_bma.setText(", ".join(files))

    def _on_pr_open_review_table_clicked(self) -> None:
        """Open the Review Table dialog for the first selected Source BMA file."""
        from pathlib import Path
        text = self.input_pr_source_bma.text().strip()
        if not text:
            QMessageBox.warning(self, "No File", "Please select a Source BMA file first.")
            return
        # Use the first file listed (continue mode = single file; new review = first of multiple)
        first_file = text.split(",")[0].strip()
        if not Path(first_file).exists():
            QMessageBox.warning(self, "File Not Found", f"File not found:\n{first_file}")
            return
        # Determine SDC folder for "Import another BMA" starting directory
        sdc_folder = ""
        project_review_folder = ""
        project_folder = self._get_active_project_folder()
        if project_folder:
            from pathlib import Path as _Path
            sdc = _Path(project_folder) / "SAP Data Compare"
            if sdc.is_dir():
                sdc_folder = str(sdc)
            pr = _Path(project_folder) / "Project Review"
            pr.mkdir(parents=True, exist_ok=True)
            project_review_folder = str(pr)
        # Parse tracker files
        tracker_files: list[str] = []
        if hasattr(self, "input_pr_project_tracker"):
            tf_text = self.input_pr_project_tracker.text().strip()
            if tf_text:
                from pathlib import Path as _P2
                tracker_files = [p.strip() for p in tf_text.split(",") if _P2(p.strip()).exists()]
        dlg = _ReviewTableDialog(
            first_file,
            sdc_folder=sdc_folder,
            project_review_folder=project_review_folder,
            tracker_files=tracker_files,
            load_colors=self.radio_pr_continue.isChecked(),
            parent=None,          # non-modal: no parent → independent window
        )
        dlg.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        # Keep a Python reference so the GC doesn't collect the window while it's open
        if not hasattr(self, "_review_table_dialogs"):
            self._review_table_dialogs = []
        self._review_table_dialogs.append(dlg)
        dlg.destroyed.connect(
            lambda: self._review_table_dialogs.remove(dlg)
            if dlg in self._review_table_dialogs else None
        )
        dlg.show()

    class _MissingValuesResultDialog(QDialog):
        """Dialog that displays missing values and offers clipboard copying."""

        def __init__(self, values: list[str], parent=None):
            super().__init__(parent)
            self.setWindowTitle("Missing Values Result")
            self.resize(560, 420)

            layout = QVBoxLayout(self)
            layout.setContentsMargins(20, 20, 20, 20)
            layout.setSpacing(12)

            title_row = QHBoxLayout()
            title_row.setContentsMargins(0, 0, 0, 0)
            title = QLabel("Missing Values")
            title.setObjectName("collectorTitle")
            title_row.addWidget(title)
            title_row.addSpacing(100)
            count_label = QLabel(f"count: {len(values)}")
            count_label.setObjectName("collectorSectionLabel")
            count_label.setContentsMargins(0, 0, 0, 0)
            title_row.addWidget(count_label)
            layout.addLayout(title_row)

            body = QTextEdit(self)
            body.setReadOnly(True)
            body.setPlainText("\n".join(values) if values else "No missing values found.")
            layout.addWidget(body, 1)

            actions = QHBoxLayout()
            actions.addStretch(1)
            copy_btn = QPushButton("Copy values to clipboard")
            copy_btn.clicked.connect(lambda: QApplication.clipboard().setText(body.toPlainText()))
            actions.addWidget(copy_btn)
            close_btn = QPushButton("Close")
            close_btn.clicked.connect(self.close)
            actions.addWidget(close_btn)
            layout.addLayout(actions)

    class _MissingValuesCheckWindow(QDialog):
        """Dialog for checking missing values between an Excel main file and a target file."""

        def __init__(self, parent=None, mode: str = "Excel vs Excel"):
            super().__init__(parent)
            self.setWindowTitle("Missing Values Check")
            self.resize(640, 320)

            layout = QVBoxLayout(self)
            layout.setContentsMargins(20, 20, 20, 20)
            layout.setSpacing(12)

            mode_label = QLabel(mode.lower())
            mode_label.setObjectName("collectorSectionLabel")
            layout.addWidget(mode_label)

            main_row = QHBoxLayout()
            self.btn_missing_main_file = QPushButton("Main File")
            self.btn_missing_main_file.clicked.connect(self._browse_main_file)
            self.input_missing_main_file = QLineEdit()
            self.input_missing_main_file.setReadOnly(True)
            self.input_missing_main_file.setPlaceholderText("Select main Excel file")
            main_row.addWidget(self.btn_missing_main_file)
            main_row.addWidget(self.input_missing_main_file, 1)
            layout.addLayout(main_row)

            target_row = QHBoxLayout()
            self.btn_missing_target_file = QPushButton("Target File")
            self.btn_missing_target_file.clicked.connect(self._browse_target_file)
            self.input_missing_target_file = QLineEdit()
            self.input_missing_target_file.setReadOnly(True)
            self.input_missing_target_file.setPlaceholderText("Select target Excel file")
            target_row.addWidget(self.btn_missing_target_file)
            target_row.addWidget(self.input_missing_target_file, 1)
            layout.addLayout(target_row)

            main_column_row = QHBoxLayout()
            main_column_label = QLabel("Main file column")
            main_column_row.addWidget(main_column_label)
            self.combo_missing_main_column = QComboBox()
            self.combo_missing_main_column.setEnabled(False)
            self.combo_missing_main_column.setMinimumHeight(34)
            self.combo_missing_main_column.setMaximumWidth(260)
            self.combo_missing_main_column.currentTextChanged.connect(self._refresh_missing_values_button_state)
            main_column_row.addWidget(self.combo_missing_main_column, 0)
            layout.addLayout(main_column_row)

            target_column_row = QHBoxLayout()
            target_column_label = QLabel("Target file column")
            target_column_row.addWidget(target_column_label)
            self.combo_missing_target_column = QComboBox()
            self.combo_missing_target_column.setEnabled(False)
            self.combo_missing_target_column.setMinimumHeight(34)
            self.combo_missing_target_column.setMaximumWidth(260)
            self.combo_missing_target_column.currentTextChanged.connect(self._refresh_missing_values_button_state)
            target_column_row.addWidget(self.combo_missing_target_column, 0)
            layout.addLayout(target_column_row)

            self.chk_missing_export = QCheckBox(
                "export check missing values report (file will be a .xlsx and will be saved in main file location.)"
            )
            self.chk_missing_export.setToolTip(
                "export check missing values report (file will be a .xlsx and will be saved in main file location.)"
            )
            layout.addWidget(self.chk_missing_export)

            self.btn_check_missing_values = QPushButton("Check Missing Values")
            self.btn_check_missing_values.setEnabled(False)
            self.btn_check_missing_values.clicked.connect(self._run_missing_values_check)
            layout.addWidget(self.btn_check_missing_values, alignment=Qt.AlignmentFlag.AlignRight)

            layout.addStretch(1)

            self.setStyleSheet("""
                QDialog { background-color: #F4F4F4; }
                QLabel { font-family: 'Segoe UI'; font-size: 13px; color: #333333; }
                QLineEdit {
                    background-color: #FFFFFF; color: #111111;
                    border: 1px solid #BBBBBB; border-radius: 6px;
                    padding: 0 8px; min-height: 34px;
                    font-family: 'Segoe UI'; font-size: 12px;
                }
                QComboBox {
                    background-color: #FFFFFF; color: #111111;
                    border: 1px solid #BBBBBB; border-radius: 6px;
                    padding: 4px 8px; min-height: 34px;
                    font-family: 'Segoe UI'; font-size: 12px;
                }
                QPushButton {
                    background-color: #9EA3AB; color: #000000;
                    border: 1px solid #8B9098; border-radius: 6px;
                    padding: 0 12px; min-height: 34px;
                    font-family: 'Segoe UI'; font-size: 12px; font-weight: 600;
                }
                QPushButton:hover { background-color: #ACB1B8; }
                QPushButton:pressed { background-color: #111F35; color: #ffffff; }
                QCheckBox { font-family: 'Segoe UI'; font-size: 12px; color: #333333; }
            """)

        def _browse_main_file(self) -> None:
            start_dir = self._get_browse_dir("packshot") if hasattr(self, "_get_browse_dir") else ""
            file, _ = QFileDialog.getOpenFileName(
                self,
                "Select Main Excel File",
                start_dir,
                "Excel Files (*.xlsx *.xls *.xlsm)",
            )
            if file:
                self.input_missing_main_file.setText(file)
                self._populate_column_dropdown(file, self.combo_missing_main_column)
                self._refresh_missing_values_button_state()

        def _browse_target_file(self) -> None:
            start_dir = self._get_browse_dir("packshot") if hasattr(self, "_get_browse_dir") else ""
            file, _ = QFileDialog.getOpenFileName(
                self,
                "Select Target Excel File",
                start_dir,
                "Excel Files (*.xlsx *.xls *.xlsm)",
            )
            if file:
                self.input_missing_target_file.setText(file)
                self._populate_column_dropdown(file, self.combo_missing_target_column)
                self._refresh_missing_values_button_state()

        def _populate_column_dropdown(self, file_path: str, combo: QComboBox) -> None:
            try:
                excel_file = pd.ExcelFile(file_path)
                first_sheet = excel_file.sheet_names[0]
                df = pd.read_excel(file_path, sheet_name=first_sheet)
            except Exception as exc:
                QMessageBox.warning(self, "Unable to read file", f"Could not read Excel columns from selected file.\n{exc}")
                combo.clear()
                combo.setEnabled(False)
                return

            columns = [str(col) for col in df.columns if str(col).strip()]
            combo.clear()
            combo.addItems(columns)
            combo.setEnabled(bool(columns))
            if not columns:
                QMessageBox.warning(self, "No columns found", "The selected Excel file does not contain any usable columns.")

        def _refresh_missing_values_button_state(self) -> None:
            main_path = self.input_missing_main_file.text().strip()
            target_path = self.input_missing_target_file.text().strip()
            main_has_column = self.combo_missing_main_column.count() > 0 and self.combo_missing_main_column.currentText() != ""
            target_has_column = self.combo_missing_target_column.count() > 0 and self.combo_missing_target_column.currentText() != ""
            self.btn_check_missing_values.setEnabled(bool(main_path and target_path and main_has_column and target_has_column))

        def _run_missing_values_check(self) -> None:
            main_path = self.input_missing_main_file.text().strip()
            target_path = self.input_missing_target_file.text().strip()
            main_column = self.combo_missing_main_column.currentText().strip()
            target_column = self.combo_missing_target_column.currentText().strip()
            if not main_path or not target_path or not main_column or not target_column:
                QMessageBox.warning(self, "Missing input", "Please select both Excel files and both columns before checking.")
                return
            if not Path(main_path).exists() or not Path(target_path).exists():
                QMessageBox.warning(self, "File not found", "One of the selected files could not be found.")
                return

            try:
                main_df = pd.read_excel(main_path)
                target_df = pd.read_excel(target_path)
            except Exception as exc:
                QMessageBox.warning(self, "Read error", f"Could not read the Excel files.\n{exc}")
                return

            if main_column not in main_df.columns:
                QMessageBox.warning(self, "Column not found", f"The main-file column '{main_column}' was not found in the selected main file.")
                return
            if target_column not in target_df.columns:
                QMessageBox.warning(self, "Column not found", f"The target-file column '{target_column}' was not found in the selected target file.")
                return

            main_values = main_df[main_column].dropna().astype(str).str.strip()
            target_values = target_df[target_column].dropna().astype(str).str.strip()
            target_values_set = {v.casefold() for v in target_values.tolist() if str(v).strip()}

            missing_values: list[str] = []
            seen_values: set[str] = set()
            for value in main_values.tolist():
                value_text = str(value).strip()
                if not value_text:
                    continue
                normalized = value_text.casefold()
                if normalized in target_values_set or normalized in seen_values:
                    continue
                missing_values.append(value_text)
                seen_values.add(normalized)

            result_dialog_cls = NewUIWindow._MissingValuesResultDialog
            results_dialog = result_dialog_cls(missing_values, self)
            results_dialog.show()
            results_dialog.raise_()
            results_dialog.activateWindow()

            if self.chk_missing_export.isChecked():
                export_path = Path(main_path).parent / (
                    f"missing_values_check_{datetime.now().strftime('%Y_%m_%d_%H_%M_%S')}.xlsx"
                )
                wb = Workbook()
                ws = wb.active
                ws.title = "Missing Values Check"
                ws.append(["Main Column", "Target Column", "Missing Values", "Main File", "Target File"])
                for value in missing_values:
                    ws.append([main_column, target_column, value, main_path, target_path])
                wb.save(export_path)
                QMessageBox.information(self, "Export complete", f"Report saved to:\n{export_path}")

    def _create_other_tools_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        title = QLabel("Other Tools")
        title.setObjectName("collectorTitle")
        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.addWidget(title, 1)
        layout.addLayout(title_row)

        # ── Duplicate Check ───────────────────────────────────────────────────
        section_label = QLabel("Duplicate Check")
        section_label.setObjectName("collectorSectionLabel")
        layout.addWidget(section_label)
        layout.addSpacing(-8)

        self.radio_idh_dup_single = QRadioButton("Single Tracker")
        self.radio_idh_dup_single.setObjectName("idhDupRadio")
        self.radio_idh_dup_single.setChecked(True)
        self.radio_idh_dup_multiple = QRadioButton("Multiple Trackers")
        self.radio_idh_dup_multiple.setObjectName("idhDupRadio")

        self.idh_dup_mode_group = QButtonGroup(page)
        self.idh_dup_mode_group.setExclusive(True)
        self.idh_dup_mode_group.addButton(self.radio_idh_dup_single)
        self.idh_dup_mode_group.addButton(self.radio_idh_dup_multiple)

        self.btn_idh_dup_open_window = QPushButton("Open Window")
        self.btn_idh_dup_open_window.setObjectName("collectorGrayBtn")

        controls_row = QHBoxLayout()
        controls_row.setSpacing(12)
        controls_row.addWidget(self.radio_idh_dup_single)
        controls_row.addWidget(self.radio_idh_dup_multiple)
        controls_row.addWidget(self.btn_idh_dup_open_window)
        controls_row.addStretch(1)
        layout.addLayout(controls_row)

        self.btn_idh_dup_open_window.clicked.connect(self._open_idh_dup_window)

        # ── Missing Values Check ─────────────────────────────────────────────
        layout.addSpacing(8)
        section_label = QLabel("Missing Values Check")
        section_label.setObjectName("collectorSectionLabel")
        layout.addWidget(section_label)
        layout.addSpacing(-8)

        self.radio_missing_vals_excel_vs_excel = QRadioButton("Excel vs Excel")
        self.radio_missing_vals_excel_vs_excel.setObjectName("idhDupRadio")
        self.radio_missing_vals_excel_vs_excel.setChecked(True)
        self.radio_missing_vals_excel_vs_images = QRadioButton("Excel vs Images")
        self.radio_missing_vals_excel_vs_images.setObjectName("idhDupRadio")

        self.missing_vals_mode_group = QButtonGroup(page)
        self.missing_vals_mode_group.setExclusive(True)
        self.missing_vals_mode_group.addButton(self.radio_missing_vals_excel_vs_excel)
        self.missing_vals_mode_group.addButton(self.radio_missing_vals_excel_vs_images)

        self.btn_missing_vals_open_window = QPushButton("Open Window")
        self.btn_missing_vals_open_window.setObjectName("collectorGrayBtn")

        controls_row = QHBoxLayout()
        controls_row.setSpacing(12)
        controls_row.addWidget(self.radio_missing_vals_excel_vs_excel)
        controls_row.addWidget(self.radio_missing_vals_excel_vs_images)
        controls_row.addWidget(self.btn_missing_vals_open_window)
        controls_row.addStretch(1)
        layout.addLayout(controls_row)

        self.btn_missing_vals_open_window.clicked.connect(self._open_missing_values_window)

        layout.addStretch(1)
        return page

    def _create_packshot_naming_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(16)

        title = QLabel("Packshot Naming Generator")
        title.setObjectName("packshotTitle")
        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.addWidget(title, 1)
        layout.addLayout(title_row)

        description = QLabel("NOTE: All fields must have input.")
        description.setObjectName("packshotDescription")
        layout.addWidget(description)
        layout.addSpacing(12)

        self.radioButton_pnm_from_clipboard = QRadioButton("From clipboard")
        self.radioButton_pnm_from_clipboard.setObjectName("collectorModeRadio")
        self.radioButton_pnm_generate_from_tracker = QRadioButton("From tracker")
        self.radioButton_pnm_generate_from_tracker.setObjectName("collectorModeRadio")
        self.radioButton_pnm_from_clipboard.setChecked(True)

        self.packshot_mode_group = QButtonGroup(page)
        self.packshot_mode_group.setExclusive(True)
        self.packshot_mode_group.addButton(self.radioButton_pnm_from_clipboard)
        self.packshot_mode_group.addButton(self.radioButton_pnm_generate_from_tracker)

        layout.addWidget(self.radioButton_pnm_from_clipboard)

        clipboard_row = QHBoxLayout()
        clipboard_row.setSpacing(12)
        self.btn_pnm_paste_on_table = QPushButton("Paste on Table")
        self.btn_pnm_paste_on_table.setObjectName("collectorGrayBtn")
        clipboard_row.addWidget(self.btn_pnm_paste_on_table, 0)
        clipboard_row.addStretch(1)
        layout.addLayout(clipboard_row)

        layout.addSpacing(8)

        pnm_tracker_header = QHBoxLayout()
        pnm_tracker_header.setSpacing(0)
        pnm_tracker_header.addWidget(self.radioButton_pnm_generate_from_tracker)
        pnm_tracker_header.addSpacing(80)
        self.lbl_png_input_count = QLabel("")
        self.lbl_png_input_count.setObjectName("inputCountLabel")
        pnm_tracker_header.addWidget(self.lbl_png_input_count, 0)
        pnm_tracker_header.addStretch(1)
        layout.addLayout(pnm_tracker_header)

        self.packshot_tracker_row = QHBoxLayout()
        self.packshot_tracker_row.setSpacing(12)
        self.btn_pg6_excel_tracker = QPushButton("Excel Tracker")
        self.btn_pg6_excel_tracker.setObjectName("collectorGrayBtn")
        self.packshot_tracker_row.addWidget(self.btn_pg6_excel_tracker, 0)
        self.input_pg6_excel_tracker = QLineEdit()
        self.input_pg6_excel_tracker.setObjectName("collectorLineEdit")
        self.packshot_tracker_row.addWidget(self.input_pg6_excel_tracker, 1)
        layout.addLayout(self.packshot_tracker_row)

        output_row = QHBoxLayout()
        output_row.setSpacing(12)
        self.btn_pg6_output = QPushButton("Output Location")
        self.btn_pg6_output.setObjectName("collectorGrayBtn")
        output_row.addWidget(self.btn_pg6_output, 0)
        self.input_pg6_output = QLineEdit()
        self.input_pg6_output.setObjectName("collectorLineEdit")
        output_row.addWidget(self.input_pg6_output, 1)
        layout.addLayout(output_row)

        layout.addStretch(1)
        run_row = QHBoxLayout()
        run_row.addStretch(1)
        self.btn_run_process_generate_packshot_naming = QPushButton("Run Process")
        self.btn_run_process_generate_packshot_naming.setObjectName("collectorRunBtn")
        self.btn_run_process_generate_packshot_naming.setFixedHeight(50)
        self.btn_run_process_generate_packshot_naming.setMinimumWidth(180)
        run_row.addWidget(self.btn_run_process_generate_packshot_naming)
        run_row.addStretch(1)
        layout.addLayout(run_row)

        self.radioButton_pnm_from_clipboard.toggled.connect(self._sync_packshot_mode_ui)
        self.radioButton_pnm_generate_from_tracker.toggled.connect(self._sync_packshot_mode_ui)
        self.btn_pnm_paste_on_table.clicked.connect(self._open_packshot_clipboard_table)
        self.input_pg6_excel_tracker.textChanged.connect(lambda: self._update_png_input_count())
        self._sync_packshot_mode_ui()

        return page

    def _open_packshot_clipboard_table(self) -> None:
        self._packshot_clipboard_dialog = _PackshotClipboardTableDialog(self, start_dir=self._get_browse_dir("packshot"))
        if hasattr(self, "packshot_naming_generator"):
            self.packshot_naming_generator.attach_table_dialog(self._packshot_clipboard_dialog)
        self._packshot_clipboard_dialog.show()
        self._packshot_clipboard_dialog.raise_()
        self._packshot_clipboard_dialog.activateWindow()

    def _open_idh_dup_window(self) -> None:
        start_dir = self._get_browse_dir("packshot")
        if getattr(self, "radio_idh_dup_single", None) and self.radio_idh_dup_single.isChecked():
            self._idh_dup_window = _SingleTrackerWindow(self, start_dir=start_dir)
        else:
            self._idh_dup_window = _MultipleTrackerWindow(self, start_dir=start_dir)
        self._idh_dup_window.show()
        self._idh_dup_window.raise_()
        self._idh_dup_window.activateWindow()

    def _open_missing_values_window(self) -> None:
        mode = "Excel vs Excel"
        if getattr(self, "radio_missing_vals_excel_vs_images", None) and self.radio_missing_vals_excel_vs_images.isChecked():
            mode = "Excel vs Images"
        self._missing_values_window = self._MissingValuesCheckWindow(self, mode=mode)
        self._missing_values_window.show()
        self._missing_values_window.raise_()
        self._missing_values_window.activateWindow()

    def _open_tsc_option1_table(self) -> None:
        self._tsc_option1_dialog = _TscOption1TableDialog(
            self,
            start_dir=self._get_browse_dir("tsc"),
            tsc_output_dir=self._get_browse_dir("tsc_output"),
        )
        # Connect Import Trackers button to status_collector logic
        if hasattr(self, "status_collector"):
            self.status_collector.attach_tracker_window_dialog(self._tsc_option1_dialog)
        self._tsc_option1_dialog.show()
        self._tsc_option1_dialog.raise_()
        self._tsc_option1_dialog.activateWindow()

    def _open_mapper_option1_table(self) -> None:
        from pathlib import Path as _Path
        _export_dir = ""
        _is_project = (self.left_nav_stack.currentWidget() is self.project_left)
        if _is_project:
            _proj = self._get_active_project_folder()
            if _proj:
                _export_dir = str(_Path(_proj) / "SAP Data Compare")
        else:
            _export_dir = self._config.sap_reformat_input() or ""
        self._mapper_option1_dialog = _MapperOption1TableDialog(
            self, start_dir=self._get_browse_dir("sap_reformat"),
            export_dir=_export_dir)
        self._mapper_option1_dialog.show()
        self._mapper_option1_dialog.raise_()
        self._mapper_option1_dialog.activateWindow()

    def _sync_packshot_mode_ui(self) -> None:
        from_clipboard = getattr(self, "radioButton_pnm_from_clipboard", None)
        if from_clipboard is None:
            return
        is_clipboard = from_clipboard.isChecked()

        if hasattr(self, "btn_pnm_paste_on_table"):
            self.btn_pnm_paste_on_table.setEnabled(is_clipboard)
        if hasattr(self, "btn_pg6_excel_tracker"):
            self.btn_pg6_excel_tracker.setEnabled(not is_clipboard)
        if hasattr(self, "input_pg6_excel_tracker"):
            self.input_pg6_excel_tracker.setEnabled(not is_clipboard)
            if is_clipboard:
                self.input_pg6_excel_tracker.clear()
        if hasattr(self, "btn_pg6_output"):
            self.btn_pg6_output.setEnabled(not is_clipboard)
        if hasattr(self, "input_pg6_output"):
            self.input_pg6_output.setEnabled(not is_clipboard)
            if is_clipboard:
                self.input_pg6_output.clear()
        if hasattr(self, "btn_run_process_generate_packshot_naming"):
            self.btn_run_process_generate_packshot_naming.setEnabled(not is_clipboard)

    def _wrap_button_text(self, text: str, words_per_line: int = 2) -> str:
        words = text.split()
        lines = []
        for i in range(0, len(words), words_per_line):
            lines.append(" ".join(words[i : i + words_per_line]))
        return "\n".join(lines)

    def _get_browse_dir(self, tool: str) -> str:
        """Return the configured start directory for a file dialog.
        When the configure TSB is ON returns the tool's configured folder
        (falls back to Desktop if the path doesn't exist). When OFF returns
        an empty string so the OS uses its default last-used location."""
        if not self._use_root_folders:
            return ""
        from pathlib import Path
        desktop = str(Path.home() / "Desktop")

        if tool == "tsc":
            path = self._config.briefing_tracker_input()
            return path if path and Path(path).is_dir() else desktop

        if tool == "tsc_output":
            path = self._config.tsc_output()
            return path if path and Path(path).is_dir() else desktop

        if tool == "thumbnail":
            path = self._config.thumbnail_input()
            return path if path and Path(path).is_dir() else desktop

        is_project = (self.left_nav_stack.currentWidget() is self.project_left)

        if tool == "excel_library":
            path = self._config.excel_library_folder()
            return path if path and Path(path).is_dir() else desktop

        if tool == "sap_compare_master":
            path = self._config.rsd_master_folder()
            return path if path and Path(path).is_dir() else desktop

        if tool == "sap_compare":
            if is_project:
                proj = self._get_active_project_folder()
                if proj:
                    p = Path(proj) / "SAP Data Reformat"
                    return str(p) if p.is_dir() else desktop
            else:
                path = self._config.sap_reformat_input()
                return path if path and Path(path).is_dir() else desktop

        if tool == "sap_reformat":
            if is_project:
                proj = self._get_active_project_folder()
                if proj:
                    p = Path(proj) / "SAP Data Reformat"
                    return str(p) if p.is_dir() else desktop
            else:
                path = self._config.sap_reformat_input()
                return path if path and Path(path).is_dir() else desktop

        if tool == "packshot":
            proj = self._get_active_project_folder()
            if proj:
                p = Path(proj) / "Packshot Naming Generator"
                return str(p) if p.is_dir() else desktop

        return desktop

    def _open_settings_dialog(self) -> None:
        btn = self.btn_settings
        btn.setStyleSheet(
            "QPushButton#iconBtn { background-color: #8A244B; color: #ffffff; "
            "border: 1px solid #8A244B; border-radius: 19px; font-size: 16px; font-weight: 700; padding: 0; }"
        )
        QTimer.singleShot(300, lambda: btn.setStyleSheet(""))
        dlg = _SettingsDialog(self._config, self)
        dlg.exec()

    def _init_modules(self) -> None:
        self._init_status_collector_module()
        self._init_thumbnail_generator_module()
        self._init_packshot_naming_module()

    def _init_status_collector_module(self) -> None:
        # Bridge line edits to legacy API expected by status_collector.py.
        self.textEdit_sc_selected_trackers = _PlainTextLineAdapter(self.input_trackers)
        self.textEdit_sc_output_location = _PlainTextLineAdapter(self.input_output)
        self.textEdit_status_collector_status_input = _PlainTextLineAdapter(self._input_status_line)

        # Optional clear-all control expected by the module; hidden in this UI.
        self.btn_menu_clear_all_fields = QPushButton("Clear All")
        self.btn_menu_clear_all_fields.setVisible(False)

        self.status_collector = StatusCollector(self)
        self.status_collector.bind_input_type_to_radioboxes()
        self.status_collector.run_process()

        # Override browse buttons to respect the configure TSB start directory.
        self.status_collector.btn_select_trackers.clicked.disconnect()
        self.status_collector.btn_select_trackers.clicked.connect(
            lambda: self.status_collector.get_tracker_files(
                self._get_browse_dir("tsc")))
        self.status_collector.btn_output_location.clicked.disconnect()
        self.status_collector.btn_output_location.clicked.connect(
            lambda: self.status_collector.get_output_location(
                self._get_browse_dir("tsc_output")))

    def _init_thumbnail_generator_module(self) -> None:
        self.thumbnail_generator = ThumbnailGenerator(self)
        self.thumbnail_generator.run_process()

        # Override browse button to respect the configure TSB start directory.
        self.thumbnail_generator.btn_images_folder.clicked.disconnect()
        self.thumbnail_generator.btn_images_folder.clicked.connect(
            lambda: self.thumbnail_generator.browse_images_folder(
                self._get_browse_dir("thumbnail")))

    def _init_packshot_naming_module(self) -> None:
        self.packshot_naming_generator = PackshotNamingGenerator(self)
        self.packshot_naming_generator.run_process()

        # Override browse buttons to respect the configure TSB start directory.
        if self.packshot_naming_generator.btn_excel_tracker is not None:
            self.packshot_naming_generator.btn_excel_tracker.clicked.disconnect()
            self.packshot_naming_generator.btn_excel_tracker.clicked.connect(
                lambda: self.packshot_naming_generator.name_gen_get_tracker_file(
                    self._get_browse_dir("packshot")))
        if self.packshot_naming_generator.btn_output_location is not None:
            self.packshot_naming_generator.btn_output_location.clicked.disconnect()
            self.packshot_naming_generator.btn_output_location.clicked.connect(
                lambda: self.packshot_naming_generator.name_gen_get_output_location(
                    self._get_browse_dir("packshot")))

    def _stylesheet(self) -> str:
        return """
        #root {
            background-color: #E5E5E5;
        }

        #shell {
            background-color: #E5E5E5;
            border-radius: 0px;
            border: none;
        }

        #topHeaderBtn {
            background-color: transparent;
            color: #ffffff;
            border: none;
            border-radius: 22px;
            padding: 0 18px;
            font-family: "Bahnschrift SemiCondensed", "Arial Narrow", "Segoe UI";
            font-size: 14px;
            font-weight: 800;
        }

        #topHeaderBtn:hover {
            background-color: transparent;
            color: #ffffff;
            border: none;
        }

        #topHeaderBtn:checked {
            background-color: #F63049;
            border: none;
            color: #ffffff;
        }

        #topHeaderBtn:checked:hover {
            background-color: #F63049;
            border: none;
            color: #ffffff;
        }

        #iconBtn {
            background-color: #171b21;
            color: #ffffff;
            border: 1px solid #4b5664;
            border-radius: 19px;
            font-size: 16px;
            font-weight: 700;
            padding: 0;
        }

        #iconBtn:hover {
            border: 1px solid #707c8b;
        }

        #contentCard {
            background-color: #E5E5E5;
            border: 1px solid #C0C0C0;
            border-radius: 18px;
        }

        #homeCard {
            background-color: #FFFFFF;
            border: 1px solid #C0C0C0;
            border-radius: 18px;
        }

        #searchPanel {
            background-color: #FFFFFF;
            border: 1px solid #C0C0C0;
            border-radius: 18px;
        }

        #artworkPanel {
            background-color: #FFFFFF;
            border: 1px solid #C0C0C0;
            border-radius: 18px;
        }

        #artworkTitle {
            color: #8A244B;
            font-family: "Segoe UI";
            font-size: 26px;
            font-weight: 800;
        }

        #artworkDescription, #artworkMutedLabel {
            color: #5D6875;
            font-family: "Segoe UI";
            font-size: 13px;
        }

        #artworkFieldLabel {
            min-width: 105px;
            color: #111F35;
            font-family: "Segoe UI";
            font-size: 13px;
            font-weight: 700;
        }

        #artworkInput, #artworkSmallInput {
            background-color: #F2F2F2;
            color: #111111;
            border: 1px solid #C8C8C8;
            border-radius: 7px;
            min-height: 34px;
            padding: 0 10px;
            font-family: "Segoe UI";
            font-size: 13px;
        }

        #artworkInput:focus, #artworkSmallInput:focus {
            border: 1px solid #8A244B;
            background-color: #FFFFFF;
        }

        #artworkBrowseButton {
            background-color: #111F35;
            color: #FFFFFF;
            border: none;
            border-radius: 7px;
            min-width: 82px;
            min-height: 34px;
            font-weight: 700;
        }

        #artworkBrowseButton:hover {
            background-color: #273B5A;
        }

        #artworkProcessButton {
            background-color: #D02752;
            color: #FFFFFF;
            border: none;
            border-radius: 8px;
            padding: 0 20px;
            font-weight: 800;
        }

        #artworkProcessButton:hover {
            background-color: #B51F45;
        }

        #artworkCheckBox {
            color: #111F35;
            font-size: 13px;
        }

        #artworkLog {
            background-color: #F7F7F7;
            color: #222222;
            border: 1px solid #D0D0D0;
            border-radius: 8px;
            padding: 8px;
            font-family: "Consolas";
            font-size: 12px;
        }

        #artworkSecondaryButton, #manualCutCancelButton {
            background-color: #111F35;
            color: #FFFFFF;
            border: none;
            border-radius: 8px;
            min-height: 34px;
            padding: 0 16px;
            font-weight: 700;
        }

        #artworkSecondaryButton:hover, #manualCutCancelButton:hover {
            background-color: #273B5A;
        }

        ArtworkManualCutDialog {
            background-color: #E5E5E5;
        }

        #manualCutTitle {
            color: #8A244B;
            font-size: 26px;
            font-weight: 800;
        }

        #manualCutDescription, #manualCutWarning {
            color: #5D6875;
            font-size: 13px;
        }

        #manualCutScroll {
            background-color: #E5E5E5;
            border: none;
        }

        #manualCutCard {
            background-color: transparent;
            border: none;
        }

        #manualCutPreview {
            background-color: #F2F2F2;
            border: 1px solid #D0D0D0;
            border-radius: 8px;
        }

        #cropControlPanel {
            background-color: #FFFFFF;
            border: 1px solid #C0C0C0;
            border-radius: 10px;
        }

        #cropTabs::pane {
            border: 1px solid #D0D0D0;
            background-color: #FFFFFF;
        }

        #cropTabs QTabBar::tab {
            background-color: #E5E5E5;
            color: #111F35;
            padding: 8px 10px;
            border: 1px solid #D0D0D0;
        }

        #cropTabs QTabBar::tab:selected {
            background-color: #8A244B;
            color: #FFFFFF;
        }

        #cropOptionCheck {
            color: #111F35;
            padding: 4px 0;
        }

        #cropOptionCheck::indicator {
            width: 16px;
            height: 16px;
            border: 1px solid #7D8694;
            border-radius: 4px;
            background-color: #FFFFFF;
        }

        #cropOptionCheck::indicator:checked {
            background-color: #E3262E;
            border-color: #A91920;
        }

        #manualCutZoomSlider::groove:horizontal {
            height: 4px;
            background-color: #C8C8C8;
        }

        #manualCutZoomSlider::handle:horizontal {
            width: 14px;
            margin: -5px 0;
            border-radius: 7px;
            background-color: #8A244B;
        }

        #manualCutFileName {
            color: #111F35;
            font-size: 15px;
            font-weight: 800;
        }

        #manualCutSectionLabel {
            color: #8A244B;
            font-size: 13px;
            font-weight: 700;
        }

        #manualCutRadio {
            color: #111F35;
            font-size: 13px;
            spacing: 8px;
            padding: 3px 0;
        }

        #manualCutRadio::indicator {
            width: 16px;
            height: 16px;
            border-radius: 4px;
            border: 1px solid #7D8694;
            background-color: #FFFFFF;
        }

        #manualCutRadio::indicator:checked {
            background-color: #B8F35A;
            border-color: #7A9E2A;
        }

        #searchExpandCheckbox {
            font-family: "Segoe UI";
            font-size: 13px;
            color: #333333;
            spacing: 8px;
        }

        /* Pill-shaped compact search input */
        #searchPillInput {
            background-color: #EFEFEF;
            color: #111111;
            border: none;
            border-radius: 23px;
            min-height: 46px;
            padding: 0 20px 0 20px;
            font-family: "Segoe UI";
            font-size: 15px;
        }
        #searchPillInput:focus {
            background-color: #E8E8E8;
        }

        /* Dark circular search button */
        #searchIconBtn {
            background-color: #1A1A1A;
            color: #FFFFFF;
            border: none;
            border-top-left-radius: 0px;
            border-bottom-left-radius: 0px;
            border-top-right-radius: 10px;
            border-bottom-right-radius: 10px;
            font-size: 20px;
            padding: 0;
            margin-left: 0px;
        }
        #searchIconBtn:hover  { background-color: #D02752; }
        #searchIconBtn:pressed { background-color: #111F35; }

        /* Expanded search – label cell (QLineEdit) and compact field dropdown (QComboBox) */
        #searchFieldLabel, QComboBox#searchFieldLabel {
            background-color: #E0E0E0;
            color: #111111;
            border: none;
            border-top-left-radius: 10px;
            border-bottom-left-radius: 10px;
            border-top-right-radius: 0px;
            border-bottom-right-radius: 0px;
            min-height: 40px;
            padding: 0 8px 0 12px;
            font-family: "Segoe UI";
            font-size: 13px;
            font-weight: 600;
        }
        QComboBox#searchFieldLabel::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: center right;
            width: 20px;
            border: none;
            background-color: transparent;
        }
        QComboBox#searchFieldLabel::down-arrow {
            image: none;
            width: 0;
            height: 0;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-top: 5px solid #555555;
            margin-right: 6px;
        }
        QComboBox#searchFieldLabel:hover { background-color: #D4D4D4; }
        QComboBox#searchFieldLabel QAbstractItemView {
            background-color: #FFFFFF;
            color: #111111;
            selection-background-color: #E8EEF8;
            selection-color: #111F35;
            border: 1px solid #BBBBBB;
        }

        /* Compact multi-value add button (SAP-style, inline) */
        #searchMultiAddBtn {
            background-color: #D8D8D8;
            color: #333333;
            border: none;
            border-top-left-radius: 0px;
            border-bottom-left-radius: 0px;
            border-top-right-radius: 0px;
            border-bottom-right-radius: 0px;
            font-size: 14px;
            padding: 0;
        }
        #searchMultiAddBtn:hover  { background-color: #C0C8D8; color: #111F35; }
        #searchMultiAddBtn:pressed { background-color: #A8B4C8; }

        /* Expanded search – value cell */
        #searchFieldInput {
            background-color: #F5F5F5;
            color: #111111;
            border: none;
            border-top-left-radius: 0px;
            border-bottom-left-radius: 0px;
            border-top-right-radius: 10px;
            border-bottom-right-radius: 10px;
            min-height: 40px;
            padding: 0 12px;
            font-family: "Segoe UI";
            font-size: 13px;
        }
        #searchFieldInput:focus {
            background-color: #EAEAEA;
            border: 2px solid rgba(208, 39, 82, 128);
        }

        #searchFieldInput:disabled {
            background-color: #E8E8E8;
            color: #AAAAAA;
        }

        /* Expanded search – combo cell */
        /* Dropdown label text */
        #searchDropdownLabel {
            font-family: "Segoe UI";
            font-size: 12px;
            color: #555555;
        }

        /* Standalone rounded-rectangle combo — closed state */
        QComboBox#searchDropdown {
            background-color: #FFFFFF;
            color: #111111;
            border: 1.5px solid #BBBBBB;
            border-radius: 10px;
            min-height: 40px;
            padding: 0 36px 0 14px;
            font-family: "Segoe UI";
            font-size: 13px;
        }
        QComboBox#searchDropdown:hover {
            border: 1.5px solid #111F35;
        }
        QComboBox#searchDropdown:on {
            border: 1.5px solid #111F35;
        }
        QComboBox#searchDropdown::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: center right;
            width: 32px;
            border: none;
            background-color: transparent;
        }
        QComboBox#searchDropdown::down-arrow {
            image: none;
            width: 0;
            height: 0;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 6px solid #555555;
            margin-right: 10px;
        }
        QComboBox#searchDropdown::up-arrow {
            image: none;
            width: 0;
            height: 0;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-bottom: 6px solid #111111;
            margin-right: 10px;
        }
        QComboBox#searchDropdown QAbstractItemView {
            background-color: #FFFFFF;
            border: 1px solid #CCCCCC;
            border-radius: 8px;
            outline: none;
            padding: 4px;
            font-family: "Segoe UI";
            font-size: 13px;
            color: #111111;
            selection-background-color: #EEEEEE;
            selection-color: #111111;
        }
        QComboBox#searchDropdown QAbstractItemView::item {
            padding: 6px 12px;
            border-radius: 6px;
            min-height: 28px;
        }
        QComboBox#searchDropdown QAbstractItemView::item:selected {
            background-color: #EEEEEE;
            font-weight: 700;
        }
        QComboBox#searchDropdown QAbstractItemView::item:hover {
            background-color: #F0F0F0;
        }

        /* Mini combos — Source, Asset Type, Sample Limit, Multiple Selection */
        QComboBox#searchDropdownMini {
            background-color: #FFFFFF;
            color: #111111;
            border: 1.5px solid #BBBBBB;
            border-radius: 6px;
            min-height: 22px;
            max-height: 22px;
            padding: 0 24px 0 8px;
            font-family: "Segoe UI";
            font-size: 10px;
        }
        QComboBox#searchDropdownMini:hover {
            border: 1.5px solid #111F35;
        }
        QComboBox#searchDropdownMini:on {
            border: 1.5px solid #111F35;
        }
        QComboBox#searchDropdownMini::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: center right;
            width: 20px;
            border: none;
            background-color: transparent;
        }
        QComboBox#searchDropdownMini::down-arrow {
            image: none;
            width: 0; height: 0;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-top: 5px solid #555555;
            margin-right: 6px;
        }
        QComboBox#searchDropdownMini::up-arrow {
            image: none;
            width: 0; height: 0;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-bottom: 5px solid #111111;
            margin-right: 6px;
        }
        QComboBox#searchDropdownMini QAbstractItemView {
            background-color: #FFFFFF;
            border: 1px solid #CCCCCC;
            border-radius: 6px;
            outline: none;
            padding: 2px;
            font-family: "Segoe UI";
            font-size: 10px;
            color: #111111;
            selection-background-color: #EEEEEE;
            selection-color: #111111;
        }
        QComboBox#searchDropdownMini QAbstractItemView::item {
            padding: 4px 8px;
            border-radius: 4px;
            min-height: 22px;
        }
        QComboBox#searchDropdownMini QAbstractItemView::item:selected {
            background-color: #EEEEEE;
            font-weight: 700;
        }
        QComboBox#searchDropdownMini QAbstractItemView::item:hover {
            background-color: #F0F0F0;
        }

        #headerBand {
            background-color: #000000;
            border: 1px solid #1F1F1F;
            border-radius: 12px;
        }

        #sectionTitle {
            color: #111F35;
            font-family: "Segoe UI";
            font-size: 21px;
            font-weight: 900;
            letter-spacing: 1px;
        }

        #projectSelectorBand {
            background-color: #9EA3AB;
            border: 1px solid #8B9098;
            border-radius: 8px;
        }

        #projectSelectorLabel {
            color: #000000;
            font-family: "Segoe UI";
            font-size: 13px;
            font-weight: 600;
            background: transparent;
        }

        #projectSelectorCombo {
            background-color: transparent;
            color: #000000;
            border: none;
            padding: 0 4px;
            font-family: "Segoe UI";
            font-size: 13px;
            font-weight: 700;
        }

        #projectSelectorCombo::drop-down {
            border: none;
            width: 22px;
        }

        #projectSelectorCombo QAbstractItemView {
            background-color: #9EA3AB;
            color: #000000;
            selection-background-color: #6B7280;
            selection-color: #FFFFFF;
            border: 1px solid #8B9098;
        }

        #basicToolActionBtn {
            background-color: #8A244B;
            color: #ffffff;
            border: none;
            border-radius: 16px;
            padding: 10px;
            text-align: center;
            font-family: "Roboto", "Segoe UI", Arial, sans-serif;
            font-size: 14px;
            font-weight: 700;
            letter-spacing: 0.2px;
        }

        #basicToolActionBtn:hover {
            background-color: #8A244B;
        }

        #basicToolActionBtn:checked {
            background-color: #D02752;
        }

        #basicToolActionBtn:pressed {
            background-color: #D02752;
        }

        #basicToolsRightPanel {
            background-color: #EDEDED;
            border: 1px solid #CDCDCD;
            border-radius: 18px;
        }

        #collectorTitle {
            color: #8A244B;
            font-family: "Segoe UI";
            font-size: 30px;
            font-weight: 800;
        }

        #collectorSectionLabel {
            color: #111F35;
            font-family: "Segoe UI";
            font-size: 15px;
            font-weight: 700;
        }

        #collectorDescription {
            color: #000000;
            font-family: "Segoe UI";
            font-size: 14px;
            font-weight: 500;
        }

        #thumbnailTitle {
            color: #8A244B;
            font-family: "Segoe UI";
            font-size: 30px;
            font-weight: 800;
        }

        #thumbnailDescription {
            color: #000000;
            font-family: "Segoe UI";
            font-size: 14px;
            font-weight: 500;
        }

        #packshotTitle {
            color: #8A244B;
            font-family: "Segoe UI";
            font-size: 30px;
            font-weight: 800;
        }

        #packshotDescription {
            color: #000000;
            font-family: "Segoe UI";
            font-size: 14px;
            font-weight: 500;
        }

        #mapperTitle {
            color: #8A244B;
            font-family: "Segoe UI";
            font-size: 30px;
            font-weight: 800;
        }

        #mapperDescription {
            color: #000000;
            font-family: "Segoe UI";
            font-size: 14px;
            font-weight: 500;
        }

        #collectorGrayBtn {
            background-color: #9EA3AB;
            color: #000000;
            border: 1px solid #8B9098;
            border-radius: 12px;
            padding: 0 14px;
            min-height: 42px;
            font-family: "Segoe UI";
            font-size: 14px;
            font-weight: 600;
        }

        #collectorGrayBtn:hover {
            background-color: #9EA3AB;
            color: #000000;
            border: 1px solid #8B9098;
        }

        #collectorGrayBtn:pressed {
            background-color: #111F35;
            color: #ffffff;
            border: 1px solid #111F35;
        }

        #collectorLineEdit {
            background-color: #D3D3D3;
            color: #000000;
            border: 1px solid #6F6F6F;
            border-radius: 16px;
            min-height: 40px;
            padding: 0 12px;
            font-family: "Segoe UI";
            font-size: 14px;
        }

        QLineEdit#comparePathLabel {
            background-color: #D3D3D3;
            color: #000000;
            border: 1px solid #6F6F6F;
            border-radius: 11px;
            padding: 0 8px;
            font-family: "Segoe UI";
            font-size: 10px;
        }

        QLineEdit#comparePathLabel:disabled {
            background-color: #E8E8E8;
            color: #aaaaaa;
            border: 1px solid #c0c0c0;
        }

        #compareSmallLabel {
            color: #000000;
            font-family: "Segoe UI";
            font-size: 11px;
        }

        QLineEdit#compareSmallInput {
            background-color: #D3D3D3;
            color: #000000;
            border: 1px solid #6F6F6F;
            border-radius: 6px;
            padding: 0 6px;
            font-family: "Segoe UI";
            font-size: 11px;
        }

        QLineEdit#compareSmallInput:disabled {
            background-color: #E8E8E8;
            color: #aaaaaa;
            border: 1px solid #c0c0c0;
        }

        #compareGrayBtn {
            background-color: #9EA3AB;
            color: #000000;
            border: 1px solid #8B9098;
            border-radius: 8px;
            padding: 0 10px;
            min-height: 34px;
            font-family: "Segoe UI";
            font-size: 11px;
            font-weight: 600;
        }

        #compareGrayBtn:hover { background-color: #ACB1B8; }

        #compareGrayBtn:pressed {
            background-color: #111F35;
            color: #ffffff;
            border: 1px solid #111F35;
        }

        #compareGrayBtn:disabled {
            background-color: #C8CBD0;
            color: #888888;
            border: 1px solid #B0B3B8;
        }

        #compareRunBtn {
            background-color: #111F35;
            color: #ffffff;
            border: none;
            border-radius: 8px;
            padding: 0 14px;
            font-family: "Segoe UI";
            font-size: 13px;
            font-weight: 700;
        }

        #compareRunBtn:pressed { background-color: #D02752; }

        #compareProgressBar {
            background-color: #C4C7CC;
            border: 1px solid #8D939D;
            border-radius: 5px;
        }

        #compareProgressBar::chunk {
            background-color: #111F35;
            border-radius: 5px;
            margin: 0px;
        }

        #collectorModeRadio {
            color: #8A244B;
            font-family: "Segoe UI";
            font-size: 16px;
            font-weight: 700;
            spacing: 8px;
        }

        #inputCountLabel {
            color: #8A244B;
            font-family: "Segoe UI";
            font-size: 13px;
            font-weight: 600;
        }

        #collectorModeRadio::indicator {
            width: 18px;
            height: 18px;
            border: 1px solid #8A244B;
            border-radius: 9px;
            background: transparent;
        }

        #collectorModeRadio::indicator:checked {
            background: #8A244B;
            border: 1px solid #8A244B;
        }

        #idhDupRadio {
            color: #8A244B;
            font-family: "Segoe UI";
            font-size: 13px;
            font-weight: 400;
            spacing: 8px;
        }

        #idhDupRadio::indicator {
            width: 14px;
            height: 14px;
            border: 1px solid #8A244B;
            border-radius: 7px;
            background: transparent;
        }

        #idhDupRadio::indicator:checked {
            background: #8A244B;
            border: 1px solid #8A244B;
        }

        #compareModeRadioSmall {
            color: #8A244B;
            font-family: "Segoe UI";
            font-size: 13px;
            font-weight: 400;
            spacing: 6px;
        }

        #compareModeRadioSmall::indicator {
            width: 14px;
            height: 14px;
            border: 1px solid #8A244B;
            border-radius: 7px;
            background: transparent;
        }

        #compareModeRadioSmall::indicator:checked {
            background: #8A244B;
            border: 1px solid #8A244B;
        }

        #compareModeRadioSmall:disabled {
            color: #aaaaaa;
        }

        #compareModeRadioSmall::indicator:disabled {
            border: 1px solid #aaaaaa;
        }

        #collectorCleanupRadio {
            color: #000000;
            font-family: "Segoe UI";
            font-size: 14px;
            font-weight: 500;
            spacing: 8px;
        }

        #collectorCleanupRadio::indicator {
            width: 18px;
            height: 18px;
            border: 1px solid #6F6F6F;
            border-radius: 9px;
            background: #D3D3D3;
        }

        #collectorCleanupRadio::indicator:checked {
            background: #111F35;
            border: 1px solid #111F35;
        }

        #mapperCleanupCheck {
            color: #000000;
            font-family: "Segoe UI";
            font-size: 14px;
            font-weight: 500;
            spacing: 8px;
        }

        #mapperCleanupCheck::indicator {
            width: 16px;
            height: 16px;
            border: 1px solid #6F6F6F;
            border-radius: 3px;
            background: #D3D3D3;
        }

        #mapperCleanupCheck::indicator:checked {
            background: #111F35;
            border: 1px solid #111F35;
        }

        #mapperCleanupCheck::indicator:disabled {
            background: #E8E8E8;
            border: 1px solid #BBBBBB;
        }

        #collectorCheck {
            color: #333333;
            font-family: "Segoe UI";
            font-size: 13px;
            spacing: 8px;
        }

        #collectorCheckSmall {
            color: #333333;
            font-family: "Segoe UI";
            font-size: 13px;
            spacing: 8px;
        }

        #collectorRunBtn {
            background-color: #111F35;
            color: #ffffff;
            border: none;
            border-radius: 12px;
            padding: 0 18px;
            font-family: "Segoe UI";
            font-size: 16px;
            font-weight: 700;
        }

        #collectorRunBtn:pressed {
            background-color: #D02752;
            color: #ffffff;
        }

        #collectorProgressBar {
            background-color: #C4C7CC;
            border: 1px solid #8D939D;
            border-radius: 7px;
        }

        #collectorProgressBar::chunk {
            background-color: #111F35;
            border-radius: 7px;
            margin: 0px;
        }

        #hatBadge {
            background: transparent;
            color: #ff1e1e;
            border: none;
            font-family: "Segoe UI";
            font-size: 40px;
            font-weight: 900;
            letter-spacing: 1px;
            padding: 0;
        }
        """
