import sys
import csv
import json
import os
import random
import cv2
import openpyxl

try:
    import pymupdf as fitz  # PyMuPDF, used to render Work Instruction PDFs
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QSlider, QTreeWidget, QTreeWidgetItem, QFileDialog,
    QLabel, QHeaderView, QComboBox, QGraphicsView, QGraphicsScene, 
    QGraphicsPixmapItem, QLineEdit, QTextEdit, QMessageBox, QAction,
    QAbstractItemView, QStyle, QToolBar, QStyledItemDelegate, QInputDialog, 
    QMenu, QProgressDialog, QStackedWidget, QSplitter, QListWidget, QListWidgetItem,
    QGroupBox, QSizePolicy, QButtonGroup, QScrollArea, QDialog, QCheckBox,
    QToolButton
)
from PyQt5.QtCore import Qt, QTimer, QEvent, QObject, QRect, QRectF, QPointF, QSettings
from PyQt5.QtGui import (
    QImage, QPixmap, QColor, QBrush, QFont, QPainter, QPen, QPolygonF,
    QFontMetrics
)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECTS_DIR = os.path.join(SCRIPT_DIR, "Projects")


# Google Sheets Dropdown Chip Color Palette
CATEGORY_COLORS = {
    "Tighten (Manual)": ("#FEE2E2", "#991B1B"),
    "Look for Tool": ("#FCE7F3", "#9D174D"),
    "Rework": ("#FFE4E6", "#9F1239"),
    "Pliers/Scissors": ("#FDE3E3", "#A82E2E"),
    "Tighten (Power)": ("#FFEDD5", "#9A3412"),
    "Fit/Place": ("#FFDAB9", "#8B4513"),
    "Grease/Loctite": ("#FEF3C7", "#92400E"),
    "Grab Parts": ("#FEF08A", "#854D0E"),
    "Other": ("#FFF4E6", "#996600"),
    "Read WI": ("#DBEAFE", "#1E3A8A"),
    "Setup": ("#E0F2FE", "#075985"),
    "Feed/Route": ("#CFFAFE", "#155E75"),
    "Write": ("#E6F0FF", "#003B99"),
    "Check": ("#ECFCCB", "#3F6212"),
    "Inspection": ("#DCFCE7", "#166534"),
    "Assembly": ("#D1FAE5", "#065F46"),
    "Connection": ("#CCFBF1", "#115E59"),
    "Clean": ("#E0FFE9", "#00662A"),
    "Flip/Rotate": ("#E0E7FF", "#3730A3"),
    "Peel/Stick": ("#EDE9FE", "#5B21B6"),
    "Fixturing/Clamping": ("#F3E8FF", "#6B21A8"),
    "Material Handling": ("#FAE8FF", "#86198F"),
    "Paperwork/Tulip": ("#E6E6FA", "#3D1466"),
    "Wait": ("#F5F5F5", "#404040"),
    "Break": ("#E2E8F0", "#0F172A"),
    "Delay": ("#E5E5E5", "#262626"),
    "De-Trash": ("#F4F4F5", "#18181B"),
    "Walk": ("#E7E5E4", "#292524"),
    "Place Screw": ("#FFEDD5", "#9A3412"),
    "Peel/Stick": ("#EDE9FE", "#5B21B6"),
    "Grab Parts From Kit": ("#FEF08A", "#854D0E"),
    "Grab Parts From Bin": ("#FEF08A", "#854D0E"),
    "Grab/Place Tool": ("#FCE7F3", "#9D174D"),
    "Swap Bit": ("#CCFBF1", "#115E59"),
    "Ergotranz": ("#E0F2FE", "#075985"),
    "Read WI": ("#E0E7FF", "#3730A3"),
}


def clean_slide_str(val):
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        if float(val).is_integer():
            return str(int(val))
        return str(val)
    s = str(val).strip()
    if s.endswith(".0"):
        try:
            return str(int(float(s)))
        except ValueError:
            pass
    return s


class CategoryDelegate(QStyledItemDelegate):
    def __init__(self, app, cat_type, parent=None):
        super().__init__(parent)
        self.app = app
        self.cat_type = cat_type

    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        text = index.data(Qt.DisplayRole)
        
        if text and text != "+ Add New Category...":
            rect = option.rect
            del_rect = QRect(rect.right() - 26, rect.top() + 4, 18, rect.height() - 8)
            painter.save()
            painter.setBrush(QBrush(QColor("#FDE2E2")))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(del_rect, 4, 4)
            painter.setPen(QColor("#9B1C1C"))
            font = painter.font()
            font.setBold(True)
            font.setPointSize(10)
            painter.setFont(font)
            painter.drawText(del_rect, Qt.AlignCenter, "X")
            
            edit_rect = QRect(rect.right() - 48, rect.top() + 4, 18, rect.height() - 8)
            painter.setBrush(QBrush(QColor("#E1F0FF")))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(edit_rect, 4, 4)
            painter.setPen(QColor("#1E429F"))
            painter.drawText(edit_rect, Qt.AlignCenter, "✎")
            
            painter.restore()


class ComboViewFilter(QObject):
    def __init__(self, app, combo, cat_type):
        super().__init__(combo)
        self.app = app
        self.combo = combo
        self.cat_type = cat_type

    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            pos = event.pos()
            index = self.combo.view().indexAt(pos)
            if index.isValid():
                text = index.data(Qt.DisplayRole)
                if text and text != "+ Add New Category...":
                    rect = self.combo.view().visualRect(index)
                    del_rect = QRect(rect.right() - 26, rect.top() + 4, 18, rect.height() - 8)
                    edit_rect = QRect(rect.right() - 48, rect.top() + 4, 18, rect.height() - 8)
                    
                    if del_rect.contains(pos):
                        self.combo.hidePopup()
                        QTimer.singleShot(0, lambda t=text, ct=self.cat_type: self.app.delete_category(t, ct))
                        return True
                    elif edit_rect.contains(pos):
                        self.combo.hidePopup()
                        QTimer.singleShot(0, lambda t=text, ct=self.cat_type: self.app.rename_category(t, ct))
                        return True
        return super().eventFilter(obj, event)


class NoScrollComboBox(QComboBox):
    def wheelEvent(self, event):
        event.ignore()


class ExcelDelegate(QStyledItemDelegate):
    def setEditorData(self, editor, index):
        super().setEditorData(editor, index)
        if isinstance(editor, QLineEdit):
            QTimer.singleShot(0, editor.selectAll)


class SelectAllFilter(QObject):
    def eventFilter(self, obj, event):
        if event.type() == QEvent.FocusIn:
            if isinstance(obj, QLineEdit):
                QTimer.singleShot(0, obj.selectAll)
        return super().eventFilter(obj, event)


class JumpSlider(QSlider):
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            val = QStyle.sliderValueFromPosition(
                self.minimum(), self.maximum(), event.x(), self.width()
            )
            self.setValue(val)
            self.sliderMoved.emit(val)
        super().mousePressEvent(event)


class ZoomableVideoView(QGraphicsView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.StrongFocus)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.pixmap_item = QGraphicsPixmapItem()
        self.scene.addItem(self.pixmap_item)

        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.zoom_factor = 1.15
        self._current_scale = 1.0
        self.speed_btn = None

    def set_speed_button(self, btn):
        self.speed_btn = btn
        self.speed_btn.setParent(self)
        self.update_button_position()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_button_position()

    def update_button_position(self):
        if self.speed_btn:
            margin = 15
            w = self.speed_btn.width()
            h = self.speed_btn.height()
            self.speed_btn.move(self.width() - w - margin, self.height() - h - margin)
            self.speed_btn.raise_()

    def set_frame(self, pixmap):
        self.pixmap_item.setPixmap(pixmap)
        self.scene.setSceneRect(self.pixmap_item.boundingRect())

    def wheelEvent(self, event):
        delta = event.angleDelta().y()
        if delta == 0:
            delta = event.angleDelta().x()
        if delta == 0:
            return

        if delta > 0:
            zoom = self.zoom_factor
        else:
            zoom = 1.0 / self.zoom_factor

        new_scale = self._current_scale * zoom
        if 0.5 <= new_scale <= 10.0:
            self._current_scale = new_scale
            self.scale(zoom, zoom)

    def reset_zoom(self):
        self.resetTransform()
        self._current_scale = 1.0
        if not self.pixmap_item.pixmap().isNull():
            self.fitInView(self.pixmap_item, Qt.KeepAspectRatio)


class MultiVideoTreeWidget(QTreeWidget):
    def __init__(self, app_instance, parent=None):
        super().__init__(parent)
        self.app = app_instance
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.SingleSelection)

        self.setEditTriggers(
            QAbstractItemView.DoubleClicked | 
            QAbstractItemView.SelectedClicked | 
            QAbstractItemView.EditKeyPressed | 
            QAbstractItemView.AnyKeyPressed
        )
        self.setItemDelegate(ExcelDelegate(self))

    def dropEvent(self, event):
        if self.app.view_mode != "video":
            event.ignore()
            return
            
        source_item = self.currentItem()
        target_item = self.itemAt(event.pos())
        event.setDropAction(Qt.IgnoreAction)
        event.accept()

        if (source_item and target_item and 
            source_item != target_item and 
            source_item.parent() is not None and 
            target_item.parent() is not None and 
            "END VIDEO" not in source_item.text(3) and 
            "END VIDEO" not in target_item.text(3)):
            
            self.app.swap_rows_content(source_item, target_item)


class TimeColumnWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.fill_ratio = 0.0
        self.setStyleSheet("background: transparent;")

    def set_fill_ratio(self, ratio):
        clamped_ratio = max(0.0, min(1.0, ratio))
        if self.fill_ratio != clamped_ratio:
            self.fill_ratio = clamped_ratio
            self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        if self.fill_ratio > 0:
            fill_width = int(self.width() * self.fill_ratio)
            painter.fillRect(0, 0, fill_width, self.height(), QColor(0, 0, 0, 35))
        painter.end()


class ParetoChartWidget(QWidget):
    """Pareto chart matching the Excel/Sheets combo-chart format.

    Bars = total delta per category (left axis, seconds), descending.
    Line = running cumulative percent of the charted total (right axis).
    """

    BAR_COLOR = "#4F81BD"       # theme accent1 used by the source workbook
    LINE_COLOR = "#C0504D"      # theme accent2
    GRID_COLOR = "#B7B7B7"
    TEXT_COLOR = "#000000"
    MAX_LABEL_BAND_FRACTION = 0.42  # cap on how much chart height the rotated labels may claim

    def __init__(self, parent=None):
        super().__init__(parent)
        self.data = []  # list of (category, total_seconds, cumulative_fraction)
        self.title = "Time Delay by Motion Category"
        self.x_title = "Motion Category"
        self.y_title = "Time Difference [mm:ss]"
        self.setMinimumHeight(420)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setAutoFillBackground(True)
        pal = self.palette()
        pal.setColor(self.backgroundRole(), QColor("#FFFFFF"))
        self.setPalette(pal)

    def set_data(self, data, title=None):
        self.data = list(data)
        if title:
            self.title = title
        self.update()

    @staticmethod
    def _nice_axis_max(value, tick_count=5):
        # Tick steps chosen so the mm:ss labels land on round time values.
        steps = [1, 2, 5, 10, 15, 20, 30, 60, 120, 300, 600, 900, 1200, 1800,
                 3600, 7200, 10800, 21600, 43200, 86400]
        if value <= 0:
            return float(tick_count), 1.0
        raw_step = value / tick_count
        step = next((s for s in steps if s >= raw_step), steps[-1])
        while step * tick_count < value:
            step += steps[-1]
        return float(step * tick_count), float(step)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#FFFFFF"))

        if not self.data:
            painter.setPen(QColor("#94A3B8"))
            painter.setFont(QFont("Segoe UI", 12))
            painter.drawText(self.rect(), Qt.AlignCenter,
                             "No category data to chart.\nLoad or open a project first.")
            return

        w, h = self.width(), self.height()
        title_font = QFont("Segoe UI", 13)
        axis_title_font = QFont("Segoe UI", 10, QFont.Bold)
        tick_font = QFont("Segoe UI", 9)

        # Title
        painter.setFont(title_font)
        painter.setPen(QColor(self.TEXT_COLOR))
        painter.drawText(QRectF(0, 8, w, 26), Qt.AlignCenter, self.title)

        # Reserve room for the rotated category labels along the bottom.
        fm_tick = QFontMetrics(tick_font)
        line_h = fm_tick.height()
        labels = [str(c) for c, _, _ in self.data]
        longest = max((fm_tick.width(t) for t in labels), default=0)
        label_band = min(longest + 12, int(h * self.MAX_LABEL_BAND_FRACTION))

        left = 78
        right = w - 82
        top = 46
        bottom = h - int(0.72 * label_band + line_h + 40)
        if right - left < 60 or bottom - top < 60:
            return
        plot = QRectF(left, top, right - left, bottom - top)

        max_val = max(v for _, v, _ in self.data)
        y_max, y_step = self._nice_axis_max(max_val)

        # Horizontal gridlines + left (seconds) and right (percent) tick labels
        painter.setFont(tick_font)
        tick_count = int(round(y_max / y_step))
        for i in range(tick_count + 1):
            frac = i / tick_count
            y = bottom - frac * plot.height()
            painter.setPen(QPen(QColor(self.GRID_COLOR), 1))
            painter.drawLine(QPointF(left, y), QPointF(right, y))
            painter.setPen(QColor(self.TEXT_COLOR))
            painter.drawText(QRectF(left - 72, y - 9, 66, 18),
                             Qt.AlignRight | Qt.AlignVCenter, self._fmt_seconds(y_step * i))
            painter.drawText(QRectF(right + 6, y - 9, 70, 18),
                             Qt.AlignLeft | Qt.AlignVCenter, f"{frac * 100:.2f}%")

        # Axis lines
        painter.setPen(QPen(QColor("#1A1A1A"), 1))
        painter.drawLine(QPointF(left, top), QPointF(left, bottom))
        painter.drawLine(QPointF(left, bottom), QPointF(right, bottom))
        painter.drawLine(QPointF(right, top), QPointF(right, bottom))

        # Bars
        n = len(self.data)
        slot = plot.width() / n
        bar_w = slot * 0.62
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(self.BAR_COLOR))
        for i, (_, val, _) in enumerate(self.data):
            bar_h = (val / y_max) * plot.height() if y_max else 0
            x = left + slot * i + (slot - bar_w) / 2
            painter.drawRect(QRectF(x, bottom - bar_h, bar_w, bar_h))

        # Cumulative percent line (right axis 0-100%)
        points = [QPointF(left + slot * (i + 0.5), bottom - cum * plot.height())
                  for i, (_, _, cum) in enumerate(self.data)]
        painter.setBrush(Qt.NoBrush)
        painter.setPen(QPen(QColor(self.LINE_COLOR), 2))
        painter.drawPolyline(QPolygonF(points))
        painter.setBrush(QColor(self.LINE_COLOR))
        painter.setPen(Qt.NoPen)
        for p in points:
            painter.drawEllipse(p, 3.2, 3.2)

        # Rotated category labels
        painter.setFont(tick_font)
        painter.setPen(QColor(self.TEXT_COLOR))
        for i, text in enumerate(labels):
            painter.save()
            painter.translate(left + slot * (i + 0.5), bottom + 8)
            painter.rotate(-45)
            painter.drawText(QRectF(-label_band - 10, -9, label_band + 4, 18),
                             Qt.AlignRight | Qt.AlignVCenter, text)
            painter.restore()

        # Axis titles
        painter.setFont(axis_title_font)
        painter.drawText(QRectF(left, h - 30, plot.width(), 22), Qt.AlignCenter, self.x_title)
        painter.save()
        painter.translate(18, (top + bottom) / 2)
        painter.rotate(-90)
        painter.drawText(QRectF(-plot.height() / 2, -10, plot.height(), 20),
                         Qt.AlignCenter, self.y_title)
        painter.restore()
        painter.save()
        painter.translate(w - 14, (top + bottom) / 2)
        painter.rotate(-90)
        painter.drawText(QRectF(-plot.height() / 2, -10, plot.height(), 20),
                         Qt.AlignCenter, "Running Percent")
        painter.restore()

    @staticmethod
    def _fmt_seconds(value):
        total = int(round(value))
        hours, rem = divmod(total, 3600)
        minutes, seconds = divmod(rem, 60)
        if hours:
            return f"{hours:d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"


class PdfCanvas(QWidget):
    """Continuous vertical strip of all PDF pages, with drag text selection."""

    SELECTION_COLOR = QColor(51, 133, 255, 90)

    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.setFocusPolicy(Qt.StrongFocus)
        self.setCursor(Qt.ArrowCursor)
        self.setMouseTracking(False)
        self.sel_anchor = None  # (page_index, word_index)
        self.sel_focus = None
        self.selection_armed = False  # text selection only starts after a double-click

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(event.rect(), QColor("#3A3A3A"))
        if not self.win.doc:
            return
        sel_lo, sel_hi = self.selection_range()
        for i, geom in enumerate(self.win.page_geoms):
            page_rect = QRect(int(geom["x"]), int(geom["y"]), int(geom["w"]), int(geom["h"]))
            if not page_rect.intersects(event.rect()):
                continue
            pixmap = self.win.page_pixmap(i)
            if pixmap is None:
                continue
            painter.drawPixmap(page_rect.topLeft(), pixmap)
            if sel_lo is not None and sel_lo[0] <= i <= sel_hi[0]:
                painter.setPen(Qt.NoPen)
                painter.setBrush(QBrush(self.SELECTION_COLOR))
                for w_idx, word in enumerate(self.win.page_words(i)):
                    if not self._word_selected(i, w_idx, sel_lo, sel_hi):
                        continue
                    s = geom["scale"]
                    painter.drawRect(QRectF(
                        geom["x"] + word[0] * s, geom["y"] + word[1] * s,
                        (word[2] - word[0]) * s, (word[3] - word[1]) * s
                    ))

    @staticmethod
    def _word_selected(page_idx, word_idx, lo, hi):
        return lo <= (page_idx, word_idx) <= hi

    def selection_range(self):
        if self.sel_anchor is None or self.sel_focus is None:
            return None, None
        lo, hi = sorted([self.sel_anchor, self.sel_focus])
        return lo, hi

    def clear_selection(self):
        self.sel_anchor = self.sel_focus = None
        self.selection_armed = False
        self.setCursor(Qt.ArrowCursor)
        self.update()

    def mousePressEvent(self, event):
        if event.button() != Qt.LeftButton:
            return
        if self.sel_anchor is not None or self.selection_armed:
            self.clear_selection()

    def mouseMoveEvent(self, event):
        if not self.selection_armed or not (event.buttons() & Qt.LeftButton) or self.sel_anchor is None:
            return
        hit = self.win.word_at(event.pos())
        if hit is not None and hit != self.sel_focus:
            self.sel_focus = hit
            self.update()

    def mouseDoubleClickEvent(self, event):
        if event.button() != Qt.LeftButton:
            return
        hit = self.win.word_at(event.pos())
        if hit is not None:
            self.selection_armed = True
            self.setCursor(Qt.IBeamCursor)
            self.sel_anchor = self.sel_focus = hit
            self.update()

    def keyPressEvent(self, event):
        if not self.win.handle_nav_key(event.key(), event.modifiers()):
            super().keyPressEvent(event)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        copy_action = menu.addAction("Copy")
        copy_action.setEnabled(bool(self.win.selected_text()))
        select_all_action = menu.addAction("Select All Text")
        chosen = menu.exec_(event.globalPos())
        if chosen == copy_action:
            self.win.copy_selection()
        elif chosen == select_all_action:
            self.win.select_all_text()


class WorkInstructionWindow(QWidget):
    """Pop-out window that streams a Work Instruction PDF like Chrome's PDF viewer."""

    PAGE_GAP = 14
    MARGIN = 12

    def __init__(self, parent=None):
        super().__init__(parent, Qt.Window)
        self.setWindowTitle("Work Instructions")
        self.resize(960, 760)
        self.setFocusPolicy(Qt.StrongFocus)

        self.doc = None
        self.pdf_path = ""
        self.user_zoom = 1.0
        self.page_geoms = []
        self._pixmap_cache = {}
        self._words_cache = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(False)
        self.scroll.setAlignment(Qt.AlignCenter)
        self.scroll.setFocusPolicy(Qt.NoFocus)  # keep key events on the window so arrows page slides
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)  # stable viewport width while laying out
        self.scroll.setStyleSheet("background-color: #3A3A3A; border: 1px solid #222222;")
        self.canvas = PdfCanvas(self)
        self.scroll.setWidget(self.canvas)
        layout.addWidget(self.scroll, 1)

        nav = QHBoxLayout()
        self.zoom_out_btn = QPushButton("−")
        self.zoom_out_btn.setFocusPolicy(Qt.NoFocus)
        self.zoom_out_btn.setFixedWidth(32)
        self.zoom_out_btn.setToolTip("Zoom out (Ctrl+Scroll / Ctrl+-)")
        self.zoom_out_btn.clicked.connect(lambda: self.adjust_zoom(1 / 1.15))
        nav.addWidget(self.zoom_out_btn)
        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setFocusPolicy(Qt.NoFocus)
        self.zoom_in_btn.setFixedWidth(32)
        self.zoom_in_btn.setToolTip("Zoom in (Ctrl+Scroll / Ctrl++)")
        self.zoom_in_btn.clicked.connect(lambda: self.adjust_zoom(1.15))
        nav.addWidget(self.zoom_in_btn)
        nav.addStretch()

        self.prev_btn = QPushButton("◄ Prev Slide")
        self.prev_btn.setFocusPolicy(Qt.NoFocus)
        self.prev_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.prev_btn.clicked.connect(lambda: self.step_page(-1))
        nav.addWidget(self.prev_btn)
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setMinimumWidth(150)
        self.status_label.setStyleSheet("font-weight: bold; color: #1E293B; font-size: 13px;")
        nav.addWidget(self.status_label)
        self.next_btn = QPushButton("Next Slide ►")
        self.next_btn.setFocusPolicy(Qt.NoFocus)
        self.next_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.next_btn.clicked.connect(lambda: self.step_page(1))
        nav.addWidget(self.next_btn)
        nav.addStretch()

        self.zoom_reset_btn = QPushButton("Fit")
        self.zoom_reset_btn.setFocusPolicy(Qt.NoFocus)
        self.zoom_reset_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.zoom_reset_btn.setToolTip("Reset zoom to fit width (Ctrl+0)")
        self.zoom_reset_btn.clicked.connect(self.reset_zoom)
        nav.addWidget(self.zoom_reset_btn)
        self.link_btn = QPushButton("Link Slide to Current Row")
        self.link_btn.setFocusPolicy(Qt.NoFocus)
        self.link_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.link_btn.setToolTip("Write this slide number into the Slide column of the selected row")
        self.link_btn.setStyleSheet("""
            QPushButton { background-color: #15803D; color: white; font-weight: bold; padding: 5px 12px; border-radius: 4px; }
            QPushButton:hover { background-color: #166534; }
        """)
        self.link_btn.clicked.connect(self.link_slide_to_row)
        nav.addWidget(self.link_btn)
        layout.addLayout(nav)

        self.scroll.viewport().installEventFilter(self)
        self.canvas.installEventFilter(self)
        self.scroll.verticalScrollBar().valueChanged.connect(self.update_status)

        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self.relayout)

    # ---------- document ----------

    def load_pdf(self, path):
        if fitz is None:
            raise RuntimeError("PyMuPDF (fitz) is not installed.")
        doc = fitz.open(path)
        if self.doc:
            self.doc.close()
        self.doc = doc
        self.pdf_path = path
        self.user_zoom = 1.0
        self.page_geoms = []
        self._pixmap_cache.clear()
        self._words_cache.clear()
        self.canvas.clear_selection()
        self.setWindowTitle(f"Work Instructions - {os.path.basename(path)}")
        self.relayout()
        self.scroll.verticalScrollBar().setValue(0)

    def page_count(self):
        return len(self.doc) if self.doc else 0

    def page_words(self, index):
        """Words as (x0, y0, x1, y1, text, block, line, word) in PDF points, reading order."""
        words = self._words_cache.get(index)
        if words is None:
            words = self.doc[index].get_text("words")
            self._words_cache[index] = words
        return words

    def page_pixmap(self, index):
        pixmap = self._pixmap_cache.get(index)
        if pixmap is None:
            geom = self.page_geoms[index]
            page = self.doc[index]
            pix = page.get_pixmap(matrix=fitz.Matrix(geom["scale"], geom["scale"]), alpha=False)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(img.copy())
            self._pixmap_cache[index] = pixmap
            self._prune_pixmap_cache(index)
        return pixmap

    def _prune_pixmap_cache(self, around):
        if len(self._pixmap_cache) <= 12:
            return
        for key in [k for k in self._pixmap_cache if abs(k - around) > 4]:
            del self._pixmap_cache[key]

    # ---------- layout ----------

    def relayout(self):
        self._pixmap_cache.clear()
        if not self.doc or self.page_count() == 0:
            self.canvas.resize(self.scroll.viewport().size())
            self.page_geoms = []
            self.update_status()
            return
        bar = self.scroll.verticalScrollBar()
        old_geoms = self.page_geoms
        anchor_page, anchor_ratio = 0, 0.0
        if old_geoms:
            anchor_page = self.current_page()
            g = old_geoms[anchor_page]
            anchor_ratio = (bar.value() - g["y"]) / g["h"] if g["h"] else 0.0
        viewport_w = max(220, self.scroll.viewport().width())
        target_w = max(120.0, (viewport_w - 2 * self.MARGIN) * self.user_zoom)
        content_w = max(viewport_w, target_w + 2 * self.MARGIN)
        geoms = []
        y = float(self.MARGIN)
        for i in range(self.page_count()):
            rect = self.doc[i].rect
            scale = target_w / rect.width if rect.width else 1.0
            h = rect.height * scale
            geoms.append({"x": (content_w - target_w) / 2.0, "y": y, "w": target_w, "h": h, "scale": scale})
            y += h + self.PAGE_GAP
        self.page_geoms = geoms
        self.canvas.resize(int(content_w), int(y - self.PAGE_GAP + self.MARGIN))
        if old_geoms:
            g = geoms[min(anchor_page, len(geoms) - 1)]
            bar.setValue(int(g["y"] + anchor_ratio * g["h"]))
        self.canvas.update()
        self.update_status()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.doc:
            self._resize_timer.start(120)

    # ---------- navigation ----------

    def current_page(self):
        if not self.page_geoms:
            return 0
        y = self.scroll.verticalScrollBar().value() + self.MARGIN + 2
        for i, geom in enumerate(self.page_geoms):
            if y < geom["y"] + geom["h"]:
                return i
        return len(self.page_geoms) - 1

    def goto_page(self, index):
        if not self.page_geoms:
            return
        index = max(0, min(int(index), self.page_count() - 1))
        self.scroll.verticalScrollBar().setValue(int(self.page_geoms[index]["y"] - self.MARGIN))
        self.update_status()

    def goto_slide(self, slide_value):
        """Slide numbers are 1-based and map directly onto PDF pages."""
        digits = "".join(ch for ch in str(slide_value) if ch.isdigit())
        if not digits:
            return
        self.goto_page(int(digits) - 1)

    def step_page(self, delta):
        if not self.page_geoms:
            return
        cur = self.current_page()
        if delta < 0:
            top = int(self.page_geoms[cur]["y"] - self.MARGIN)
            # Snap back to the top of the current slide first, like Chrome does.
            self.goto_page(cur if self.scroll.verticalScrollBar().value() > top + 4 else cur - 1)
        else:
            self.goto_page(cur + 1)

    def scroll_by(self, pixels):
        bar = self.scroll.verticalScrollBar()
        bar.setValue(bar.value() + pixels)

    def update_status(self):
        if not self.doc:
            self.status_label.setText("No Work Instruction loaded.")
            return
        zoom_txt = "" if abs(self.user_zoom - 1.0) < 1e-6 else f"  ({self.user_zoom * 100:.0f}%)"
        self.status_label.setText(f"Slide {self.current_page() + 1} / {self.page_count()}{zoom_txt}")

    # ---------- zoom ----------

    def adjust_zoom(self, factor):
        new_zoom = max(0.25, min(6.0, self.user_zoom * factor))
        if abs(new_zoom - self.user_zoom) < 1e-6:
            return
        self.user_zoom = new_zoom
        self.relayout()

    def reset_zoom(self):
        if abs(self.user_zoom - 1.0) < 1e-6:
            return
        self.user_zoom = 1.0
        self.relayout()

    # ---------- text selection ----------

    def word_at(self, pos):
        """Maps a canvas point to (page_index, word_index), snapping to the nearest word."""
        if not self.page_geoms:
            return None
        page_idx = None
        for i, geom in enumerate(self.page_geoms):
            if pos.y() < geom["y"] + geom["h"] + self.PAGE_GAP / 2:
                page_idx = i
                break
        if page_idx is None:
            page_idx = len(self.page_geoms) - 1
        geom = self.page_geoms[page_idx]
        scale = geom["scale"] or 1.0
        x = (pos.x() - geom["x"]) / scale
        y = (pos.y() - geom["y"]) / scale
        words = self.page_words(page_idx)
        if not words:
            return None
        best_idx, best_dist = 0, None
        for idx, w in enumerate(words):
            if w[0] <= x <= w[2] and w[1] <= y <= w[3]:
                return (page_idx, idx)
            dx = max(w[0] - x, 0, x - w[2])
            dy = max(w[1] - y, 0, y - w[3])
            dist = dx * dx + (dy * 4) ** 2  # weight vertical distance so lines win over columns
            if best_dist is None or dist < best_dist:
                best_idx, best_dist = idx, dist
        return (page_idx, best_idx)

    def select_all_text(self):
        if not self.doc or self.page_count() == 0:
            return
        last_page = self.page_count() - 1
        last_words = self.page_words(last_page)
        self.canvas.selection_armed = True
        self.canvas.setCursor(Qt.IBeamCursor)
        self.canvas.sel_anchor = (0, 0)
        self.canvas.sel_focus = (last_page, max(0, len(last_words) - 1))
        self.canvas.update()

    def selected_text(self):
        lo, hi = self.canvas.selection_range()
        if lo is None:
            return ""
        chunks = []
        prev_line = None
        for page_idx in range(lo[0], hi[0] + 1):
            words = self.page_words(page_idx)
            start = lo[1] if page_idx == lo[0] else 0
            end = hi[1] if page_idx == hi[0] else len(words) - 1
            for idx in range(start, min(end, len(words) - 1) + 1):
                w = words[idx]
                line_key = (page_idx, w[5], w[6])
                if prev_line is not None and line_key != prev_line:
                    chunks.append("\n")
                elif chunks:
                    chunks.append(" ")
                chunks.append(w[4])
                prev_line = line_key
        return "".join(chunks).strip()

    def copy_selection(self):
        text = self.selected_text()
        if text:
            QApplication.clipboard().setText(text)

    # ---------- input ----------

    def eventFilter(self, obj, event):
        if obj in (self.scroll.viewport(), self.canvas) and self.doc:
            if event.type() == QEvent.Wheel:
                delta = event.angleDelta().y() or event.angleDelta().x()
                if event.modifiers() & Qt.ControlModifier:
                    if delta:
                        self.adjust_zoom(1.15 if delta > 0 else 1 / 1.15)
                    return True
                if event.modifiers() & Qt.ShiftModifier:
                    if delta:
                        bar = self.scroll.horizontalScrollBar()
                        bar.setValue(bar.value() - delta)
                    return True
            if event.type() == QEvent.KeyPress and self.handle_nav_key(event.key(), event.modifiers()):
                return True
        return super().eventFilter(obj, event)

    def handle_nav_key(self, key, modifiers=Qt.NoModifier):
        if modifiers & Qt.ControlModifier:
            if key in (Qt.Key_Plus, Qt.Key_Equal):
                self.adjust_zoom(1.15)
            elif key in (Qt.Key_Minus, Qt.Key_Underscore):
                self.adjust_zoom(1 / 1.15)
            elif key == Qt.Key_0:
                self.reset_zoom()
            elif key == Qt.Key_C:
                self.copy_selection()
            elif key == Qt.Key_A:
                self.select_all_text()
            else:
                return False
            return True
        if key in (Qt.Key_Left, Qt.Key_Backspace):
            self.step_page(-1)
        elif key in (Qt.Key_Right, Qt.Key_Space):
            self.step_page(1)
        elif key == Qt.Key_Up:
            self.scroll_by(-60)
        elif key == Qt.Key_Down:
            self.scroll_by(60)
        elif key == Qt.Key_PageUp:
            self.scroll_by(-self.scroll.viewport().height())
        elif key == Qt.Key_PageDown:
            self.scroll_by(self.scroll.viewport().height())
        elif key == Qt.Key_Home:
            self.goto_page(0)
        elif key == Qt.Key_End:
            self.goto_page(self.page_count() - 1)
        elif key == Qt.Key_Escape:
            self.canvas.clear_selection()
        else:
            return False
        return True

    def keyPressEvent(self, event):
        if not self.handle_nav_key(event.key(), event.modifiers()):
            super().keyPressEvent(event)

    def focus_viewer(self):
        self.canvas.setFocus(Qt.OtherFocusReason)

    def link_slide_to_row(self):
        owner = self.parent()
        if self.doc and hasattr(owner, "link_slide_to_selected_row"):
            owner.link_slide_to_selected_row(self.current_page() + 1)

    def closeEvent(self, event):
        self._pixmap_cache.clear()
        self._words_cache.clear()
        self.page_geoms = []
        if self.doc:
            self.doc.close()
            self.doc = None
        super().closeEvent(event)


class SettingsDialog(QDialog):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.setWindowTitle("Settings")
        self.setMinimumSize(760, 480)
        self.resize(860, 540)
        self.setStyleSheet("""
            QDialog { background-color: #202328; color: #E6E8EB; }
            QListWidget { background-color: #17191D; border: none; color: #C7CBD1; font-size: 14px; padding-top: 10px; }
            QListWidget::item { padding: 11px 16px; border-left: 3px solid transparent; }
            QListWidget::item:selected { background-color: #2B3038; color: #FFFFFF; border-left: 3px solid #2D8CEB; }
            QLabel#settingsTitle { color: #FFFFFF; font-size: 22px; font-weight: bold; }
            QLabel#settingsSection { color: #AEB4BD; font-size: 12px; font-weight: bold; }
            QLabel#settingsName { color: #F2F3F5; font-size: 14px; font-weight: bold; }
            QLabel#settingsDescription { color: #9DA3AC; font-size: 12px; }
            QCheckBox { spacing: 8px; }
            QCheckBox::indicator { width: 38px; height: 20px; }
            QCheckBox::indicator:unchecked { background-color: #464B53; border: 1px solid #656B74; border-radius: 10px; }
            QCheckBox::indicator:checked { background-color: #2D8CEB; border: 1px solid #58A8F5; border-radius: 10px; }
            QPushButton { background-color: #30343B; color: #E6E8EB; border: 1px solid #484D55; padding: 7px 18px; border-radius: 3px; }
            QPushButton:hover { background-color: #3A3F47; }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        categories = QListWidget()
        categories.setFixedWidth(215)
        categories.addItem("WI pop-out")
        categories.setCurrentRow(0)
        layout.addWidget(categories)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(32, 26, 32, 24)
        content_layout.setSpacing(18)

        title = QLabel("WI pop-out")
        title.setObjectName("settingsTitle")
        content_layout.addWidget(title)

        section = QLabel("BEHAVIOR")
        section.setObjectName("settingsSection")
        content_layout.addWidget(section)

        self.auto_jump_checkbox = self._add_setting(
            content_layout,
            "Auto-jump to slide",
            "Jump the WI PDF viewer to the slide number of the selected row.",
            app.auto_jump_to_slide,
            app.set_auto_jump_to_slide,
        )
        self.auto_open_checkbox = self._add_setting(
            content_layout,
            "Auto-open WI on startup",
            "Open the saved WI pop-out automatically when a project is loaded.",
            app.auto_open_wi_on_startup,
            app.set_auto_open_wi_on_startup,
        )

        content_layout.addStretch()
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)
        content_layout.addLayout(button_layout)
        layout.addWidget(content, 1)

    def _add_setting(self, layout, name, description, checked, callback):
        row = QHBoxLayout()
        text_layout = QVBoxLayout()
        text_layout.setSpacing(3)
        name_label = QLabel(name)
        name_label.setObjectName("settingsName")
        description_label = QLabel(description)
        description_label.setObjectName("settingsDescription")
        description_label.setWordWrap(True)
        text_layout.addWidget(name_label)
        text_layout.addWidget(description_label)
        row.addLayout(text_layout, 1)
        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.toggled.connect(callback)
        row.addWidget(checkbox, 0, Qt.AlignVCenter)
        layout.addLayout(row)
        return checkbox


class TimeStudyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Multi-Video Time Study Logger - [Untitled Project]")
        self.setGeometry(100, 100, 1420, 800)

        self.settings = QSettings("TimeStudyTool", "TimeStudyTool")
        self.auto_jump_to_slide = self.settings.value("wi/auto_jump_to_slide", True, type=bool)
        self.auto_open_wi_on_startup = self.settings.value("wi/auto_open_on_startup", True, type=bool)

        self.unsaved_changes = False
        self.view_mode = "video"
        self.saved_video_state = []
        self.active_category_item = None
        # Last (group index, row index) selected in each tree, so mode switches return to the same row.
        self.last_video_selection = None
        self.last_category_selection = None
        # Mirrors the source spreadsheet's Pareto query, which drops Break and Other.
        self.pareto_excluded_cats = {"Break", "Other"}

        self.cap = None
        self.active_video_path = ""
        self.active_group_item = None
        self.project_path = ""
        self.wi_path = ""
        self.wi_window = None
        self.is_playing = False
        self.fps = 30.0
        self.total_frames = 0
        self.duration_ms = 0
        self.playback_speed = 1.0

        self.custom_gen_cats = set()
        self.custom_spec_cats = set()
        self.custom_category_colors = {}
        self.custom_color_hue_offset = random.randrange(24)
        self.time_editing_item = None

        self.undo_stack = []
        self.redo_stack = []
        self.is_restoring_state = False
        self.select_all_filter = SelectAllFilter(self)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.next_frame)

        self.default_cat_general_options = [
            "Assembly", "Inspection", "Rework", "Delay", "Setup", "Material Handling"
        ]
        self.default_cat_specific_options = [
            "Tighten (Power)", "Fit/Place", "Place Screw", "Walk", "Peel/Stick",
            "De-Trash", "Break", "Tighten (Manual)", "Paperwork/Tulip", "Grab Parts From Kit",
            "Grab Parts From Bin", "Grab/Place Tool", "Check", "Fixturing/Clamping", "Feed/Route",
            "Other", "Swap Bit", "Ergotranz", "Connection", "Grease/Loctite", "Inspection",
            "Pliers/Scissors", "Wait", "Write", "Clean", "Flip/Rotate", "Read WI"
        ]

        self.cat_general_options = self.default_cat_general_options.copy()
        self.cat_specific_options = self.default_cat_specific_options.copy()
        self.cat_general_options.sort()
        
        self.cat_specific_options.sort(key=lambda x: (
            self.default_cat_specific_options.index(x) if x in self.default_cat_specific_options else 999,
            x.lower()
        ))

        self.init_ui()
        self.create_menu_bar()
        QApplication.instance().installEventFilter(self)
        
        # Add to the end of def __init__(self):
        self.seek_timer = QTimer(self)
        self.seek_timer.setSingleShot(True)
        self.seek_timer.timeout.connect(self._perform_seek)
        self.target_seek_ms = 0

    def format_ms(self, ms):
        """Formats milliseconds into MM:SS or HH:MM:SS."""
        if not isinstance(ms, (int, float)): return "00:00"
        total_seconds = int(ms) // 1000
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

    def next_frame(self):
        """Advances the video by a single frame and updates the UI."""
        if self.cap and self.cap.isOpened() and self.is_playing:
            
            # --- AUTO-ADVANCE LOGIC FOR CATEGORY MODE ---
            if self.view_mode == "category" and getattr(self, 'active_category_item', None):
                current_ms = int(self.cap.get(cv2.CAP_PROP_POS_MSEC))
                end_ms = self.active_category_item.data(2, Qt.UserRole) or 0
                
                # If we've reached or passed the segment's end time
                if current_ms >= end_ms and end_ms > 0:
                    slice_index = getattr(self.active_category_item, "category_slice_index", 0)
                    slices = getattr(self.active_category_item, "category_slices", [])
                    if slice_index + 1 < len(slices):
                        self._play_category_slice(self.active_category_item, slice_index + 1)
                        return
                    next_item = self.get_next_category_segment(self.active_category_item)
                    if next_item:
                        self.play_category_segment(next_item) # Auto-play next clip
                        return 
                    else:
                        self.toggle_play() # Pause at the end of the category
                        return
            # --------------------------------------------
            
            # Skip frames to match playback speed without taxing the CPU
            frames_to_skip = int(self.playback_speed) - 1
            for _ in range(frames_to_skip):
                self.cap.grab()
            
            ret, frame = self.cap.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                h, w, ch = frame.shape
                bytes_per_line = ch * w
                q_img = QImage(frame.data, w, h, bytes_per_line, QImage.Format_RGB888)
                self.video_view.set_frame(QPixmap.fromImage(q_img))
                
                current_ms = int(self.cap.get(cv2.CAP_PROP_POS_MSEC))
                self.slider.blockSignals(True)
                self.slider.setValue(current_ms)
                self.slider.blockSignals(False)
                
                self.time_label.setText(f"{self.format_ms(current_ms)} / {self.format_ms(self.duration_ms)} (◄ / ► Arrow Keys = ±1s)")
                self.refresh_playback_ui()
            else:
                self.toggle_play() # Pause at physical end of video

    def toggle_play(self):
        """Starts or pauses the video playback timer."""
        if self.cap and self.cap.isOpened():
            self.is_playing = not self.is_playing
            if self.is_playing:
                self.timer.start()
            else:
                self.timer.stop()

    def toggle_speed(self):
        """Cycles through playback speed multipliers."""
        speeds = [1.0, 2.0, 4.0, 8.0]
        try:
            idx = speeds.index(self.playback_speed)
            self.set_playback_speed(speeds[(idx + 1) % len(speeds)])
        except ValueError:
            self.set_playback_speed(1.0)

    def set_playback_speed(self, speed):
        """Sets playback speed and updates its on-video control."""
        self.playback_speed = speed
        self.speed_btn.setText(f"{int(self.playback_speed)}x")
        if self.fps > 0:
            # Keep interval locked to the natural FPS
            interval = max(1, int(1000 / self.fps))
            self.timer.setInterval(interval)

    def _default_browse_dir(self):
        if self.project_path:
            return os.path.dirname(self.project_path)
        if os.path.isdir(PROJECTS_DIR):
            return PROJECTS_DIR
        return ""

    def open_file_dialog(self):
        """Triggers the video loading dialog."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Video File", self._default_browse_dir(), "Video Files (*.mp4 *.avi *.mkv *.mov)")
        if file_path:
            self.add_video_group(file_path)

    def seek_position(self, ms):
        """Jumps to a specific timestamp in the video, correcting for keyframe-snap seek imprecision."""
        if self.cap and self.cap.isOpened():
            target_ms = max(0, ms)
            self.cap.set(cv2.CAP_PROP_POS_MSEC, target_ms)
            ret, frame = self.cap.read()
            if ret:
                # Sparse-keyframe sources (e.g. GoPro proxy files) can land the decoder
                # well before the requested time; step forward until we reach it for real.
                actual_ms = self.cap.get(cv2.CAP_PROP_POS_MSEC)
                frame_duration_ms = (1000.0 / self.fps) if self.fps else 33.0
                guard = 0
                while actual_ms < target_ms - frame_duration_ms and guard < 300:
                    ret2, next_frame = self.cap.read()
                    if not ret2:
                        break
                    frame = next_frame
                    actual_ms = self.cap.get(cv2.CAP_PROP_POS_MSEC)
                    guard += 1

                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                h, w, ch = frame.shape
                q_img = QImage(frame.data, w, h, ch * w, QImage.Format_RGB888)
                self.video_view.set_frame(QPixmap.fromImage(q_img))

                display_ms = int(actual_ms)
                self.slider.blockSignals(True)
                self.slider.setValue(display_ms)
                self.slider.blockSignals(False)

                self.time_label.setText(f"{self.format_ms(display_ms)} / {self.format_ms(self.duration_ms)} (◄ / ► Arrow Keys = ±1s)")
                self.refresh_playback_ui()

    def step_time(self, delta_ms):
        """Nudges the video forward or backwards (arrow keys) with UI debouncing."""
        if self.cap and self.cap.isOpened():
            # Base the jump on the target if we are already scrubbing, else use current position
            base_ms = self.target_seek_ms if self.seek_timer.isActive() else self.cap.get(cv2.CAP_PROP_POS_MSEC)
            self.target_seek_ms = max(0, min(self.duration_ms, base_ms + delta_ms))
            
            # Update the UI immediately so it feels snappy to the user
            self.slider.blockSignals(True)
            self.slider.setValue(int(self.target_seek_ms))
            self.slider.blockSignals(False)
            self.time_label.setText(f"{self.format_ms(self.target_seek_ms)} / {self.format_ms(self.duration_ms)} (Scrubbing...)")
            
            # Delay the actual heavy OpenCV seek by 100ms
            self.seek_timer.start(10)

    def _perform_seek(self):
        """Executes the delayed seek."""
        self.seek_position(self.target_seek_ms)

    def _segment_boundary_epsilon_ms(self):
        """~1 frame of tolerance so fps-quantized frame timestamps don't misclassify a row boundary,
        without being large enough to noticeably pre-empt row switching during normal playback."""
        frame_duration_ms = (1000.0 / self.fps) if self.fps else 33.0
        return frame_duration_ms + 10

    def skip_to_segment_start(self):
        """Ctrl+Left: jump to the start of the current segment. If already there, jump to the segment immediately prior."""
        if self.view_mode != "video" or not self.cap or not self.cap.isOpened() or not self.active_group_item:
            return
        current_ms = int(self.target_seek_ms if self.seek_timer.isActive() else self.cap.get(cv2.CAP_PROP_POS_MSEC))
        group = self.active_group_item
        child_count = group.childCount()
        # Seeks can land a few ms before their target (fps-quantized frame timestamps),
        # so nudge forward before matching against row boundaries to avoid misclassifying
        # the segment we're actually in.
        query_ms = current_ms + self._segment_boundary_epsilon_ms()
        segment_index = 0
        segment_start_ms = 0
        for i in range(max(0, child_count - 1)):
            item = group.child(i)
            next_item = group.child(i + 1)
            ms1 = item.data(4, Qt.UserRole) or 0
            ms2 = next_item.data(4, Qt.UserRole) or 0
            if ms1 <= query_ms < ms2:
                segment_index = i
                segment_start_ms = ms1
                break
        else:
            if child_count >= 2:
                last_real_index = child_count - 2
                last_ms = group.child(last_real_index).data(4, Qt.UserRole) or 0
                if query_ms >= last_ms:
                    segment_index = last_real_index
                    segment_start_ms = last_ms

        if current_ms - segment_start_ms > 300:
            self.seek_position(segment_start_ms)
        elif segment_index > 0:
            prev_segment_start_ms = group.child(segment_index - 1).data(4, Qt.UserRole) or 0
            self.seek_position(prev_segment_start_ms)
        else:
            idx = self.video_tree.indexOfTopLevelItem(group)
            total_groups = self.video_tree.topLevelItemCount()
            # Wrap around to the last video in the project if we're already on the first one
            prev_idx = idx - 1 if idx > 0 else total_groups - 1
            prev_group = self.video_tree.topLevelItem(prev_idx)
            prev_child_count = prev_group.childCount()
            # Last real segment start (excludes the END VIDEO sentinel row)
            last_segment_start_ms = prev_group.child(prev_child_count - 2).data(4, Qt.UserRole) or 0 if prev_child_count >= 2 else 0
            self.switch_active_video(prev_group, last_segment_start_ms)

    def skip_to_next_segment_start(self):
        """Ctrl+Right: jump to the start of the next segment. If there is none, jump to the start of the next video."""
        if self.view_mode != "video" or not self.cap or not self.cap.isOpened() or not self.active_group_item:
            return
        current_ms = int(self.target_seek_ms if self.seek_timer.isActive() else self.cap.get(cv2.CAP_PROP_POS_MSEC))
        group = self.active_group_item
        child_count = group.childCount()
        # See skip_to_segment_start for why we nudge forward before matching boundaries.
        query_ms = current_ms + self._segment_boundary_epsilon_ms()
        segment_index = 0
        for i in range(max(0, child_count - 1)):
            item = group.child(i)
            next_item = group.child(i + 1)
            ms1 = item.data(4, Qt.UserRole) or 0
            ms2 = next_item.data(4, Qt.UserRole) or 0
            if ms1 <= query_ms < ms2:
                segment_index = i
                break
        else:
            if child_count >= 2:
                segment_index = child_count - 2

        # Always step to the immediately following row, never skipping closely-spaced segments
        next_real_index = segment_index + 1
        if next_real_index <= child_count - 2:
            next_start_ms = group.child(next_real_index).data(4, Qt.UserRole) or 0
            self.seek_position(next_start_ms)
        else:
            idx = self.video_tree.indexOfTopLevelItem(group)
            total_groups = self.video_tree.topLevelItemCount()
            # Wrap around to the first video in the project if we're already on the last one
            next_idx = idx + 1 if idx + 1 < total_groups else 0
            self.switch_active_video(self.video_tree.topLevelItem(next_idx), 0)

    def skip_to_category_segment_start(self):
        """Ctrl+Left: restart the current category slice, then move to the preceding category segment."""
        item = self.active_category_item
        if not item or not self.cap or not self.cap.isOpened():
            return

        slices = getattr(item, "category_slices", [])
        slice_index = getattr(item, "category_slice_index", 0)
        start_ms = slices[slice_index]["start_ms"] if slices else (item.data(1, Qt.UserRole) or 0)
        current_ms = self.cap.get(cv2.CAP_PROP_POS_MSEC)
        if current_ms - start_ms > self._segment_boundary_epsilon_ms():
            self.seek_position(start_ms)
            return

        previous_item = self.get_previous_category_segment(item)
        if previous_item:
            self.play_category_segment(previous_item)

    def skip_to_next_category_segment_start(self):
        """Ctrl+Right: start the following displayed category segment."""
        item = self.active_category_item
        if not item:
            return
        next_item = self.get_next_category_segment(item)
        if next_item:
            self.play_category_segment(next_item)
        
    def _clear_custom_category_colors(self):
        for cat_name in self.custom_category_colors:
            CATEGORY_COLORS.pop(cat_name, None)
        self.custom_category_colors.clear()

    def _category_color_distance(self, first_color, second_color):
        first = QColor(first_color)
        second = QColor(second_color)
        first_rgb = (first.red(), first.green(), first.blue())
        second_rgb = (second.red(), second.green(), second.blue())
        return sum((first_rgb[i] - second_rgb[i]) ** 2 for i in range(3)) ** 0.5

    def _generate_custom_category_color(self, cat_name):
        custom_colors = [colors[0] for colors in self.custom_category_colors.values() if colors[0] != CATEGORY_COLORS.get(cat_name, (None,))[0]]
        default_colors = [colors[0] for name, colors in CATEGORY_COLORS.items() if name not in self.custom_category_colors]
        custom_hues = [QColor(color).hue() for color in custom_colors if QColor(color).hue() >= 0]
        candidates = []
        candidate_hues = [(self.custom_color_hue_offset + (index * 24)) % 360 for index in range(15)]
        random.shuffle(candidate_hues)
        for hue in candidate_hues:
            background = QColor.fromHsv(hue, 75, 245).name().upper()
            hue_distance = min(
                (min(abs(hue - existing), 360 - abs(hue - existing)) for existing in custom_hues),
                default=float("inf")
            )
            custom_distance = min(
                (self._category_color_distance(background, existing) for existing in custom_colors),
                default=float("inf")
            )
            default_distance = min(
                (self._category_color_distance(background, existing) for existing in default_colors),
                default=float("inf")
            )
            candidates.append((hue_distance, custom_distance, default_distance, background))

        background = max(candidates, key=lambda candidate: (candidate[0], candidate[1], candidate[2]))[3]
        bg_color = QColor(background)
        foreground = QColor.fromHsv(bg_color.hue(), 220, 95).name().upper()
        colors = (background, foreground)
        self.custom_category_colors[cat_name] = colors
        CATEGORY_COLORS[cat_name] = colors
        return colors

    def _ensure_category_color(self, cat_name):
        if cat_name and cat_name not in CATEGORY_COLORS:
            self._generate_custom_category_color(cat_name)

    def _load_saved_category_colors(self, saved_colors):
        for cat_name, color_data in (saved_colors or {}).items():
            if not isinstance(color_data, dict):
                continue
            background = QColor(color_data.get("background", ""))
            foreground = QColor(color_data.get("foreground", ""))
            if background.isValid() and foreground.isValid():
                colors = (background.name().upper(), foreground.name().upper())
                self.custom_category_colors[cat_name] = colors
                CATEGORY_COLORS[cat_name] = colors

    def _load_custom_categories(self, gen_cats, spec_cats):
        for cat in (gen_cats or []):
            if cat and cat not in self.cat_general_options:
                self.cat_general_options.append(cat)
                self.custom_gen_cats.add(cat)
                self._ensure_category_color(cat)
                    
        for cat in (spec_cats or []):
            if cat and cat not in self.cat_specific_options:
                self.cat_specific_options.append(cat)
                self.custom_spec_cats.add(cat)
                self._ensure_category_color(cat)

        self.cat_general_options.sort()
        self.cat_specific_options.sort(key=lambda x: (
            self.default_cat_specific_options.index(x) if x in self.default_cat_specific_options else 999,
            x.lower()
        ))

    def _scan_and_add_custom_categories(self, groups_data):
        for g_data in groups_data:
            for row in g_data.get("rows", []):
                cat_gen = row.get("cat_gen", "")
                cat_spec = row.get("cat_spec", "")
                
                if cat_gen and cat_gen not in self.cat_general_options:
                    self.cat_general_options.append(cat_gen)
                    self.custom_gen_cats.add(cat_gen)
                    self._ensure_category_color(cat_gen)
                        
                if cat_spec and cat_spec not in self.cat_specific_options:
                    self.cat_specific_options.append(cat_spec)
                    self.custom_spec_cats.add(cat_spec)
                    self._ensure_category_color(cat_spec)

        self.cat_general_options.sort()
        self.cat_specific_options.sort(key=lambda x: (
            self.default_cat_specific_options.index(x) if x in self.default_cat_specific_options else 999,
            x.lower()
        ))

    def parse_time_ms(self, ms_val, time_str):
        if time_str:
            try:
                parts = [int(p) for p in str(time_str).strip().split(':')]
                if len(parts) == 3:
                    return ((parts[0] * 3600) + (parts[1] * 60) + parts[2]) * 1000
                elif len(parts) == 2:
                    return ((parts[0] * 60) + parts[1]) * 1000
                elif len(parts) == 1:
                    return parts[0] * 1000
            except (ValueError, TypeError):
                pass
        if isinstance(ms_val, (int, float)) and ms_val > 0:
            return int(ms_val if ms_val > 1000 else ms_val * 1000)
        return 0

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress:
            # The WI pop-out owns its own arrow-key navigation while it is focused.
            if self.wi_window is not None and self._event_belongs_to_wi(obj):
                return super().eventFilter(obj, event)
            focused = QApplication.focusWidget()
            key = event.key()

            is_time_edit_focused = False
            if isinstance(focused, QLineEdit) and self.time_editing_item:
                try:
                    w = self.video_tree.itemWidget(self.time_editing_item, 4)
                    if w and hasattr(w, 'time_edit') and w.time_edit == focused:
                        is_time_edit_focused = True
                except Exception:
                    self.time_editing_item = None

            if is_time_edit_focused and key in (Qt.Key_Left, Qt.Key_Right):
                delta = -1000 if key == Qt.Key_Left else 1000
                self.nudge_editing_time(delta)
                return True

            if isinstance(focused, (QLineEdit, QTextEdit)) and not is_time_edit_focused:
                if key in (Qt.Key_Escape, Qt.Key_Return, Qt.Key_Enter):
                    if self.time_editing_item:
                        self.commit_time_edit(self.time_editing_item)
                        self.set_time_edit_mode(None)
                return super().eventFilter(obj, event)

            if key in (Qt.Key_Escape, Qt.Key_Return, Qt.Key_Enter):
                if self.time_editing_item:
                    self.commit_time_edit(self.time_editing_item)
                    self.set_time_edit_mode(None)
                    return True

            if key == Qt.Key_Left and (event.modifiers() & Qt.ControlModifier):
                if self.view_mode == "category":
                    self.skip_to_category_segment_start()
                else:
                    self.skip_to_segment_start()
                return True

            if key == Qt.Key_Right and (event.modifiers() & Qt.ControlModifier):
                if self.view_mode == "category":
                    self.skip_to_next_category_segment_start()
                else:
                    self.skip_to_next_segment_start()
                return True

            if key in (Qt.Key_Left, Qt.Key_Right):
                self.step_time(-1000 if key == Qt.Key_Left else 1000)
                return True

            elif key == Qt.Key_Space:
                self.toggle_play()
                return True
            elif focused is self.video_view and key in (Qt.Key_1, Qt.Key_2, Qt.Key_4, Qt.Key_8):
                self.set_playback_speed(float(event.text()))
                return True
            elif key == Qt.Key_T and self.view_mode == "video":
                self.add_timestamp_row()
                return True

        return super().eventFilter(obj, event)

    def create_menu_bar(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("File")
        new_action = QAction("New Project", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_project)
        file_menu.addAction(new_action)
        open_action = QAction("Open Project...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_project_dialog)
        file_menu.addAction(open_action)
        import_action = QAction("Import Project (.tsproject)...", self)
        import_action.setShortcut("Ctrl+I")
        import_action.triggered.connect(self.import_project_dialog)
        file_menu.addAction(import_action)
        import_excel_action = QAction("Import from Excel (.xlsx)...", self)
        import_excel_action.triggered.connect(self.import_from_excel)
        file_menu.addAction(import_excel_action)
        open_wi_action = QAction("Open WI (PDF)...", self)
        open_wi_action.triggered.connect(self.open_work_instruction)
        file_menu.addAction(open_wi_action)
        file_menu.addSeparator()
        save_action = QAction("Save Project", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_project)
        file_menu.addAction(save_action)
        save_as_action = QAction("Save Project As...", self)
        save_as_action.triggered.connect(self.save_project_as)
        file_menu.addAction(save_as_action)
        file_menu.addSeparator()
        export_csv_action = QAction("Export to CSV (.csv)...", self)
        export_csv_action.triggered.connect(self.export_to_csv)
        file_menu.addAction(export_csv_action)
        edit_menu = menu_bar.addMenu("Edit")
        self.undo_action = QAction("↩ Undo", self)
        self.undo_action.setShortcut("Ctrl+Z")
        self.undo_action.setEnabled(False)
        self.undo_action.triggered.connect(self.undo)
        edit_menu.addAction(self.undo_action)
        self.redo_action = QAction("↪ Redo", self)
        self.redo_action.setShortcut("Ctrl+Y")
        self.redo_action.setEnabled(False)
        self.redo_action.triggered.connect(self.redo)
        edit_menu.addAction(self.redo_action)

        self.settings_button = QToolButton(self)
        self.settings_button.setText("⚙")
        self.settings_button.setToolTip("Settings")
        self.settings_button.setFixedSize(32, 28)
        self.settings_button.setStyleSheet("""
            QToolButton { border: none; border-radius: 4px; color: #334155; font-size: 20px; }
            QToolButton:hover { background-color: #E2E8F0; color: #0F172A; }
        """)
        self.settings_button.clicked.connect(self.open_settings)
        menu_bar.setCornerWidget(self.settings_button, Qt.TopRightCorner)

        toolbar = QToolBar("Main Toolbar", self)
        toolbar.setMovable(False)
        toolbar.setStyleSheet("""
            QToolBar { border: none; background: transparent; spacing: 6px; padding: 2px 6px; }
            QToolButton { font-size: 13px; font-weight: bold; padding: 4px 8px; border-radius: 4px; background-color: #E2E8F0; color: #1E293B; }
            QToolButton:hover { background-color: #CBD5E1; }
            QToolButton:disabled { color: #94A3B8; background-color: #F1F5F9; }
        """)
        self.addToolBar(Qt.TopToolBarArea, toolbar)
        toolbar.addAction(self.undo_action)
        toolbar.addAction(self.redo_action)

    def open_settings(self):
        SettingsDialog(self).exec_()

    def set_auto_jump_to_slide(self, enabled):
        self.auto_jump_to_slide = enabled
        self.settings.setValue("wi/auto_jump_to_slide", enabled)
        self.settings.sync()

    def set_auto_open_wi_on_startup(self, enabled):
        self.auto_open_wi_on_startup = enabled
        self.settings.setValue("wi/auto_open_on_startup", enabled)
        self.settings.sync()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        self.video_view = ZoomableVideoView()
        self.video_view.setStyleSheet("background-color: #1e1e1e; border: 1px solid #333;")
        left_layout.addWidget(self.video_view, stretch=5)

        self.speed_btn = QPushButton("1x", self.video_view)
        self.speed_btn.setFocusPolicy(Qt.NoFocus)
        self.speed_btn.setToolTip("Toggle playback speed (1x, 2x, 4x, 8x)")
        self.speed_btn.setFixedSize(46, 46)
        self.speed_btn.setStyleSheet("""
            QPushButton { background-color: rgba(30, 30, 30, 0.85); color: #FFCA28; font-weight: bold; font-size: 13px; border: 2px solid #FFCA28; border-radius: 23px; }
            QPushButton:hover { background-color: rgba(50, 50, 50, 0.95); border-color: #FFE082; color: #FFE082; }
        """)
        self.speed_btn.clicked.connect(self.toggle_speed)
        self.video_view.set_speed_button(self.speed_btn)

        slider_layout = QVBoxLayout()
        slider_layout.setSpacing(2)
        glob_layout = QHBoxLayout()
        glob_lbl = QLabel("Global:")
        glob_lbl.setFixedWidth(65)
        glob_lbl.setStyleSheet("font-weight: bold; color: #555; font-size: 12px;")
        self.slider = JumpSlider(Qt.Horizontal)
        self.slider.setFocusPolicy(Qt.NoFocus)
        self.slider.setToolTip("Global Timeline: Shows entire video progress")
        self.slider.setRange(0, 0)
        self.slider.sliderMoved.connect(self.seek_position)
        glob_layout.addWidget(glob_lbl)
        glob_layout.addWidget(self.slider)
        slider_layout.addLayout(glob_layout)
        left_layout.addLayout(slider_layout)

        controls_layout = QHBoxLayout()
        self.open_vid_btn = QPushButton("+ Load Video")
        self.open_vid_btn.setFocusPolicy(Qt.NoFocus)
        self.open_vid_btn.setStyleSheet("font-weight: bold;")
        self.open_vid_btn.clicked.connect(self.open_file_dialog)
        controls_layout.addWidget(self.open_vid_btn)

        self.play_btn = QPushButton("Play / Pause (Space)")
        self.play_btn.setFocusPolicy(Qt.NoFocus)
        self.play_btn.clicked.connect(self.toggle_play)
        controls_layout.addWidget(self.play_btn)

        self.reset_zoom_btn = QPushButton("Reset Zoom")
        self.reset_zoom_btn.setFocusPolicy(Qt.NoFocus)
        self.reset_zoom_btn.clicked.connect(self.video_view.reset_zoom)
        controls_layout.addWidget(self.reset_zoom_btn)

        self.mark_btn = QPushButton("Add Timestamp (T)")
        self.mark_btn.setFocusPolicy(Qt.NoFocus)
        self.mark_btn.setStyleSheet("background-color: #2b7cff; color: white; font-weight: bold;")
        self.mark_btn.clicked.connect(self.add_timestamp_row)
        controls_layout.addWidget(self.mark_btn)

        left_layout.addLayout(controls_layout)
        self.time_label = QLabel("00:00 / 00:00 (◄ / ► Arrow Keys = ±1s)")
        self.time_label.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(self.time_label)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        table_controls = QHBoxLayout()

        tabs_container = QWidget()
        tabs_layout = QHBoxLayout(tabs_container)
        tabs_layout.setContentsMargins(0, 0, 0, 0)
        tabs_layout.setSpacing(-1)  # overlap borders so tabs look connected like Chrome tabs

        TAB_BTN_STYLE = """
            QPushButton {
                background-color: #E5E7EB;
                border: 1px solid #CBD5E1;
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 6px 16px;
                font-weight: bold;
                font-size: 14px;
                color: #6B7280;
            }
            QPushButton:checked {
                background-color: #FFFFFF;
                color: #111827;
                border: 1px solid #2b7cff;
                border-bottom: 2px solid #FFFFFF;
            }
            QPushButton:hover:!checked {
                background-color: #F1F5F9;
            }
        """

        self.view_mode_group = QButtonGroup(self)
        self.view_mode_group.setExclusive(True)
        self.view_mode_buttons = []
        for i, label in enumerate(["🕒 Chronological Mode", "📁 Category Mode", "📊 Pareto Mode"]):
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setFocusPolicy(Qt.NoFocus)
            btn.setStyleSheet(TAB_BTN_STYLE)
            self.view_mode_group.addButton(btn, i)
            self.view_mode_buttons.append(btn)
            tabs_layout.addWidget(btn)
        self.view_mode_buttons[0].setChecked(True)
        self.view_mode_group.idClicked.connect(self.toggle_view_mode)
        table_controls.addWidget(tabs_container)

        self.proj_status_label = QLabel("Project: Unsaved")
        self.proj_status_label.setStyleSheet("color: #888; font-style: italic; margin-left: 10px;")
        table_controls.addWidget(self.proj_status_label)
        table_controls.addStretch()

        self.export_vid_btn = QPushButton("Export Category Videos")
        self.export_vid_btn.setFocusPolicy(Qt.NoFocus)
        self.export_vid_btn.setStyleSheet("background-color: #8B5CF6; color: white; font-weight: bold;")
        self.export_vid_btn.clicked.connect(self.export_category_videos)
        self.export_vid_btn.setVisible(False) # Hide initially
        table_controls.addWidget(self.export_vid_btn)

        self.import_excel_btn = QPushButton("Import from Excel (.xlsx)")
        self.import_excel_btn.setFocusPolicy(Qt.NoFocus)
        self.import_excel_btn.setStyleSheet("background-color: #0284c7; color: white; font-weight: bold;")
        self.import_excel_btn.clicked.connect(self.import_from_excel)
        table_controls.addWidget(self.import_excel_btn)
        
        self.export_btn = QPushButton("Export to CSV (.csv)")
        self.export_btn.setFocusPolicy(Qt.NoFocus)
        self.export_btn.setStyleSheet("background-color: #28a745; color: white; font-weight: bold;")
        self.export_btn.clicked.connect(self.export_to_csv)
        table_controls.addWidget(self.export_btn)
        right_layout.addLayout(table_controls)

        tree_controls = QHBoxLayout()
        self.collapse_btn = QPushButton("Collapse All")
        self.collapse_btn.setFocusPolicy(Qt.NoFocus)
        self.collapse_btn.clicked.connect(lambda: self.video_tree.collapseAll() if self.view_mode == "video" else self.category_tree.collapseAll())
        tree_controls.addWidget(self.collapse_btn)
        
        self.expand_btn = QPushButton("Expand All")
        self.expand_btn.setFocusPolicy(Qt.NoFocus)
        self.expand_btn.clicked.connect(lambda: self.video_tree.expandAll() if self.view_mode == "video" else self.category_tree.expandAll())
        tree_controls.addWidget(self.expand_btn)
        tree_controls.addStretch()
        right_layout.addLayout(tree_controls)

        self.stacked_widget = QStackedWidget()

        # ---------- Video Mode Tree ----------
        self.video_tree = MultiVideoTreeWidget(self)
        self.video_tree.setHeaderLabels(["Slide", "Category, General", "Category, Specific", "Description", "Time"])
        self.video_tree.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.video_tree.setWordWrap(False)
        self.video_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.video_tree.customContextMenuRequested.connect(self.show_tree_context_menu)

        v_header = self.video_tree.header()
        v_header.setStretchLastSection(False)
        v_header.setSectionResizeMode(0, QHeaderView.Interactive)
        v_header.setSectionResizeMode(1, QHeaderView.Interactive)
        v_header.setSectionResizeMode(2, QHeaderView.Interactive)
        v_header.setSectionResizeMode(3, QHeaderView.Stretch)
        v_header.setSectionResizeMode(4, QHeaderView.Interactive)
        self.video_tree.setColumnWidth(0, 110)
        self.video_tree.setColumnWidth(1, 120)
        self.video_tree.setColumnWidth(2, 220)
        self.video_tree.setColumnWidth(4, 160)
        v_header.sectionResized.connect(self.on_column_resized)

        tree_style = """
            QTreeWidget { background-color: #F4F5F7; border: 1px solid #CCCCCC; font-size: 15px; }
            QHeaderView::section { background-color: #E2E8F0; color: #1E293B; font-weight: bold; font-size: 14px; padding: 8px; border-right: 1px solid #CBD5E1; border-bottom: 2px solid #CBD5E1; }
            QTreeWidget::item { padding: 4px 6px; font-size: 15px; border-right: 1px solid #E2E8F0; border-bottom: 1px solid #E2E8F0; }
            QTreeWidget::item:selected { background-color: #CBD5E1; color: #0F172A; }
        """
        self.video_tree.setStyleSheet(tree_style)
        self.video_tree.itemDoubleClicked.connect(self.on_tree_item_double_clicked)
        self.video_tree.itemClicked.connect(self.on_tree_item_clicked)
        self.video_tree.itemChanged.connect(self.on_tree_item_changed)
        self.stacked_widget.addWidget(self.video_tree)

        # ---------- Category Mode Tree ----------
        self.category_tree = QTreeWidget()
        self.category_tree.setHeaderLabels(["Slide", "Video File", "Description", "Duration"])
        self.category_tree.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.category_tree.setWordWrap(False)
        self.category_tree.setSelectionMode(QAbstractItemView.SingleSelection)
        self.category_tree.setStyleSheet(tree_style)
        
        c_header = self.category_tree.header()
        c_header.setStretchLastSection(False)
        c_header.setSectionResizeMode(0, QHeaderView.Interactive)
        c_header.setSectionResizeMode(1, QHeaderView.Interactive)
        c_header.setSectionResizeMode(2, QHeaderView.Stretch)
        c_header.setSectionResizeMode(3, QHeaderView.Interactive)
        self.category_tree.setColumnWidth(0, 200)
        self.category_tree.setColumnWidth(1, 220)
        self.category_tree.setColumnWidth(3, 160)
        self.category_tree.itemClicked.connect(lambda item, col: self.jump_wi_to_slide(item.text(0)))
        self.stacked_widget.addWidget(self.category_tree)

        # ---------- Pareto Mode Page ----------
        self.pareto_page = QWidget()
        pareto_layout = QHBoxLayout(self.pareto_page)
        pareto_layout.setContentsMargins(0, 0, 0, 0)

        self.pareto_chart = ParetoChartWidget()
        self.pareto_chart.setStyleSheet("border: 1px solid #CCCCCC;")
        pareto_layout.addWidget(self.pareto_chart, 4)

        exclude_box = QGroupBox("Categories in Pareto")
        exclude_box.setStyleSheet("QGroupBox { font-weight: bold; font-size: 13px; margin-top: 6px; }")
        exclude_layout = QVBoxLayout(exclude_box)

        hint = QLabel("Uncheck a category to remove it from the Pareto.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #64748B; font-weight: normal; font-size: 12px;")
        exclude_layout.addWidget(hint)

        self.pareto_exclude_list = QListWidget()
        self.pareto_exclude_list.setStyleSheet(
            "QListWidget { background-color: #F4F5F7; border: 1px solid #CBD5E1; font-size: 13px; }"
            "QListWidget::item { padding: 3px 4px; }"
        )
        self.pareto_exclude_list.itemChanged.connect(self.on_pareto_exclusion_changed)
        exclude_layout.addWidget(self.pareto_exclude_list)

        exclude_btns = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.setFocusPolicy(Qt.NoFocus)
        select_all_btn.clicked.connect(lambda: self.set_all_pareto_categories(True))
        exclude_btns.addWidget(select_all_btn)
        clear_all_btn = QPushButton("Clear All")
        clear_all_btn.setFocusPolicy(Qt.NoFocus)
        clear_all_btn.clicked.connect(lambda: self.set_all_pareto_categories(False))
        exclude_btns.addWidget(clear_all_btn)
        exclude_layout.addLayout(exclude_btns)

        exclude_box.setMaximumWidth(280)
        pareto_layout.addWidget(exclude_box, 1)
        self.stacked_widget.addWidget(self.pareto_page)

        right_layout.addWidget(self.stacked_widget)
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 5)
        main_layout.addWidget(splitter)

    def _capture_mode_selection(self):
        """Remember which row was selected in the tree we're leaving."""
        if not hasattr(self, 'video_tree'):
            return
        if self.view_mode == "video":
            item = self.video_tree.currentItem()
            if item and item.parent():
                parent = item.parent()
                self.last_video_selection = (
                    self.video_tree.indexOfTopLevelItem(parent),
                    parent.indexOfChild(item)
                )
        elif self.view_mode == "category":
            item = self.category_tree.currentItem()
            if item is not None and getattr(item, 'g_idx', -1) >= 0:
                self.last_category_selection = (item.g_idx, item.r_idx)

    def _restore_video_selection(self):
        if not self.last_video_selection or self.video_tree.topLevelItemCount() == 0:
            return
        g_idx, r_idx = self.last_video_selection
        grp = self.video_tree.topLevelItem(min(g_idx, self.video_tree.topLevelItemCount() - 1))
        segment_count = max(0, grp.childCount() - 1)  # last child is the END VIDEO sentinel
        if segment_count == 0:
            target = grp
        else:
            grp.setExpanded(True)
            target = grp.child(max(0, min(r_idx, segment_count - 1)))
        self.video_tree.setCurrentItem(target)
        self.video_tree.scrollToItem(target)

    def _restore_category_selection(self):
        if not self.last_category_selection:
            return
        g_idx, r_idx = self.last_category_selection
        best = None
        best_distance = None
        for i in range(self.category_tree.topLevelItemCount()):
            cat_grp = self.category_tree.topLevelItem(i)
            for c in range(cat_grp.childCount()):
                child = cat_grp.child(c)
                if getattr(child, 'g_idx', -1) != g_idx:
                    continue
                distance = abs(getattr(child, 'r_idx', 0) - r_idx)
                if best_distance is None or distance < best_distance:
                    best, best_distance = child, distance
                    if distance == 0:
                        break
            if best_distance == 0:
                break
        if best is None:
            return
        self.category_tree.setCurrentItem(best)
        self.category_tree.scrollToItem(best)

    def toggle_view_mode(self, index):
        target_mode = ("video", "category", "pareto")[index]
        if self.view_mode == target_mode:
            self.view_mode_buttons[index].setChecked(True)
            return
        self._capture_mode_selection()
        self.view_mode_buttons[index].setChecked(True)
        if index == 2:
            if self.view_mode == "video":
                self.saved_video_state = self.get_current_state()
            self.view_mode = "pareto"
            self.timer.stop()
            self.is_playing = False
            self.active_category_item = None

            self.mark_btn.setEnabled(False)
            self.mark_btn.setStyleSheet("background-color: #94A3B8; color: white; font-weight: bold;")
            self.open_vid_btn.setEnabled(False)
            self.open_vid_btn.setStyleSheet("background-color: #E2E8F0; color: #94A3B8; font-weight: bold;")

            self.export_vid_btn.setVisible(False)
            self.import_excel_btn.setVisible(False)
            self.export_btn.setVisible(False)
            self.collapse_btn.setVisible(False)
            self.expand_btn.setVisible(False)

            self.build_pareto_view()
            self.stacked_widget.setCurrentIndex(2)
            self.update_undo_redo_actions()
        elif index == 1:
            if self.view_mode == "video":
                self.saved_video_state = self.get_current_state()
            self.view_mode = "category"
            self.timer.stop()
            self.is_playing = False
            self.active_category_item = None
            
            self.mark_btn.setEnabled(False)
            self.mark_btn.setStyleSheet("background-color: #94A3B8; color: white; font-weight: bold;")
            
            # Gray out Load Video button
            self.open_vid_btn.setEnabled(False)
            self.open_vid_btn.setStyleSheet("background-color: #E2E8F0; color: #94A3B8; font-weight: bold;")
            
            # Toggle export/import buttons
            self.export_vid_btn.setVisible(True)
            self.import_excel_btn.setVisible(False)
            self.export_btn.setVisible(False)
            self.collapse_btn.setVisible(True)
            self.expand_btn.setVisible(True)
            
            self.build_category_view()
            self.stacked_widget.setCurrentIndex(1)
            self._restore_category_selection()
        else:
            self.view_mode = "video"
            self.timer.stop()
            self.is_playing = False
            self.active_category_item = None
            
            self.mark_btn.setEnabled(True)
            self.mark_btn.setStyleSheet("background-color: #2b7cff; color: white; font-weight: bold;")
            
            # Restore Load Video button
            self.open_vid_btn.setEnabled(True)
            self.open_vid_btn.setStyleSheet("font-weight: bold;")
            
            # Toggle export/import buttons
            self.export_vid_btn.setVisible(False)
            self.import_excel_btn.setVisible(True)
            self.export_btn.setVisible(True)
            self.collapse_btn.setVisible(True)
            self.expand_btn.setVisible(True)
            
            self.stacked_widget.setCurrentIndex(0)
            self._restore_video_selection()
            self.update_undo_redo_actions()

    def get_pareto_category_totals(self):
        """Total delta (ms) per specific category across the whole study, descending."""
        totals = {}
        for segment in self.get_category_segments(self.saved_video_state):
            totals[segment["cat_spec"]] = totals.get(segment["cat_spec"], 0) + segment["delta_ms"]
        return sorted(totals.items(), key=lambda x: x[1], reverse=True)

    def build_pareto_view(self):
        totals = self.get_pareto_category_totals()

        self.pareto_exclude_list.blockSignals(True)
        self.pareto_exclude_list.clear()
        for cat_name, total_ms in totals:
            item = QListWidgetItem(f"{cat_name}  ({self.format_ms(total_ms)})")
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setData(Qt.UserRole, cat_name)
            item.setCheckState(Qt.Unchecked if cat_name in self.pareto_excluded_cats else Qt.Checked)
            self.pareto_exclude_list.addItem(item)
        self.pareto_exclude_list.blockSignals(False)

        self.refresh_pareto_chart()

    def refresh_pareto_chart(self):
        included = [(name, ms) for name, ms in self.get_pareto_category_totals()
                    if name not in self.pareto_excluded_cats and ms > 0]
        grand_total = sum(ms for _, ms in included)

        chart_data = []
        running = 0
        for name, ms in included:
            running += ms
            chart_data.append((name, ms / 1000.0, running / grand_total if grand_total else 0))
        self.pareto_chart.set_data(chart_data)

    def on_pareto_exclusion_changed(self, item):
        cat_name = item.data(Qt.UserRole)
        if item.checkState() == Qt.Checked:
            self.pareto_excluded_cats.discard(cat_name)
        else:
            self.pareto_excluded_cats.add(cat_name)
        self.refresh_pareto_chart()

    def set_all_pareto_categories(self, checked):
        self.pareto_exclude_list.blockSignals(True)
        for i in range(self.pareto_exclude_list.count()):
            item = self.pareto_exclude_list.item(i)
            item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
            cat_name = item.data(Qt.UserRole)
            if checked:
                self.pareto_excluded_cats.discard(cat_name)
            else:
                self.pareto_excluded_cats.add(cat_name)
        self.pareto_exclude_list.blockSignals(False)
        self.refresh_pareto_chart()

    def build_category_view(self):
        self.category_tree.setUpdatesEnabled(False)
        self.category_tree.clear()
        cat_segments = {}

        for segment in self.get_category_segments(self.saved_video_state):
            cat_spec = segment["cat_spec"]
            if cat_spec not in cat_segments:
                cat_segments[cat_spec] = {"total_time": 0, "segments": []}

            cat_segments[cat_spec]["total_time"] += segment["delta_ms"]
            segments = cat_segments[cat_spec]["segments"]
            if segments and self._is_continuation_segment(segments[-1], segment):
                segments[-1]["slices"].append(segment)
                segments[-1]["delta_ms"] += segment["delta_ms"]
            else:
                segment["slices"] = [segment]
                segments.append(segment)
                
        sorted_cats = sorted(cat_segments.items(), key=lambda x: x[1]["total_time"], reverse=True)
        bold_font = QFont()
        bold_font.setBold(True)
        
        for cat_name, data in sorted_cats:
            cat_item = QTreeWidgetItem(self.category_tree)
            cat_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            cat_item.setText(0, cat_name)
            cat_item.setText(3, self.format_ms(data["total_time"])) # Moved to col 3
            
            bg, fg = CATEGORY_COLORS.get(cat_name, ("#E2E8F0", "#1E293B"))
            bg_brush = QBrush(QColor(bg))
            fg_brush = QBrush(QColor(fg))
            for col in range(4): # Loop 4 columns
                cat_item.setBackground(col, bg_brush)
                cat_item.setForeground(col, fg_brush)
                cat_item.setFont(col, bold_font)
                
            sorted_segs = sorted(data["segments"], key=lambda x: x["delta_ms"], reverse=True)
            
            for seg in sorted_segs:
                child = QTreeWidgetItem(cat_item)
                child.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                
                vid_display_name = f"Video {seg['video_g_idx'] + 1} ({seg['vid_name']})"
                
                child.setText(0, seg["slide"])
                child.setText(1, vid_display_name)
                child.setText(2, seg["desc"]) # Description is now col 2
                
                child.setData(0, Qt.UserRole, seg["video_path"])
                child.setData(1, Qt.UserRole, seg["start_ms"])
                child.setData(2, Qt.UserRole, seg["end_ms"])
                child.setData(3, Qt.UserRole, seg["delta_ms"])
                child.category_slices = seg["slices"]
                child.category_slice_index = 0
                child.g_idx = seg["g_idx"]
                child.r_idx = seg["r_idx"]
                
                time_str = self.format_ms(seg["delta_ms"])
                action_widget = self.create_inline_action_widget(
                    time_str, item=child, is_editable=False,
                    on_play=lambda _, it=child: self.play_category_segment(it),
                    on_delete=None,
                    on_jump=lambda _, it=child: self.jump_to_video(it)
                )
                self.category_tree.setItemWidget(child, 3, action_widget) # Set to col 3
                
        self.category_tree.expandAll()
        self.category_tree.setUpdatesEnabled(True)

    def _is_continuation_segment(self, previous_segment, segment):
        return (
            previous_segment["g_idx"] == segment["g_idx"]
            and previous_segment["r_idx"] == segment["r_idx"]
            and segment["start_ms"] == 0
        )

    def get_category_segments(self, groups_data):
        """Return category segments, including unmarked time at video boundaries."""
        for group_index, group in enumerate(groups_data):
            rows = group.get("rows", [])
            duration_ms = group.get("duration_ms", 0)
            video_path = group.get("video_path", "")

            for row_index, row in enumerate(rows):
                start_ms = row["time_ms"]
                end_ms = rows[row_index + 1]["time_ms"] if row_index + 1 < len(rows) else duration_ms
                yield self._make_category_segment(
                    row, video_path, group_index, row_index, group_index, start_ms, end_ms
                )

                if row_index != len(rows) - 1:
                    continue

                for next_group_index in range(group_index + 1, len(groups_data)):
                    next_group = groups_data[next_group_index]
                    next_rows = next_group.get("rows", [])
                    next_duration_ms = next_group.get("duration_ms", 0)
                    next_video_path = next_group.get("video_path", "")
                    continuation_end_ms = next_rows[0]["time_ms"] if next_rows else next_duration_ms
                    yield self._make_category_segment(
                        row, next_video_path, next_group_index, row_index, group_index,
                        0, continuation_end_ms
                    )
                    if next_rows:
                        break

    def _make_category_segment(self, row, video_path, video_group_index, row_index, source_group_index, start_ms, end_ms):
        start_ms = max(0, start_ms)
        end_ms = max(start_ms, end_ms)
        return {
            "video_path": video_path,
            "vid_name": os.path.basename(video_path) if video_path else "Unknown",
            "start_ms": start_ms,
            "end_ms": end_ms,
            "delta_ms": end_ms - start_ms,
            "slide": row.get("slide", ""),
            "desc": row.get("desc", ""),
            "cat_spec": row.get("cat_spec", "").strip() or "Uncategorized",
            "g_idx": source_group_index,
            "r_idx": row_index,
            "video_g_idx": video_group_index
        }

    def jump_to_category(self, video_item):
        """Switches to Category Mode, finds the matching item, and auto-scrolls to it."""
        if not video_item or not video_item.parent(): return
        g_idx = self.video_tree.indexOfTopLevelItem(video_item.parent())
        r_idx = video_item.parent().indexOfChild(video_item)
        
        self.toggle_view_mode(1)
        for i in range(self.category_tree.topLevelItemCount()):
            cat_grp = self.category_tree.topLevelItem(i)
            for c in range(cat_grp.childCount()):
                cat_child = cat_grp.child(c)
                if getattr(cat_child, 'g_idx', -1) == g_idx and getattr(cat_child, 'r_idx', -1) == r_idx:
                    self.category_tree.setCurrentItem(cat_child)
                    self.category_tree.scrollToItem(cat_child)
                    return

    def jump_to_video(self, category_item):
        """Switches to Video Mode, finds the original parent row, and auto-scrolls to it."""
        g_idx = getattr(category_item, 'g_idx', -1)
        r_idx = getattr(category_item, 'r_idx', -1)
        if g_idx >= 0 and r_idx >= 0:
            self.toggle_view_mode(0)
            if g_idx < self.video_tree.topLevelItemCount():
                grp = self.video_tree.topLevelItem(g_idx)
                grp.setExpanded(True)
                if r_idx < grp.childCount():
                    target = grp.child(r_idx)
                    self.video_tree.setCurrentItem(target)
                    self.video_tree.scrollToItem(target)

    def export_category_videos(self):
        """Merges all video segments from a selected category into an MP4 file."""
        base_state = self.saved_video_state if self.view_mode != "video" else self.get_current_state()
        if not base_state:
            QMessageBox.information(self, "Export", "There is no video data loaded to export.")
            return

        cat_segments = {}
        for segment in self.get_category_segments(base_state):
            video_path = segment["video_path"]
            if not video_path or not os.path.exists(video_path):
                continue

            cat_segments.setdefault(segment["cat_spec"], []).append(
                (video_path, segment["start_ms"], segment["end_ms"])
            )

        if not cat_segments:
            QMessageBox.information(self, "Export", "No categorized segments found.")
            return

        # Prompt user to choose which category to export
        cat_choices = sorted(list(cat_segments.keys()))
        selected_cat, ok = QInputDialog.getItem(self, "Select Category", "Choose which category to export:", cat_choices, 0, False)
        if not ok or not selected_cat: return
        
        # Filter down to only the chosen category
        cat_segments = {selected_cat: cat_segments[selected_cat]}

        speed_str, ok = QInputDialog.getItem(self, "Export Speed", "Select export speed multiplier:", ["1x", "2x", "4x", "8x"], 0, False)
        if not ok: return
        speed_factor = int(speed_str.replace("x", ""))

        out_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory for MP4", self._default_browse_dir())
        if not out_dir: return

        progress = QProgressDialog(f"Exporting '{selected_cat}'...", "Cancel", 0, len(cat_segments[selected_cat]), self)
        progress.setWindowTitle("Exporting Video")
        progress.setWindowModality(Qt.WindowModal)
        progress.show()

        for i, (cat, segments) in enumerate(cat_segments.items()):
            safe_cat_name = "".join([c if c.isalnum() or c in " _-" else "_" for c in cat])
            out_path = os.path.join(out_dir, f"{safe_cat_name}.mp4")
            
            writer = None
            target_size = None
            target_fps = 30.0

            for j, (vid_path, start_ms, end_ms) in enumerate(segments):
                if progress.wasCanceled(): break
                progress.setLabelText(f"Exporting '{cat}'... (Clip {j+1}/{len(segments)})")
                progress.setValue(j)
                QApplication.processEvents()
                
                cap = cv2.VideoCapture(vid_path)
                if not cap.isOpened(): continue

                if writer is None:
                    target_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
                    w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                    h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                    target_size = (w, h)
                    fourcc = cv2.VideoWriter_fourcc(*'mp4v')
                    writer = cv2.VideoWriter(out_path, fourcc, target_fps, target_size)

                cap.set(cv2.CAP_PROP_POS_MSEC, start_ms)
                frame_count = 0

                while True:
                    ret, frame = cap.read()
                    if not ret: break

                    curr_ms = cap.get(cv2.CAP_PROP_POS_MSEC)
                    if curr_ms > end_ms and end_ms > 0:
                        break

                    if frame_count % speed_factor == 0:
                        if (frame.shape[1], frame.shape[0]) != target_size:
                            frame = cv2.resize(frame, target_size)
                        writer.write(frame)

                    frame_count += 1
                    if frame_count % 30 == 0:
                        QApplication.processEvents()

                cap.release()

            if writer:
                writer.release()

        progress.setValue(len(cat_segments[selected_cat]))
        if not progress.wasCanceled():
            QMessageBox.information(self, "Export Complete", f"Successfully exported video to:\n{out_path}")

    def play_category_segment(self, item):
        self._play_category_slice(item, 0)

    def _play_category_slice(self, item, slice_index):
        slices = getattr(item, "category_slices", [])
        if slices:
            segment = slices[slice_index]
            video_path = segment["video_path"]
            start_ms = segment["start_ms"]
            end_ms = segment["end_ms"]
        else:
            video_path = item.data(0, Qt.UserRole)
            start_ms = item.data(1, Qt.UserRole)
            end_ms = item.data(2, Qt.UserRole)
        
        if not video_path or not os.path.exists(video_path):
            QMessageBox.warning(self, "Video Missing", f"Could not find video file:\n{video_path}")
            return
            
        self.active_category_item = item
        item.category_slice_index = slice_index
        item.setData(0, Qt.UserRole, video_path)
        item.setData(1, Qt.UserRole, start_ms)
        item.setData(2, Qt.UserRole, end_ms)
        
        if self.active_video_path != video_path or not self.cap or not self.cap.isOpened():
            if self.cap:
                self.cap.release()
            self.active_video_path = video_path
            self.cap = cv2.VideoCapture(video_path)
            self.fps = self.cap.get(cv2.CAP_PROP_FPS) or 30.0
            self.total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
            self.duration_ms = int((self.total_frames / self.fps) * 1000)
            self.slider.setRange(0, self.duration_ms)
            interval = max(1, int(1000 / (self.fps)))
            self.timer.setInterval(interval)
            self.video_view.reset_zoom()
            
        self.seek_position(start_ms)
        if not self.is_playing:
            self.toggle_play()

    def get_next_category_segment(self, current_item):
        parent = current_item.parent()
        if not parent: return None
        idx = parent.indexOfChild(current_item)
        if idx + 1 < parent.childCount():
            return parent.child(idx + 1)
        else:
            p_idx = self.category_tree.indexOfTopLevelItem(parent)
            if p_idx + 1 < self.category_tree.topLevelItemCount():
                return self.category_tree.topLevelItem(p_idx + 1).child(0)
        return None

    def get_previous_category_segment(self, current_item):
        parent = current_item.parent()
        if not parent:
            return None
        idx = parent.indexOfChild(current_item)
        if idx > 0:
            return parent.child(idx - 1)

        p_idx = self.category_tree.indexOfTopLevelItem(parent)
        if p_idx > 0:
            previous_parent = self.category_tree.topLevelItem(p_idx - 1)
            if previous_parent.childCount() > 0:
                return previous_parent.child(previous_parent.childCount() - 1)
        return None

    def create_category_combo(self, item, cat_type, current_text):
        combo = NoScrollComboBox()
        options = self.cat_general_options if cat_type == "general" else self.cat_specific_options
        combo.addItems([""] + options + ["+ Add New Category..."])
        combo.setEditable(True)
        combo.setCurrentText(current_text)

        if combo.lineEdit():
            combo.lineEdit().installEventFilter(self.select_all_filter)

        delegate = CategoryDelegate(self, cat_type, combo)
        combo._custom_delegate = delegate
        combo.setItemDelegate(delegate)
        
        combo_filter = ComboViewFilter(self, combo, cat_type)
        combo._custom_filter = combo_filter
        combo.view().viewport().installEventFilter(combo_filter)
        self.apply_chip_style(combo)

        def on_change(text):
            if text == "+ Add New Category...":
                self.handle_add_new_category(combo, cat_type, item)
            else:
                role_col = 1 if cat_type == "general" else 2
                item.setData(role_col, Qt.UserRole, text)
                self.apply_chip_style(combo)
                self.push_state()

        combo.currentTextChanged.connect(on_change)
        return combo

    def handle_add_new_category(self, combo, cat_type, item):
        role_col = 1 if cat_type == "general" else 2
        prev_val = item.data(role_col, Qt.UserRole) or ""
        cat_title = "General" if cat_type == "general" else "Specific"
        new_cat, ok = QInputDialog.getText(self, f"Add New {cat_title} Category", f"Enter name for new {cat_title} category:")

        if ok and new_cat.strip():
            clean_cat = new_cat.strip()
            options = self.cat_general_options if cat_type == "general" else self.cat_specific_options

            if clean_cat not in options:
                options.append(clean_cat)

            if cat_type == "general":
                self.custom_gen_cats.add(clean_cat)
            else:
                self.custom_spec_cats.add(clean_cat)

            self._ensure_category_color(clean_cat)

            item.setData(role_col, Qt.UserRole, clean_cat)
            self.refresh_all_combos()
            self.push_state()
        else:
            combo.setCurrentText(prev_val)
            self.apply_chip_style(combo)

    def rename_category(self, old_name, cat_type):
        cat_title = "General" if cat_type == "general" else "Specific"
        new_cat, ok = QInputDialog.getText(self, f"Rename {cat_title} Category", f"Enter new name for '{old_name}':", QLineEdit.Normal, old_name)

        if ok and new_cat.strip() and new_cat.strip() != old_name:
            clean_cat = new_cat.strip()
            options = self.cat_general_options if cat_type == "general" else self.cat_specific_options

            if clean_cat not in options:
                options.append(clean_cat)
            if old_name in CATEGORY_COLORS and clean_cat not in CATEGORY_COLORS:
                CATEGORY_COLORS[clean_cat] = CATEGORY_COLORS[old_name]
                if old_name in self.custom_category_colors:
                    self.custom_category_colors[clean_cat] = self.custom_category_colors.pop(old_name)
            if old_name in options:
                options.remove(old_name)

            if cat_type == "general":
                if old_name in self.custom_gen_cats:
                    self.custom_gen_cats.remove(old_name)
                    self.custom_gen_cats.add(clean_cat)
            else:
                if old_name in self.custom_spec_cats:
                    self.custom_spec_cats.remove(old_name)
                    self.custom_spec_cats.add(clean_cat)

            self.video_tree.blockSignals(True)
            for i in range(self.video_tree.topLevelItemCount()):
                group = self.video_tree.topLevelItem(i)
                for c in range(group.childCount()):
                    child = group.child(c)
                    if "END VIDEO" in child.text(3): continue
                    col = 1 if cat_type == "general" else 2
                    if child.data(col, Qt.UserRole) == old_name:
                        child.setData(col, Qt.UserRole, clean_cat)
            self.video_tree.blockSignals(False)
            
            self.refresh_all_combos()
            self.push_state()

    def delete_category(self, cat_name, cat_type):
        reply = QMessageBox.question(self, "Delete Category", f"Are you sure you want to completely delete the category '{cat_name}'? It will be removed from all tasks.", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.push_state()
            options = self.cat_general_options if cat_type == "general" else self.cat_specific_options
            if cat_name in options:
                options.remove(cat_name)
            
            if cat_type == "general" and cat_name in self.custom_gen_cats:
                self.custom_gen_cats.remove(cat_name)
            elif cat_type == "specific" and cat_name in self.custom_spec_cats:
                self.custom_spec_cats.remove(cat_name)

            self.custom_category_colors.pop(cat_name, None)
            if cat_name not in self.default_cat_general_options and cat_name not in self.default_cat_specific_options:
                CATEGORY_COLORS.pop(cat_name, None)
            
            self.video_tree.blockSignals(True)
            for i in range(self.video_tree.topLevelItemCount()):
                group = self.video_tree.topLevelItem(i)
                for c in range(group.childCount()):
                    child = group.child(c)
                    if "END VIDEO" in child.text(3): continue
                    col = 1 if cat_type == "general" else 2
                    if child.data(col, Qt.UserRole) == cat_name:
                        child.setData(col, Qt.UserRole, "")
            self.video_tree.blockSignals(False)
            self.refresh_all_combos()
            self.push_state()

    def refresh_all_combos(self):
        if self.view_mode != "video": return
        gen_counts = {}
        spec_counts = {}
        for i in range(self.video_tree.topLevelItemCount()):
            group = self.video_tree.topLevelItem(i)
            for c in range(group.childCount()):
                child = group.child(c)
                if "END VIDEO" in child.text(3): continue
                g_val = child.data(1, Qt.UserRole) or ""
                s_val = child.data(2, Qt.UserRole) or ""
                if g_val: gen_counts[g_val] = gen_counts.get(g_val, 0) + 1
                if s_val: spec_counts[s_val] = spec_counts.get(s_val, 0) + 1

        self.cat_general_options.sort(key=lambda x: (-gen_counts.get(x, 0), x.lower()))
        self.cat_specific_options.sort(key=lambda x: (
            -spec_counts.get(x, 0), 
            self.default_cat_specific_options.index(x) if x in self.default_cat_specific_options else 999,
            x.lower()
        ))

        for i in range(self.video_tree.topLevelItemCount()):
            group = self.video_tree.topLevelItem(i)
            for c in range(group.childCount()):
                child = group.child(c)
                if "END VIDEO" in child.text(3): continue
                combo_gen = self.video_tree.itemWidget(child, 1)
                if combo_gen:
                    cur = child.data(1, Qt.UserRole) or ""
                    combo_gen.blockSignals(True)
                    combo_gen.clear()
                    combo_gen.addItems([""] + self.cat_general_options + ["+ Add New Category..."])
                    combo_gen.setCurrentText(cur)
                    self.apply_chip_style(combo_gen)
                    combo_gen.blockSignals(False)
                combo_spec = self.video_tree.itemWidget(child, 2)
                if combo_spec:
                    cur = child.data(2, Qt.UserRole) or ""
                    combo_spec.blockSignals(True)
                    combo_spec.clear()
                    combo_spec.addItems([""] + self.cat_specific_options + ["+ Add New Category..."])
                    combo_spec.setCurrentText(cur)
                    self.apply_chip_style(combo_spec)
                    combo_spec.blockSignals(False)

    def show_tree_context_menu(self, pos):
        if self.view_mode != "video": return
        item = self.video_tree.itemAt(pos)
        if not item: return
        if item.parent() is None:
            menu = QMenu(self)
            replace_action = menu.addAction("Replace Video")
            action = menu.exec_(self.video_tree.viewport().mapToGlobal(pos))
            if action == replace_action:
                self.replace_video_for_group(item)

    def replace_video_for_group(self, group_item):
        old_path = group_item.data(0, Qt.UserRole) or ""
        grp_name = os.path.basename(old_path) if old_path else "Unlinked Video"
        new_path, _ = QFileDialog.getOpenFileName(self, f"Select Replacement Video File for {grp_name}", self._default_browse_dir(), "Video Files (*.mp4 *.avi *.mkv *.mov)")
        if not new_path or not os.path.exists(new_path): return
        self.push_state()
        dur_ms = 0
        cap = cv2.VideoCapture(new_path)
        if cap.isOpened():
            fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            dur_ms = int((total_frames / fps) * 1000)
            cap.release()
        group_item.setData(0, Qt.UserRole, new_path)
        group_item.setData(1, Qt.UserRole, dur_ms)
        child_count = group_item.childCount()
        if child_count > 0:
            end_item = group_item.child(child_count - 1)
            if "END VIDEO" in end_item.text(3):
                end_str = self.format_ms(dur_ms)
                end_item.setText(4, end_str)
                end_item.setData(4, Qt.UserRole, dur_ms)
                end_item.setData(4, Qt.UserRole + 1, end_str)
        if self.active_group_item == group_item or self.active_video_path == old_path:
            self.active_video_path = ""
            self.switch_active_video(group_item, 0)
        self.renumber_video_groups()
        self.refresh_playback_ui()

    def swap_rows_content(self, item1, item2):
        self.push_state()
        self.video_tree.blockSignals(True)
        slide1, slide2 = item1.text(0), item2.text(0)
        item1.setText(0, slide2)
        item2.setText(0, slide1)
        cat_gen1, cat_gen2 = item1.data(1, Qt.UserRole) or "", item2.data(1, Qt.UserRole) or ""
        item1.setData(1, Qt.UserRole, cat_gen2)
        item2.setData(1, Qt.UserRole, cat_gen1)
        combo_gen1 = self.video_tree.itemWidget(item1, 1)
        if combo_gen1:
            combo_gen1.blockSignals(True)
            combo_gen1.setCurrentText(cat_gen2)
            self.apply_chip_style(combo_gen1)
            combo_gen1.blockSignals(False)
        combo_gen2 = self.video_tree.itemWidget(item2, 1)
        if combo_gen2:
            combo_gen2.blockSignals(True)
            combo_gen2.setCurrentText(cat_gen1)
            self.apply_chip_style(combo_gen2)
            combo_gen2.blockSignals(False)
        cat_spec1, cat_spec2 = item1.data(2, Qt.UserRole) or "", item2.data(2, Qt.UserRole) or ""
        item1.setData(2, Qt.UserRole, cat_spec2)
        item2.setData(2, Qt.UserRole, cat_spec1)
        combo_spec1 = self.video_tree.itemWidget(item1, 2)
        if combo_spec1:
            combo_spec1.blockSignals(True)
            combo_spec1.setCurrentText(cat_spec2)
            self.apply_chip_style(combo_spec1)
            combo_spec1.blockSignals(False)
        combo_spec2 = self.video_tree.itemWidget(item2, 2)
        if combo_spec2:
            combo_spec2.blockSignals(True)
            combo_spec2.setCurrentText(cat_spec1)
            self.apply_chip_style(combo_spec2)
            combo_spec2.blockSignals(False)
        desc1, desc2 = item1.text(3), item2.text(3)
        item1.setText(3, desc2)
        item2.setText(3, desc1)
        self.video_tree.blockSignals(False)
        self.video_tree.doItemsLayout()
        self.refresh_all_combos()
        self.refresh_playback_ui()

    def import_from_excel(self):
        excel_path, _ = QFileDialog.getOpenFileName(self, "Import Excel Spreadsheet", self._default_browse_dir(), "Excel Workbook (*.xlsx *.xls)")
        if not excel_path: return
        progress = QProgressDialog(self)
        progress.setWindowTitle("Loading")
        progress.setLabelText("Opening Excel file (this may take a moment)...")
        progress.setRange(0, 0)
        progress.setCancelButton(None)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()
        QApplication.processEvents()
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            sheet_names = wb.sheetnames
            progress.close()
            if not sheet_names:
                QMessageBox.critical(self, "Excel Load Error", "No worksheets found in workbook.")
                return
            if len(sheet_names) == 1:
                selected_sheet = sheet_names[0]
            else:
                selected_sheet, ok = QInputDialog.getItem(self, "Select Worksheet", "Choose the worksheet to import data from:", sheet_names, 0, False)
                if not ok or not selected_sheet: return
            ws = wb[selected_sheet]
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Excel Load Error", f"Failed to open Excel file:\n{str(e)}")
            return
        range_str, ok = QInputDialog.getText(self, "Define Import Cell Range", f"Enter cell range from sheet '{selected_sheet}'\n(Col 1: Slide, 2: Gen Cat, 3: Spec Cat, 4: Desc, 5: Time):\nExample: A4:E250 or A2:E100", QLineEdit.Normal, "A2:E100")
        if not ok or not range_str.strip(): return
        progress = QProgressDialog(self)
        progress.setWindowTitle("Importing")
        progress.setLabelText("Parsing rows...")
        progress.setRange(0, 0)
        progress.setCancelButton(None)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()
        QApplication.processEvents()
        try:
            cells_grid = ws[range_str.strip().upper()]
            if isinstance(cells_grid, openpyxl.cell.cell.Cell): cells_grid = ((cells_grid,),)
            elif len(cells_grid) > 0 and isinstance(cells_grid[0], openpyxl.cell.cell.Cell): cells_grid = (cells_grid,)
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Invalid Range", f"Invalid cell range '{range_str}':\n{str(e)}")
            return
        parsed_groups = []
        current_rows = []
        current_group_name = "Imported Video 1"
        has_active_group = False
        for row_cells in cells_grid:
            vals = [c.value if c.value is not None else "" for c in row_cells]
            if len(vals) < 5: vals += [""] * (5 - len(vals))
            slide = clean_slide_str(vals[0])
            cat_gen = str(vals[1]).strip()
            cat_spec = str(vals[2]).strip()
            desc = str(vals[3]).strip()
            time_val = str(vals[4]).strip()
            if not any([slide, cat_gen, cat_spec, desc, time_val]): continue
            is_start = "START VIDEO" in desc.upper() or "START VIDEO" in slide.upper()
            is_end = "END VIDEO" in desc.upper() or "END VIDEO" in slide.upper()
            if is_start:
                if current_rows or has_active_group:
                    parsed_groups.append({"name": current_group_name, "rows": current_rows})
                    current_rows = []
                current_group_name = desc if desc else (slide if slide else f"Video {len(parsed_groups) + 1}")
                has_active_group = True
                continue
            if is_end:
                if current_rows or has_active_group:
                    parsed_groups.append({"name": current_group_name, "rows": current_rows})
                    current_rows = []
                    has_active_group = False
                continue
            time_ms = self.parse_time_ms(None, time_val)
            formatted_time = self.format_ms(time_ms) if time_ms > 0 else time_val
            current_rows.append({
                "slide": slide, "cat_gen": cat_gen, "cat_spec": cat_spec, 
                "desc": desc, "time_str": formatted_time, "time_ms": time_ms
            })
        if current_rows or has_active_group:
            parsed_groups.append({"name": current_group_name, "rows": current_rows})
        progress.close()
        if not parsed_groups:
            QMessageBox.information(self, "No Data Found", f"No valid rows found in range {range_str} of sheet '{selected_sheet}'.")
            return
        self._scan_and_add_custom_categories(parsed_groups)
        if self.video_tree.topLevelItemCount() > 0:
            msg = QMessageBox(self)
            msg.setWindowTitle("Import Options")
            msg.setText(f"Found {len(parsed_groups)} video group(s) in sheet '{selected_sheet}'. Choose import mode:")
            btn_replace = msg.addButton("Replace Session", QMessageBox.AcceptRole)
            btn_merge = msg.addButton("Merge Groups", QMessageBox.ActionRole)
            btn_cancel = msg.addButton(QMessageBox.Cancel)
            msg.exec_()
            clicked = msg.clickedButton()
            if clicked == btn_cancel: return
            elif clicked == btn_replace:
                self.set_time_edit_mode(None)
                self.video_tree.clear()
                self.undo_stack.clear()
                self.redo_stack.clear()
                self.update_undo_redo_actions()
        for idx, grp in enumerate(parsed_groups):
            grp_name = grp["name"]
            grp_rows = grp["rows"]
            video_path, _ = QFileDialog.getOpenFileName(self, f"Select Video File for '{grp_name}' (Group {idx + 1}/{len(parsed_groups)})\n(Click Cancel to skip picking a video)", self._default_browse_dir(), "Video Files (*.mp4 *.avi *.mkv *.mov)")
            self.add_video_group(video_path, existing_rows=grp_rows)
        self.renumber_video_groups()
        self.refresh_all_combos()
        self.refresh_playback_ui()
        QMessageBox.information(self, "Import Complete", f"Successfully imported {len(parsed_groups)} video group(s) from sheet '{selected_sheet}'!")

    def on_tree_item_changed(self, item, column):
        if not self.is_restoring_state and self.view_mode == "video":
            self.video_tree.blockSignals(True)
            if column in (0, 3):
                txt = item.text(column)
                if "\n" in txt or "\r" in txt:
                    clean_txt = txt.replace("\r", "").replace("\n", " ")
                    item.setText(column, clean_txt)
            self.video_tree.doItemsLayout()
            self.push_state()
            self.video_tree.blockSignals(False)

    def on_column_resized(self, logical_index, old_size, new_size):
        if self.is_restoring_state: return
        try:
            viewport_w = self.video_tree.viewport().width()
            if viewport_w <= 0: return
            min_desc_w = 60
            w0 = self.video_tree.columnWidth(0)
            w1 = self.video_tree.columnWidth(1)
            w2 = self.video_tree.columnWidth(2)
            w4 = self.video_tree.columnWidth(4)
            max_time_w = max(80, viewport_w - (w0 + w1 + w2 + min_desc_w))
            header = self.video_tree.header()
            header.blockSignals(True)
            if logical_index == 4 and new_size > max_time_w:
                self.video_tree.setColumnWidth(4, max_time_w)
            elif logical_index in (0, 1, 2):
                other_fixed = sum(self.video_tree.columnWidth(c) for c in (0, 1, 2) if c != logical_index)
                max_this_col = max(50, viewport_w - (w4 + min_desc_w + other_fixed))
                if new_size > max_this_col:
                    self.video_tree.setColumnWidth(logical_index, max_this_col)
            header.blockSignals(False)
        except Exception:
            pass

    def sync_all_widget_data(self):
        if self.view_mode != "video": return
        for i in range(self.video_tree.topLevelItemCount()):
            group = self.video_tree.topLevelItem(i)
            for c in range(group.childCount()):
                child = group.child(c)
                if "END VIDEO" in child.text(3): continue
                combo_gen = self.video_tree.itemWidget(child, 1)
                if combo_gen: child.setData(1, Qt.UserRole, combo_gen.currentText())
                combo_spec = self.video_tree.itemWidget(child, 2)
                if combo_spec: child.setData(2, Qt.UserRole, combo_spec.currentText())

    def rebind_tree_widgets(self):
        if self.view_mode != "video": return
        for i in range(self.video_tree.topLevelItemCount()):
            group_item = self.video_tree.topLevelItem(i)
            if not self.video_tree.itemWidget(group_item, 4):
                start_time_w = self.create_inline_action_widget("00:00", item=group_item, is_editable=False, on_play=None, on_delete=lambda _, g=group_item: self.delete_video_group(g))
                self.video_tree.setItemWidget(group_item, 4, start_time_w)
            child_count = group_item.childCount()
            for c in range(child_count):
                child = group_item.child(c)
                if "END VIDEO" in child.text(3): continue
                cat_gen = child.data(1, Qt.UserRole) or ""
                cat_spec = child.data(2, Qt.UserRole) or ""
                time_ms = child.data(4, Qt.UserRole) if child.data(4, Qt.UserRole) is not None else 0
                time_str = child.data(4, Qt.UserRole + 1) or self.format_ms(time_ms)
                if not self.video_tree.itemWidget(child, 1):
                    combo_gen = self.create_category_combo(child, "general", cat_gen)
                    self.video_tree.setItemWidget(child, 1, combo_gen)
                if not self.video_tree.itemWidget(child, 2):
                    combo_spec = self.create_category_combo(child, "specific", cat_spec)
                    self.video_tree.setItemWidget(child, 2, combo_spec)
                if not self.video_tree.itemWidget(child, 4):
                    action_widget = self.create_inline_action_widget(time_str, item=child, is_editable=True, on_play=lambda _, g=group_item, item=child: self.switch_active_video(g, item.data(4, Qt.UserRole) or 0), on_delete=lambda _, item=child: self.delete_row_item(item), on_jump=lambda _, item=child: self.jump_to_category(item))
                    self.video_tree.setItemWidget(child, 4, action_widget)

    def push_state(self):
        if self.is_restoring_state or self.view_mode != "video": return
        self.unsaved_changes = True
        state = self.get_current_state()
        self.undo_stack.append(state)
        if len(self.undo_stack) > 50: self.undo_stack.pop(0)
        self.redo_stack.clear()
        self.update_undo_redo_actions()

    def update_undo_redo_actions(self):
        self.undo_action.setEnabled(len(self.undo_stack) > 0 and self.view_mode == "video")
        self.redo_action.setEnabled(len(self.redo_stack) > 0 and self.view_mode == "video")

    def undo(self):
        if not self.undo_stack or self.view_mode != "video": return
        current_state = self.get_current_state()
        self.redo_stack.append(current_state)
        prev_state = self.undo_stack.pop()
        self.restore_state(prev_state)
        self.update_undo_redo_actions()

    def redo(self):
        if not self.redo_stack or self.view_mode != "video": return
        current_state = self.get_current_state()
        self.undo_stack.append(current_state)
        next_state = self.redo_stack.pop()
        self.restore_state(next_state)
        self.update_undo_redo_actions()

    def get_current_state(self):
        if self.view_mode != "video":
            return self.saved_video_state
        groups_data = []
        for i in range(self.video_tree.topLevelItemCount()):
            group = self.video_tree.topLevelItem(i)
            video_path = group.data(0, Qt.UserRole)
            duration_ms = group.data(1, Qt.UserRole)
            rows_data = []
            child_count = group.childCount()
            for c in range(max(0, child_count - 1)):
                child = group.child(c)
                slide = clean_slide_str(child.text(0))
                cat_gen = child.data(1, Qt.UserRole) or ""
                cat_spec = child.data(2, Qt.UserRole) or ""
                desc = child.text(3)
                time_ms = child.data(4, Qt.UserRole) if child.data(4, Qt.UserRole) is not None else 0
                time_str = child.data(4, Qt.UserRole + 1) or self.format_ms(time_ms)
                rows_data.append({
                    "slide": slide, "cat_gen": cat_gen, "cat_spec": cat_spec,
                    "desc": desc, "time_str": time_str, "time_ms": time_ms
                })
            groups_data.append({
                "video_path": video_path, "duration_ms": duration_ms, "rows": rows_data
            })
        return groups_data

    def restore_state(self, groups_data):
        self.is_restoring_state = True
        self.set_time_edit_mode(None)
        self.video_tree.setUpdatesEnabled(False)
        self.video_tree.clear()
        self._scan_and_add_custom_categories(groups_data)
        for g_data in groups_data:
            file_path = g_data.get("video_path", "")
            rows = g_data.get("rows", [])
            dur_ms = g_data.get("duration_ms", 0)
            self.add_video_group(file_path, existing_rows=rows, default_duration_ms=dur_ms)
        found_active = False
        if self.active_video_path:
            for i in range(self.video_tree.topLevelItemCount()):
                group = self.video_tree.topLevelItem(i)
                if group.data(0, Qt.UserRole) == self.active_video_path:
                    self.active_group_item = group
                    found_active = True
                    break
        if not found_active and self.video_tree.topLevelItemCount() > 0:
            self.switch_active_video(self.video_tree.topLevelItem(0), 0)
        else:
            self.renumber_video_groups()
        self.refresh_all_combos()
        self.video_tree.setUpdatesEnabled(True)
        self.video_tree.doItemsLayout()
        self.refresh_playback_ui()
        self.is_restoring_state = False

    def update_play_button_state(self, widget, is_active):
        if not hasattr(widget, 'play_btn'): return
        if getattr(widget, '_is_active_style', None) != is_active:
            widget._is_active_style = is_active
            if is_active:
                widget.play_btn.setStyleSheet("""QPushButton { border: none; background-color: rgba(40, 167, 69, 0.15); color: #28a745; font-size: 14px; font-weight: bold; padding: 2px 4px; border-radius: 4px; }""")
            else:
                widget.play_btn.setStyleSheet("""QPushButton { border: none; background-color: transparent; color: #2563EB; font-size: 14px; font-weight: bold; padding: 2px 4px; border-radius: 4px; } QPushButton:hover { background-color: rgba(37, 99, 235, 0.15); color: #1D4ED8; }""")

    def refresh_playback_ui(self):
        if self.cap and self.cap.isOpened():
            current_ms = int(self.cap.get(cv2.CAP_PROP_POS_MSEC))
            self.highlight_active_row(current_ms)
        else:
            self.highlight_active_row(0)

    def highlight_active_row(self, pos_ms):
        if self.view_mode == "pareto":
            return
        if self.view_mode == "category":
            for g in range(self.category_tree.topLevelItemCount()):
                group = self.category_tree.topLevelItem(g)
                for c in range(group.childCount()):
                    child = group.child(c)
                    w = self.category_tree.itemWidget(child, 3)
                    if not w: continue
                    is_active = (child == getattr(self, "active_category_item", None))
                    self.update_play_button_state(w, is_active)
                    if isinstance(w, TimeColumnWidget):
                        if is_active:
                            start_ms = child.data(1, Qt.UserRole) or 0
                            delta_ms = child.data(3, Qt.UserRole) or 1
                            progress = max(0, pos_ms - start_ms)
                            w.set_fill_ratio(min(1.0, progress / delta_ms))
                        else:
                            w.set_fill_ratio(0.0)
            return

        for g in range(self.video_tree.topLevelItemCount()):
            group = self.video_tree.topLevelItem(g)
            is_active_group = (group == self.active_group_item)
            child_count = group.childCount()
            active_item = None
            if is_active_group:
                # Match the epsilon used in skip_to_segment_start/skip_to_next_segment_start so the
                # highlighted row always agrees with which segment navigation thinks we're in.
                query_ms = pos_ms + self._segment_boundary_epsilon_ms()
                for i in range(child_count - 1):
                    item = group.child(i)
                    next_item = group.child(i + 1)
                    ms1 = item.data(4, Qt.UserRole) or 0
                    ms2 = next_item.data(4, Qt.UserRole) or 0
                    if ms1 <= query_ms < ms2:
                        active_item = item
                        break
                if not active_item and child_count > 0:
                    last_item = group.child(child_count - 1)
                    if query_ms >= (last_item.data(4, Qt.UserRole) or 0):
                        active_item = last_item

            for i in range(child_count):
                item = group.child(i)
                w = self.video_tree.itemWidget(item, 4)
                if not w: continue
                is_active_item = (item == active_item)
                self.update_play_button_state(w, is_active_item)
                if isinstance(w, TimeColumnWidget):
                    if is_active_item and i < child_count - 1:
                        ms_start = item.data(4, Qt.UserRole) or 0
                        ms_end = group.child(i + 1).data(4, Qt.UserRole) or 0
                        seg_dur = max(1, ms_end - ms_start)
                        progress = max(0, pos_ms - ms_start)
                        ratio = min(1.0, progress / seg_dur)
                        w.set_fill_ratio(ratio)
                    elif is_active_item and i == child_count - 1:
                        w.set_fill_ratio(1.0)
                    else:
                        w.set_fill_ratio(0.0)

    def create_inline_action_widget(self, time_str, item=None, is_editable=True, on_play=None, on_delete=None, on_jump=None):
        widget = TimeColumnWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(2, 0, 2, 0)
        layout.setSpacing(4)
        if is_editable and item is not None:
            time_edit = QLineEdit(time_str)
            time_edit.setFixedWidth(55)
            time_edit.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            time_edit.setStyleSheet("QLineEdit { font-size: 14px; font-weight: bold; font-family: monospace; background: transparent; border: 1px solid transparent; color: inherit; padding: 1px 2px; } QLineEdit:focus { background-color: #2563EB; color: #FFFFFF; border-radius: 4px; }")
            time_edit.editingFinished.connect(lambda: self.commit_time_edit(item))
            time_edit.selectionChanged.connect(lambda: self.set_time_edit_mode(item, focus_ui=False))
            layout.addWidget(time_edit)
            widget.time_edit = time_edit
        else:
            lbl = QLabel(time_str)
            lbl.setStyleSheet("font-size: 14px; font-weight: bold; font-family: monospace; background: transparent; color: inherit;")
            layout.addWidget(lbl)
            widget.time_label = lbl

        layout.addStretch()
        if on_play:
            play_btn = QPushButton("▶")
            play_btn.setFocusPolicy(Qt.NoFocus)
            play_btn.setToolTip("Jump to timestamp")
            layout.addWidget(play_btn)
            widget.play_btn = play_btn
            self.update_play_button_state(widget, False)
            play_btn.clicked.connect(on_play)

        if on_jump:
            jump_btn = QPushButton("↪")
            jump_btn.setFocusPolicy(Qt.NoFocus)
            jump_btn.setToolTip("Jump to this row in the other view mode")
            jump_btn.setStyleSheet("""
                QPushButton { border: none; background-color: transparent; color: #000000; font-size: 16px; font-weight: bold; padding: 2px 4px; border-radius: 4px; }
                QPushButton:hover { background-color: rgba(0, 0, 0, 0.08); color: #000000; }
            """)
            layout.addWidget(jump_btn)
            widget.jump_btn = jump_btn
            jump_btn.clicked.connect(on_jump)

        if on_delete:
            del_btn = QPushButton("🗑")
            del_btn.setFocusPolicy(Qt.NoFocus)
            del_btn.setToolTip("Delete item / video group")
            del_btn.setStyleSheet("QPushButton { border: none; background-color: transparent; color: #DC2626; font-size: 14px; padding: 2px 4px; border-radius: 4px; } QPushButton:hover { background-color: rgba(220, 38, 38, 0.15); color: #991B1B; }")
            del_btn.clicked.connect(on_delete)
            layout.addWidget(del_btn)
        return widget

    def apply_chip_style(self, combo):
        text = combo.currentText().strip()
        if text == "+ Add New Category...":
            combo.setStyleSheet("QComboBox { border-radius: 11px; padding: 3px 8px; background-color: #FFFFFF; border: 1px dashed #2563EB; color: #2563EB; font-size: 13px; font-weight: bold; }")
            return
        bg, fg = CATEGORY_COLORS.get(text, ("#F3F4F6", "#374151"))
        if not text:
            combo.setStyleSheet("QComboBox { border-radius: 11px; padding: 3px 8px; background-color: #FFFFFF; border: 1px solid #D1D5DB; color: #9CA3AF; font-size: 14px; }")
        else:
            combo.setStyleSheet(f"QComboBox {{ border-radius: 11px; padding: 3px 10px; background-color: {bg}; color: {fg}; font-weight: bold; font-size: 14px; border: 1px solid {bg}; }} QComboBox::drop-down {{ subcontrol-origin: padding; subcontrol-position: top right; width: 16px; border: none; }} QComboBox::down-arrow {{ image: none; border-left: 4px solid transparent; border-right: 4px solid transparent; border-top: 5px solid {fg}; margin-right: 4px; }}")

    def set_time_edit_mode(self, item, focus_ui=True):
        self.time_editing_item = item
        if focus_ui and item and item.parent() is not None and "END VIDEO" not in item.text(3):
            w = self.video_tree.itemWidget(item, 4)
            if w and hasattr(w, 'time_edit'): w.time_edit.setFocus()

    def commit_time_edit(self, item):
        if not item or not item.parent(): return
        w = self.video_tree.itemWidget(item, 4)
        if w and hasattr(w, 'time_edit'):
            raw_text = w.time_edit.text()
            new_ms = self.parse_time_ms(None, raw_text)
            group_dur = item.parent().data(1, Qt.UserRole) or self.duration_ms
            new_ms = max(0, min(group_dur, new_ms))
            new_str = self.format_ms(new_ms)
            old_ms = item.data(4, Qt.UserRole)
            item.setData(4, Qt.UserRole, new_ms)
            item.setData(4, Qt.UserRole + 1, new_str)
            w.time_edit.setText(new_str)
            if old_ms != new_ms:
                self.seek_position(new_ms)
                self.push_state()
        self.refresh_playback_ui()

    def nudge_editing_time(self, delta_ms):
        item = self.time_editing_item
        if not item or not item.parent(): return
        w = self.video_tree.itemWidget(item, 4)
        if w:
            time_val_str = w.time_edit.text() if hasattr(w, 'time_edit') else item.data(4, Qt.UserRole + 1)
            curr_ms = self.parse_time_ms(item.data(4, Qt.UserRole), time_val_str)
            group_dur = item.parent().data(1, Qt.UserRole) or self.duration_ms
            new_ms = max(0, min(group_dur, curr_ms + delta_ms))
            new_str = self.format_ms(new_ms)
            item.setData(4, Qt.UserRole, new_ms)
            item.setData(4, Qt.UserRole + 1, new_str)
            if hasattr(w, 'time_edit'):
                w.time_edit.setText(new_str)
                w.time_edit.selectAll()
            self.seek_position(new_ms)
            self.push_state()
        self.refresh_playback_ui()

    def on_tree_item_clicked(self, item, column):
        self.jump_wi_to_slide(item.text(0))
        if self.view_mode != "video": return
        if item.parent() is not None and "END VIDEO" not in item.text(3):
            if column == 4:
                self.set_time_edit_mode(item, focus_ui=True)
            elif column in (0, 3):
                self.set_time_edit_mode(None)
                self.video_tree.editItem(item, column)
            else:
                self.set_time_edit_mode(None)
        else:
            self.set_time_edit_mode(None)

    def on_tree_item_double_clicked(self, item, column):
        if self.view_mode != "video": return
        if item.parent() is None:
            self.switch_active_video(item, 0)
        elif column == 4 and "END VIDEO" not in item.text(3):
            self.set_time_edit_mode(item, focus_ui=True)
            group_item = item.parent()
            time_ms = item.data(4, Qt.UserRole) or 0
            self.switch_active_video(group_item, time_ms)
        elif column in (0, 3) and "END VIDEO" not in item.text(3):
            self.set_time_edit_mode(None)
            self.video_tree.editItem(item, column)

    def add_video_group(self, file_path, existing_rows=None, default_duration_ms=0, target_ms=0):
        if not self.is_restoring_state: self.push_state()
        duration_ms = default_duration_ms
        video_available = False
        if file_path and os.path.exists(file_path):
            video_available = True
            if duration_ms <= 0:
                cap = cv2.VideoCapture(file_path)
                if cap.isOpened():
                    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
                    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
                    duration_ms = int((total_frames / fps) * 1000)
                    cap.release()

        group_item = QTreeWidgetItem(self.video_tree)
        group_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled)
        group_item.setData(0, Qt.UserRole, file_path)
        group_item.setData(1, Qt.UserRole, duration_ms)
        filename = os.path.basename(file_path) if file_path else "Unlinked Video"
        status_suffix = "" if video_available else " [Missing Video]"
        group_item.setText(3, f"START VIDEO ({filename}){status_suffix}")
        group_item.setText(4, "")
        group_item.setData(4, Qt.UserRole, 0)
        group_item.setData(4, Qt.UserRole + 1, "00:00")
        green_bg = QBrush(QColor("#DEF7EC"))
        green_fg = QBrush(QColor("#03543F"))
        bold_font = QFont()
        bold_font.setBold(True)
        for c in range(5):
            group_item.setBackground(c, green_bg)
            group_item.setForeground(c, green_fg)
            group_item.setFont(c, bold_font)
            
        start_time_w = self.create_inline_action_widget("00:00", item=group_item, is_editable=False, on_play=None, on_delete=lambda _, g=group_item: self.delete_video_group(g))
        self.video_tree.setItemWidget(group_item, 4, start_time_w)
        
        sorted_rows = sorted(existing_rows or [], key=lambda r: self.parse_time_ms(r.get("time_ms"), r.get("time_str")))
        for idx, r in enumerate(sorted_rows):
            t_str = r.get("time_str", "00:00")
            t_ms = self.parse_time_ms(r.get("time_ms"), t_str)
            t_str = self.format_ms(t_ms)
            child_item = QTreeWidgetItem()
            child_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled)
            group_item.insertChild(idx, child_item)
            child_item.setText(0, clean_slide_str(r.get("slide", "")))
            child_item.setData(1, Qt.UserRole, r.get("cat_gen", ""))
            child_item.setData(2, Qt.UserRole, r.get("cat_spec", ""))
            child_item.setText(3, r.get("desc", ""))
            child_item.setText(4, "")
            child_item.setData(4, Qt.UserRole, t_ms)
            child_item.setData(4, Qt.UserRole + 1, t_str)
            combo_gen = self.create_category_combo(child_item, "general", r.get("cat_gen", ""))
            self.video_tree.setItemWidget(child_item, 1, combo_gen)
            combo_spec = self.create_category_combo(child_item, "specific", r.get("cat_spec", ""))
            self.video_tree.setItemWidget(child_item, 2, combo_spec)
            
            action_widget = self.create_inline_action_widget(
                t_str, item=child_item, is_editable=True, 
                on_play=lambda _, g=group_item, item=child_item: self.switch_active_video(g, item.data(4, Qt.UserRole) or 0), 
                on_delete=lambda _, item=child_item: self.delete_row_item(item),
                on_jump=lambda _, item=child_item: self.jump_to_category(item)
            )
            self.video_tree.setItemWidget(child_item, 4, action_widget)
            
        end_item = QTreeWidgetItem(group_item)
        end_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        end_item.setText(3, "END VIDEO")
        end_time_str = self.format_ms(duration_ms)
        end_item.setText(4, end_time_str)
        end_item.setData(4, Qt.UserRole, duration_ms)
        end_item.setData(4, Qt.UserRole + 1, end_time_str)
        red_bg = QBrush(QColor("#FDE2E2"))
        red_fg = QBrush(QColor("#9B1C1C"))
        for c in range(5):
            end_item.setBackground(c, red_bg)
            end_item.setForeground(c, red_fg)
            end_item.setFont(c, bold_font)
        group_item.setExpanded(True)
        self.renumber_video_groups()
        if video_available and not self.is_restoring_state:
            self.switch_active_video(group_item, target_ms)
        return group_item

    def renumber_video_groups(self):
        for i in range(self.video_tree.topLevelItemCount()):
            group = self.video_tree.topLevelItem(i)
            file_path = group.data(0, Qt.UserRole)
            filename = os.path.basename(file_path) if file_path else "Unlinked Video"
            status = "" if file_path and os.path.exists(file_path) else " [Missing Video]"
            prefix = "★ " if group == self.active_group_item else ""
            group.setText(3, f"{prefix}START VIDEO {i + 1} ({filename}){status}")
            child_count = group.childCount()
            if child_count > 0:
                end_item = group.child(child_count - 1)
                end_item.setText(3, f"END VIDEO {i + 1}")

    def switch_active_video(self, group_item, target_ms=0):
        if not group_item: return
        file_path = group_item.data(0, Qt.UserRole)
        if not file_path or not os.path.exists(file_path):
            new_path, _ = QFileDialog.getOpenFileName(self, f"Locate Missing Video ({os.path.basename(str(file_path))})", self._default_browse_dir(), "Video Files (*.mp4 *.avi *.mkv *.mov)")
            if new_path and os.path.exists(new_path):
                file_path = new_path
                group_item.setData(0, Qt.UserRole, file_path)
            else:
                return
        if self.active_video_path != file_path or not self.cap or not self.cap.isOpened():
            if self.cap: self.cap.release()
            self.active_video_path = file_path
            self.cap = cv2.VideoCapture(file_path)
            self.fps = self.cap.get(cv2.CAP_PROP_FPS) or 30.0
            self.total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
            self.duration_ms = int((self.total_frames / self.fps) * 1000)
            group_item.setData(1, Qt.UserRole, self.duration_ms)
            self.slider.setRange(0, self.duration_ms)
            interval = max(1, int(1000 / (self.fps)))
            self.timer.setInterval(interval)
            self.video_view.reset_zoom()
        self.active_group_item = group_item
        self.renumber_video_groups()
        self.seek_position(target_ms)

    def add_timestamp_row(self):
        if not self.active_group_item or not self.cap:
            QMessageBox.information(self, "No Active Video", "Please load a video first.")
            return
        current_ms = int(self.cap.get(cv2.CAP_PROP_POS_MSEC))
        time_str = self.format_ms(current_ms)
        self.insert_child_row(self.active_group_item, "", "", "", "", time_str, current_ms)

    def insert_child_row(self, group_item, slide, cat_gen, cat_spec, desc, time_str, time_ms):
        if not self.is_restoring_state: self.push_state()
        insert_idx = 0
        child_count = group_item.childCount()
        for i in range(max(0, child_count - 1)):
            child = group_item.child(i)
            existing_ms = child.data(4, Qt.UserRole)
            if existing_ms is None:
                existing_ms = self.parse_time_ms(None, child.data(4, Qt.UserRole + 1))
            if time_ms >= existing_ms:
                insert_idx = i + 1
            else:
                break
        child_item = QTreeWidgetItem()
        child_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled)
        group_item.insertChild(insert_idx, child_item)
        child_item.setText(0, clean_slide_str(slide))
        child_item.setData(1, Qt.UserRole, cat_gen)
        child_item.setData(2, Qt.UserRole, cat_spec)
        child_item.setText(3, desc)
        child_item.setText(4, "")
        child_item.setData(4, Qt.UserRole, time_ms)
        child_item.setData(4, Qt.UserRole + 1, time_str)
        combo_gen = self.create_category_combo(child_item, "general", cat_gen)
        self.video_tree.setItemWidget(child_item, 1, combo_gen)
        combo_spec = self.create_category_combo(child_item, "specific", cat_spec)
        self.video_tree.setItemWidget(child_item, 2, combo_spec)
        
        action_widget = self.create_inline_action_widget(
            time_str, item=child_item, is_editable=True, 
            on_play=lambda _, g=group_item, item=child_item: self.switch_active_video(g, item.data(4, Qt.UserRole) or 0), 
            on_delete=lambda _, item=child_item: self.delete_row_item(item),
            on_jump=lambda _, item=child_item: self.jump_to_category(item)
        )
        self.video_tree.setItemWidget(child_item, 4, action_widget)
        group_item.setExpanded(True)
        self.video_tree.setCurrentItem(child_item)
        self.refresh_playback_ui()

    def delete_row_item(self, item):
        if not item or not item.parent(): return
        if "END VIDEO" in item.text(3):
            QMessageBox.warning(self, "Action Denied", "Cannot delete the END VIDEO anchor line.")
            return
        self.push_state()
        if item == self.time_editing_item: self.set_time_edit_mode(None)
        item.parent().removeChild(item)
        self.refresh_playback_ui()

    def delete_video_group(self, group_item):
        if not group_item: return
        reply = QMessageBox.question(self, "Delete Video Group?", "Are you sure you want to delete this video group and all its timestamps?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.push_state()
            self.set_time_edit_mode(None)
            idx = self.video_tree.indexOfTopLevelItem(group_item)
            self.video_tree.takeTopLevelItem(idx)
            if group_item == self.active_group_item:
                self.active_group_item = self.video_tree.topLevelItem(0) if self.video_tree.topLevelItemCount() > 0 else None
                if self.active_group_item:
                    self.switch_active_video(self.active_group_item, 0)
            self.renumber_video_groups()
            self.refresh_all_combos()
            self.refresh_playback_ui()

    def closeEvent(self, event):
        if self.unsaved_changes:
            reply = QMessageBox.question(
                self, "Unsaved Changes",
                "You have unsaved changes. Do you want to save before exiting?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )
            if reply == QMessageBox.Save:
                if not self.save_project():
                    event.ignore()
                    return
            elif reply == QMessageBox.Cancel:
                event.ignore()
                return
        event.accept()

    def new_project(self):
        if self.unsaved_changes:
            reply = QMessageBox.question(
                self, "Unsaved Changes",
                "You have unsaved changes. Do you want to save your progress before creating a new project?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )
            if reply == QMessageBox.Save:
                if not self.save_project():
                    return
            elif reply == QMessageBox.Cancel:
                return

        self.set_time_edit_mode(None)
        self.video_tree.clear()
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.update_undo_redo_actions()

        self.cat_general_options = self.default_cat_general_options.copy()
        self.cat_specific_options = self.default_cat_specific_options.copy()
        self.custom_gen_cats.clear()
        self.custom_spec_cats.clear()
        self._clear_custom_category_colors()
        
        self.project_path = ""
        self.active_video_path = ""
        self.active_group_item = None
        self.close_work_instruction()
        if self.cap:
            self.cap.release()
            self.cap = None
        self.setWindowTitle("Multi-Video Time Study Logger - [Untitled Project]")
        self.proj_status_label.setText("Project: Unsaved")
        self.time_label.setText("00:00 / 00:00 (◄ / ► Arrow Keys = ±1s)")
        self.unsaved_changes = False

    def _event_belongs_to_wi(self, obj):
        if self.wi_window is None:
            return False
        if QApplication.activeWindow() is self.wi_window:
            return True
        if not isinstance(obj, QWidget):
            return False
        return obj is self.wi_window or self.wi_window.isAncestorOf(obj)

    def open_work_instruction(self):
        if self.wi_path:
            reply = QMessageBox.warning(
                self, "Replace Work Instruction",
                f"This project is already linked to a Work Instruction:\n{os.path.basename(self.wi_path)}\n\n"
                "Opening a new one will unlink it. Continue?",
                QMessageBox.Yes | QMessageBox.Cancel, QMessageBox.Cancel
            )
            if reply != QMessageBox.Yes:
                return
        path, _ = QFileDialog.getOpenFileName(self, "Open Work Instruction PDF", self._default_browse_dir(), "PDF Files (*.pdf)")
        if not path: return
        if self.load_work_instruction(path):
            self.unsaved_changes = True

    def load_work_instruction(self, path, show=True, warn_missing=True):
        if fitz is None:
            QMessageBox.warning(self, "Work Instructions", "PyMuPDF is not installed.\nRun: pip install PyMuPDF")
            return False
        if not path or not os.path.exists(path):
            if warn_missing:
                QMessageBox.warning(self, "Work Instructions", f"Work Instruction PDF not found:\n{path}")
            return False
        if self.wi_window is None:
            self.wi_window = WorkInstructionWindow(self)
        try:
            self.wi_window.load_pdf(path)
        except Exception as e:
            QMessageBox.critical(self, "Work Instructions", f"Failed to open PDF:\n{str(e)}")
            return False
        self.wi_path = path
        if show:
            self.wi_window.show()
            self.wi_window.raise_()
            self.wi_window.activateWindow()
            self.wi_window.focus_viewer()
        return True

    def close_work_instruction(self):
        self.wi_path = ""
        if self.wi_window is not None:
            self.wi_window.close()
            self.wi_window = None

    def jump_wi_to_slide(self, slide_value):
        if self.auto_jump_to_slide and self.wi_window is not None and self.wi_window.doc:
            self.wi_window.goto_slide(slide_value)

    def link_slide_to_selected_row(self, slide_number):
        if self.view_mode != "video":
            QMessageBox.information(self, "Link Slide", "Switch to Video view to link a slide to a row.")
            return
        item = self.video_tree.currentItem()
        if item is None or item.parent() is None or "END VIDEO" in item.text(3):
            QMessageBox.information(self, "Link Slide", "Select a timestamp row first, then link the slide.")
            return
        self.push_state()
        self.video_tree.blockSignals(True)
        item.setText(0, clean_slide_str(slide_number))
        self.video_tree.blockSignals(False)
        self.unsaved_changes = True

    def save_project(self):
        if not self.project_path: return self.save_project_as()
        return self._write_project_file(self.project_path)

    def save_project_as(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Project File", self._default_browse_dir(), "Time Study Project (*.tsproject *.json)")
        if not path: return False
        self.project_path = path
        return self._write_project_file(path)

    @staticmethod
    def _make_relative_path(file_path, project_file_path):
        if not file_path or not project_file_path:
            return file_path
        try:
            return os.path.relpath(file_path, os.path.dirname(project_file_path))
        except ValueError:
            return file_path  # e.g. different drive on Windows

    @staticmethod
    def _resolve_project_path(stored_path, project_file_path):
        if not stored_path or os.path.isabs(stored_path):
            return stored_path
        return os.path.normpath(os.path.join(os.path.dirname(project_file_path), stored_path))

    def _write_project_file(self, path):
        current_ms = int(self.cap.get(cv2.CAP_PROP_POS_MSEC)) if self.cap else 0
        groups_data = []
        for g in self.get_current_state():
            g = dict(g)
            g["video_path"] = self._make_relative_path(g.get("video_path", ""), path)
            groups_data.append(g)
        project_data = {
            "version": 4, "active_video_path": self._make_relative_path(self.active_video_path, path), "last_position_ms": current_ms,
            "work_instruction_path": self._make_relative_path(self.wi_path, path),
            "video_groups": groups_data, "custom_gen_cats": list(self.custom_gen_cats), "custom_spec_cats": list(self.custom_spec_cats),
            "custom_category_colors": {
                cat_name: {"background": colors[0], "foreground": colors[1]}
                for cat_name, colors in self.custom_category_colors.items()
            }
        }
        try:
            with open(path, 'w', encoding='utf-8') as f: json.dump(project_data, f, indent=4)
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to write project file:\n{str(e)}")
            return False
        filename = os.path.basename(path)
        self.setWindowTitle(f"Multi-Video Time Study Logger - {filename}")
        self.proj_status_label.setText(f"Project: {filename}")
        self.unsaved_changes = False
        return True

    def open_project_dialog(self):
        if self.unsaved_changes:
            reply = QMessageBox.question(
                self, "Unsaved Changes",
                "You have unsaved changes. Do you want to save before opening a new project?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )
            if reply == QMessageBox.Save:
                if not self.save_project():
                    return
            elif reply == QMessageBox.Cancel:
                return

        path, _ = QFileDialog.getOpenFileName(self, "Open Project", self._default_browse_dir(), "Time Study Project (*.tsproject *.json);;All Files (*)")
        if not path: return
        self.toggle_view_mode(0)
        progress = QProgressDialog(self)
        progress.setWindowTitle("Loading")
        progress.setLabelText("Loading Project File...")
        progress.setRange(0, 0)
        progress.setCancelButton(None)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()
        QApplication.processEvents()
        try:
            with open(path, 'r', encoding='utf-8') as f: data = json.load(f)
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Open Error", f"Failed to read project file:\n{str(e)}")
            return
        self.set_time_edit_mode(None)
        self.video_tree.clear()
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.update_undo_redo_actions()
        self.project_path = path
        
        self._clear_custom_category_colors()
        self._load_saved_category_colors(data.get("custom_category_colors", {}))
        self._load_custom_categories(data.get("custom_gen_cats", []), data.get("custom_spec_cats", []))

        video_groups = data.get("video_groups", [])
        if not video_groups and ("video_path" in data or "rows" in data):
            video_groups = [{"video_path": data.get("video_path", ""), "duration_ms": 0, "rows": data.get("rows", [])}]
            
        self.is_restoring_state = True
        self.video_tree.setUpdatesEnabled(False)
        
        for g_data in video_groups:
            file_path = self._resolve_project_path(g_data.get("video_path", ""), path)
            rows = g_data.get("rows", [])
            dur_ms = g_data.get("duration_ms", 0)
            self.add_video_group(file_path, existing_rows=rows, default_duration_ms=dur_ms)
            
        self.video_tree.setUpdatesEnabled(True)
        self.is_restoring_state = False
        
        active_path = self._resolve_project_path(data.get("active_video_path") or data.get("video_path", ""), path)
        
        last_ms = data.get("last_position_ms", 0)
        switched = False
        for i in range(self.video_tree.topLevelItemCount()):
            group = self.video_tree.topLevelItem(i)
            if group.data(0, Qt.UserRole) == active_path:
                self.switch_active_video(group, last_ms)
                switched = True
                break
        if not switched and self.video_tree.topLevelItemCount() > 0:
            self.switch_active_video(self.video_tree.topLevelItem(0), last_ms)
        progress.close()

        self.close_work_instruction()
        wi_stored = data.get("work_instruction_path", "")
        if wi_stored:
            resolved_wi_path = self._resolve_project_path(wi_stored, path)
            self.wi_path = resolved_wi_path
            if self.auto_open_wi_on_startup:
                self.load_work_instruction(resolved_wi_path, show=True)

        filename = os.path.basename(path)
        self.setWindowTitle(f"Multi-Video Time Study Logger - {filename}")
        self.proj_status_label.setText(f"Project: {filename}")
        self.unsaved_changes = False

    def import_project_dialog(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import Project File", self._default_browse_dir(), "Time Study Project (*.tsproject *.json);;All Files (*)")
        if not path: return
        self.toggle_view_mode(0)
        progress = QProgressDialog(self)
        progress.setWindowTitle("Importing")
        progress.setLabelText("Parsing Project File...")
        progress.setRange(0, 0)
        progress.setCancelButton(None)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()
        QApplication.processEvents()
        try:
            with open(path, 'r', encoding='utf-8') as f: data = json.load(f)
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Import Error", f"Failed to read project file:\n{str(e)}")
            return
        video_groups = data.get("video_groups", [])
        if not video_groups and ("video_path" in data or "rows" in data):
            video_groups = [{"video_path": data.get("video_path", ""), "duration_ms": 0, "rows": data.get("rows", [])}]
        progress.close()
        if self.video_tree.topLevelItemCount() > 0:
            msg = QMessageBox(self)
            msg.setWindowTitle("Import Project Options")
            msg.setText("Choose how you want to import this project:")
            btn_replace = msg.addButton("Replace Active Session", QMessageBox.AcceptRole)
            btn_merge = msg.addButton("Merge Groups into Current Session", QMessageBox.ActionRole)
            btn_cancel = msg.addButton(QMessageBox.Cancel)
            msg.exec_()
            clicked = msg.clickedButton()
            if clicked == btn_cancel: return
            elif clicked == btn_replace:
                self.set_time_edit_mode(None)
                self.video_tree.clear()
                self.undo_stack.clear()
                self.redo_stack.clear()
                self.update_undo_redo_actions()
                self.project_path = path
        progress = QProgressDialog(self)
        progress.setWindowTitle("Importing")
        progress.setLabelText("Importing Project Data...")
        progress.setRange(0, 0)
        progress.setCancelButton(None)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()
        QApplication.processEvents()
        
        self._load_saved_category_colors(data.get("custom_category_colors", {}))
        self._load_custom_categories(data.get("custom_gen_cats", []), data.get("custom_spec_cats", []))
        
        self.is_restoring_state = True
        self.video_tree.setUpdatesEnabled(False)
        
        for g_data in video_groups:
            file_path = self._resolve_project_path(g_data.get("video_path", ""), path)
            rows = g_data.get("rows", [])
            dur_ms = g_data.get("duration_ms", 0)
            self.add_video_group(file_path, existing_rows=rows, default_duration_ms=dur_ms)
            
        self.video_tree.setUpdatesEnabled(True)
        self.is_restoring_state = False
        
        if not self.project_path: self.project_path = path
        
        progress.close()
        filename = os.path.basename(path)
        self.setWindowTitle(f"Multi-Video Time Study Logger - {filename}")
        self.proj_status_label.setText(f"Project: {filename}")

    def export_to_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export to CSV", self._default_browse_dir(), "CSV Files (*.csv)")
        if not path:
            return
        if not path.lower().endswith(".csv"):
            path += ".csv"

        groups_data = self.saved_video_state if self.view_mode != "video" else self.get_current_state()

        with open(path, "w", newline="", encoding="utf-8-sig") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(["Slide", "Category, General", "Category, Specific", "Description", "Time"])

            for index, group in enumerate(groups_data, start=1):
                writer.writerow(["", "", "", f"START VIDEO {index}", "00:00"])
                for row in group.get("rows", []):
                    writer.writerow([
                        row.get("slide", ""),
                        row.get("cat_gen", ""),
                        row.get("cat_spec", ""),
                        row.get("desc", ""),
                        row.get("time_str", "00:00")
                    ])
                writer.writerow([
                    "", "", "", f"END VIDEO {index}",
                    self.format_ms(group.get("duration_ms", 0))
                ])

        QMessageBox.information(self, "Export Complete", f"Exported successfully to:\n{path}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TimeStudyApp()
    window.show()
    sys.exit(app.exec_())