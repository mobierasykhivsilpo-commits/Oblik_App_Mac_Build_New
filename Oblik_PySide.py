import sys
import os
import re
import glob
import ctypes
import json
from datetime import datetime, timedelta

import pandas as pd

from PySide6.QtCore import (
    Qt, QTimer, QPoint, QSize, QEvent
)
from PySide6.QtGui import (
    QBrush, QAction, QIcon, QCursor, QGuiApplication, QFont, QTextCursor, QPalette, QColor
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QTreeWidget, QTreeWidgetItem, QHeaderView, QLineEdit, QLabel, QPushButton,
    QFileDialog, QMessageBox, QTextEdit, QMenu, QSplitter, QFrame, QStyleFactory,
    QComboBox, QDialog, QDialogButtonBox, QSizePolicy, QScrollArea, QAbstractItemView, QStyledItemDelegate
)


# ============================ DELEGATE FOR EXTRA ROW HEIGHT ============================
class RowPaddingDelegate(QStyledItemDelegate):
    def __init__(self, extra_px=6, parent=None):
        super().__init__(parent)
        self.extra_px = extra_px

    def sizeHint(self, option, index):
        sz = super().sizeHint(option, index)
        try:
            sz.setHeight(sz.height() + int(self.extra_px))
        except Exception:
            pass
        return sz


# ============================ УТИЛІТИ ============================

def read_excel_any(path: str) -> pd.DataFrame:
    """Зчитування xls/xlsx/xlsm з явним вибором engine і fallback-логікою."""
    try:
        if path.lower().endswith(".xls"):
            return pd.read_excel(path, engine="xlrd", header=None)
        else:
            return pd.read_excel(path, engine="openpyxl", header=None)
    except Exception:
        # fallback без engine
        return pd.read_excel(path, header=None)


def format_number(val):
    if pd.isna(val):
        return ""
    try:
        num = float(val)
        return str(int(round(num)))
    except Exception:
        return str(val)


def extract_date_from_filename(filename: str) -> datetime:
    """
    Витягує дату у форматах: dd.mm(.yyyy), dd,mm(,yy), dd-mm(-yy) тощо.
    Якщо немає — повертає datetime.min
    """
    match = re.search(r'(\d{1,2})[.,-](\d{1,2})(?:[.,-](\d{2,4})?)', filename)
    if not match:
        return datetime.min
    day, month, year = match.groups()
    day = int(day)
    month = int(month)
    year = int(year) if year else datetime.now().year
    if year < 100:
        year += 2000
    try:
        return datetime(year=year, month=month, day=day)
    except ValueError:
        return datetime.min


# ============================ ДІАЛОГ МАПІНГУ КОЛОНОК ============================

class ColumnMappingDialog(QDialog):
    def __init__(self, parent, current_mapping):
        super().__init__(parent)
        self.setWindowTitle("Налаштування стовпців")
        self.setModal(True)
        self.resize(360, 240)

        self.labels = ["Найменування", "Закуп", "Приб", "Ціна", "Код", "Арт"]
        self.boxes = {}
        layout = QVBoxLayout(self)

        grid = QGridLayout()
        for i, name in enumerate(self.labels):
            grid.addWidget(QLabel(name, self), i, 0)
            cb = QComboBox(self)
            cb.setEditable(False)
            # дозволимо 0..30 як номери сирих колонок у файлі
            cb.addItems([str(x) for x in range(31)])
            cb.setCurrentText(str(current_mapping.get(name, 0)))
            self.boxes[name] = cb
            grid.addWidget(cb, i, 1)
        layout.addLayout(grid)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def mapping(self):
        mp = {}
        for k, cb in self.boxes.items():
            mp[k] = int(cb.currentText())
        return mp


# ============================ ГОЛОВНЕ ВІКНО ============================

class NomenklaturaApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Заголовок/іконка
        self.setWindowTitle("Облік товарів MobiEra")
        self._load_icon()

        # Стан/поля
        self.df = None
        self.stocks_df = None
        self.current_file = None
        self.current_stocks_file = None
        self.stores = []
        self.history = []
        self.search_delay_ms = 1000
        self.last_query = None
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(lambda: self.search_items(record_history=True))
        self.showing_stocks = False
        self.history_file_path = os.path.join(os.path.expanduser("~"), ".oblpy_history")

        self.column_mapping = {
            "Найменування": 0,
            "Закуп": 1,
            "Приб": 4,
            "Ціна": 5,
            "Код": 7,
            "Арт": 6,
        }

        self.column_order = ["Найменування", "Закуп", "Приб", "Ціна", "Код", "Арт"]
        self.default_column_widths = {
            "Найменування": 500,
            "Закуп": 40,
            "Приб": 40,
            "Ціна": 40,
            "Код": 105,
            "Арт": 105,
        }
        self.column_visibility = {c: True for c in self.column_order}
        # Hide by default
        if 'Закуп' in self.column_visibility:
            self.column_visibility['Закуп'] = False
        if 'Арт' in self.column_visibility:
            self.column_visibility['Арт'] = False

        # Шрифти
        self.font_normal = QFont("Calibri", 11)
        self.font_bold = QFont("Calibri", 11, QFont.Weight.Bold)

        # UI
        self._build_ui()
        self._apply_modern_style()

        QApplication.instance().installEventFilter(self)
        # Автозавантаження
        QTimer.singleShot(150, self.auto_load_files)
        self.load_history()

    def closeEvent(self, event):
        self.save_history()
        event.accept()

    # ----------------------- ІНІЦІАЛІЗАЦІЯ UI -----------------------

    def _build_ui(self):
        QApplication.setStyle(QStyleFactory.create("Fusion"))

        central = QWidget(self)
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(8)

        # Верхній ряд: кнопки імпорту + пошук
        top = QHBoxLayout()
        root_layout.addLayout(top)

        # Ліва частина (Імпорт)
        left = QHBoxLayout()
        top.addLayout(left, 1)

        self.btn_import = QPushButton("Облік", self)
        self.btn_import.setStyleSheet("padding: 2px 8px; font-size: 11px; height: 26px;")
        self.btn_import.setFixedHeight(22)
        self.btn_import.clicked.connect(lambda: self.import_excel(False))
        left.addWidget(self.btn_import)

        self.file_label = QLabel("", self)
        self.file_label.setStyleSheet("color:#0078D4;")
        left.addWidget(self.file_label)

        self.btn_import_stocks = QPushButton("Залишки", self)
        self.btn_import_stocks.setStyleSheet("padding: 2px 8px; font-size: 11px; height: 26px;")
        self.btn_import_stocks.setFixedHeight(22)
        self.btn_import_stocks.clicked.connect(self.import_stocks)
        left.addWidget(self.btn_import_stocks)

        self.stocks_label = QLabel("", self)
        self.stocks_label.setStyleSheet("color:#0078D4;")
        left.addWidget(self.stocks_label)
        left.addStretch()

        # Права частина (Пошук)
        right = QHBoxLayout()
        top.addLayout(right, 0)
        right.addWidget(QLabel("Пошук:", self))
        self.search_edit = QLineEdit(self)
        self.search_edit.setFixedHeight(22)
        self.search_edit.setStyleSheet("padding: 2px 6px; font-size: 11px;")
        self.search_edit.setFixedWidth(240)
        self.search_edit.returnPressed.connect(lambda: (self.search_timer.stop(), self.search_items(True), self.search_edit.selectAll()))
        self.search_edit.textEdited.connect(lambda _: self._debounce_search())
        self.search_edit.installEventFilter(self)
        right.addWidget(self.search_edit)

        # Спліттер: Таблиця + (історія/залишки)
        splitter = QSplitter(Qt.Vertical, self)
        root_layout.addWidget(splitter, 1)

        # Таблиця
        table_container = QWidget(self)
        t_layout = QVBoxLayout(table_container)
        t_layout.setContentsMargins(0, 0, 0, 0)
        t_layout.setSpacing(6)

        self.tree = QTreeWidget(self)
        self.tree.setFont(self.font_normal)
        self.tree.setColumnCount(len(self.column_order))
        self.tree.setHeaderLabels(self.column_order)

        # Центруємо заголовки колонок
        header_item = self.tree.headerItem()
        for i in range(len(self.column_order)):
            header_item.setTextAlignment(i, Qt.AlignCenter)

        # Навігація та вигляд
        self.tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tree.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tree.setFocusPolicy(Qt.StrongFocus)
        self.tree.setItemDelegate(RowPaddingDelegate(6, self.tree))
        self.tree.viewport().setFocusPolicy(Qt.StrongFocus)
        self.tree.setAlternatingRowColors(False)
        self.tree.setRootIsDecorated(False)
        self.tree.setUniformRowHeights(True)
        self.tree.itemClicked.connect(self.on_tree_item_clicked)
        self.tree.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_tree_context_menu)

        # Розміри колонок
        header = self.tree.header()
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        header.customContextMenuRequested.connect(self.show_header_menu)
        header.setStretchLastSection(False)
        for i, col in enumerate(self.column_order):
            if col == "Найменування":
                header.setSectionResizeMode(i, QHeaderView.Stretch)
            else:
                header.setSectionResizeMode(i, QHeaderView.Interactive)
            w = self.default_column_widths.get(col, 100)
            self.tree.setColumnWidth(i, w)
        # Початкова видимість
        for i, col in enumerate(self.column_order):
            self.tree.setColumnHidden(i, not self.column_visibility.get(col, True))

        self.tree.setStyleSheet(self.tree.styleSheet() + """
            QHeaderView::section {
                background-color: #F3F4F6;
                border: 1px solid #C9C9CF;
                border-radius: 4px;
                padding-left: 4px;
                padding-right: 4px;
                min-height: 20px;
            }
            QHeaderView::section:hover {
                background-color: #E9EAEE;
            }
            QHeaderView::section:pressed {
                background-color: #E1E2E7;
            }
        """)

        t_layout.addWidget(self.tree)

        # Ряд з «швидкими» кнопками пошуку
        btn_row = QHBoxLayout()
        t_layout.addLayout(btn_row)
        btn_row.addStretch()
        quick_texts = ["Балон", "Набій", "Шарік", "Ремінець", "Деколь",
                       "Клей", "Розет", "Ст. П.", "Hydrogel", "Батарейка"]
        for txt in quick_texts:
            b = QPushButton(txt, self)
            b.setProperty("preset", True)
            b.style().unpolish(b)
            b.style().polish(b)
            b.setFont(self.font_normal)
            b.clicked.connect(lambda _=False, t=txt: self.set_search_text(t))
            btn_row.addWidget(b)
        btn_row.addStretch()

        splitter.addWidget(table_container)

        # Нижня частина: (Заголовок + кнопка перемикач) + (історія/залишки)
        bottom_container = QWidget(self)
        bottom_layout = QVBoxLayout(bottom_container)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(6)

        head_row = QHBoxLayout()
        bottom_layout.addLayout(head_row)

        self.history_label = QLabel("Історія пошуку:", self)
        self.history_label.setFont(self.font_bold)
        head_row.addWidget(self.history_label)

        self.toggle_history_btn = QPushButton("Залишки", self)
        self.toggle_history_btn.setFixedHeight(22)
        self.toggle_history_btn.setStyleSheet("padding: 2px 6px; font-size: 10pt;")
        self.toggle_history_btn.setFixedWidth(80)
        self.toggle_history_btn.clicked.connect(self.toggle_history_stocks)
        head_row.addWidget(self.toggle_history_btn)
        head_row.addStretch()

        # Стек з історією та панеллю залишків
        self.history_text = QTextEdit(self)
        self.history_text.setReadOnly(True)
        self.history_text.setContextMenuPolicy(Qt.CustomContextMenu)
        self.history_text.customContextMenuRequested.connect(self.show_history_menu)
        self.history_text.viewport().installEventFilter(self)

        self.stocks_panel = ScrollPane(self)  # скрол-панель
        self.stocks_panel.setVisible(False)

        bottom_layout.addWidget(self.history_text, 1)
        bottom_layout.addWidget(self.stocks_panel, 1)

        splitter.addWidget(bottom_container)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        # Мін. розміри та позиціювання
        self._setup_window_position()
        self.search_edit.setFocus()
        self.search_edit.selectAll()

    def _apply_modern_style(self):
        """Apply a clean Windows 10/11-like Fusion palette and compact stylesheet."""
        pal = QPalette()
        pal.setColor(QPalette.Window, QColor("#F7F8FA"))
        pal.setColor(QPalette.WindowText, QColor("#111827"))
        pal.setColor(QPalette.Base, QColor("#FFFFFF"))
        pal.setColor(QPalette.AlternateBase, QColor("#F3F4F6"))
        pal.setColor(QPalette.ToolTipBase, QColor("#111827"))
        pal.setColor(QPalette.ToolTipText, QColor("#FFFFFF"))
        pal.setColor(QPalette.Text, QColor("#111827"))
        pal.setColor(QPalette.Button, QColor("#F3F4F6"))
        pal.setColor(QPalette.ButtonText, QColor("#111827"))
        pal.setColor(QPalette.BrightText, QColor("#FFFFFF"))
        pal.setColor(QPalette.Highlight, QColor("#2D7DFF"))
        pal.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
        QApplication.setPalette(pal)

        self.setStyleSheet("""
        QWidget { font-size: 10pt; }
        QMainWindow { background: #F7F8FA; }

        QLineEdit, QTextEdit, QTreeWidget, QScrollArea {
            border: 1px solid #D5D7DB;
            border-radius: 6px;
            background: #FFFFFF;
        }
        QLineEdit { padding: 4px 6px; }
        QTextEdit { padding: 4px; }
        QLineEdit:focus, QTextEdit:focus {
            border: 1px solid #2D7DFF;
        }

        QPushButton {
            padding: 4px 8px;
            border-radius: 6px;
            border: 1px solid #C9C9CF;
            background: #F3F4F6;
        }

        /* Tighter padding for preset buttons */
        QPushButton[preset="true"] {
            padding: 4px 1px;
        }
        QPushButton:hover { background: #E9EAEE; }
        QPushButton:pressed { background: #E1E2E7; }

        QComboBox {
            padding: 4px 6px;
            border: 1px solid #C9C9CF;
            border-radius: 6px;
            background: #FFFFFF;
        }
        QComboBox::drop-down { border: none; }

        QMenu {
            border: 1px solid #D5D7DB;
            border-radius: 6px;
            padding: 4px;
            background: #FFFFFF;
        }
        QMenu::item { padding: 4px 8px; border-radius: 4px; }
        QMenu::item:selected { background: #E6EFFF; }

        QSplitter::handle {
            background: #E5E7EB;
            margin: 2px;
            border-radius: 2px;
        }
    """)

    # ----------------------- ПОДІЇ/ОБРОБНИКИ -----------------------

    def eventFilter(self, source, event):
        # Автовиділення у полі пошуку при фокусі і кліку
        if source == self.search_edit and event.type() in (QEvent.FocusIn, QEvent.MouseButtonPress):
            QTimer.singleShot(0, self.search_edit.selectAll)

        # Копіювання по кліку в історії (якщо є)
        if hasattr(self, "history_text") and source == self.history_text.viewport() and event.type() == QEvent.MouseButtonPress:
            if not self.showing_stocks:
                pos = event.position().toPoint()
                self.copy_from_history_click(pos)

        # Глобальний клік -> фокус у поле пошуку (окрім інтерактивних елементів)
        try:
            from PySide6.QtWidgets import (
                QLineEdit, QTextEdit, QPlainTextEdit,
                QSpinBox, QDoubleSpinBox, QComboBox,
                QPushButton, QAbstractButton, QAbstractItemView
            )
            if event.type() == QEvent.MouseButtonRelease:
                safe_widgets = (QLineEdit, QTextEdit, QPlainTextEdit,
                                QSpinBox, QDoubleSpinBox,
                                QPushButton, QAbstractButton, QAbstractItemView)
                is_safe = isinstance(source, safe_widgets)
                if isinstance(source, QComboBox) and source.isEditable():
                    is_safe = True
                if not is_safe:
                    QTimer.singleShot(0, lambda: (self.search_edit.setFocus(), self.search_edit.selectAll()))
        except Exception:
            pass

        return super().eventFilter(source, event)

    def _debounce_search(self):
        self.search_timer.start(self.search_delay_ms)

    def set_search_text(self, text):
        self.search_edit.setText(text)
        self.search_items(record_history=True)
        self.search_edit.selectAll()

    def on_tree_item_clicked(self, item: QTreeWidgetItem, column: int):
        # Копіюємо значення в буфер
        text = item.text(column)
        QApplication.clipboard().setText(text)
        # Якщо показуємо залишки — оновити праву панель
        if self.showing_stocks:
            self.update_stocks_display()

    # ----------------------- КОНТЕКСТНІ МЕНЮ -----------------------

    def show_tree_context_menu(self, point: QPoint):
        item = self.tree.itemAt(point)
        if not item:
            return
        col = self.tree.currentColumn()
        menu = QMenu(self)

        act_copy_no_code = QAction("Копіювати строку", self)
        def copy_row_without_code():
            try:
                values = [item.text(i) for i in range(self.tree.columnCount())]
                name_val = values[self.column_order.index("Найменування")] if "Найменування" in self.column_order else ""
                price_val = format_number(values[self.column_order.index("Ціна")]) if "Ціна" in self.column_order else ""
                profit_val = format_number(values[self.column_order.index("Приб")]) if "Приб" in self.column_order else ""
                text = f"{name_val}\t1\t{price_val}\t{profit_val}"
                QApplication.clipboard().setText(text)
            except Exception:
                QApplication.clipboard().setText("")
        act_copy_no_code.triggered.connect(copy_row_without_code)
        menu.addAction(act_copy_no_code)

        act_copy = QAction("Копіювати строку (з кодом)", self)
        def copy_full_row():
            try:
                values = [item.text(i) for i in range(self.tree.columnCount())]
                code_val = values[self.column_order.index("Код")] if "Код" in self.column_order else ""
                name_val = values[self.column_order.index("Найменування")] if "Найменування" in self.column_order else ""
                price_val = format_number(values[self.column_order.index("Ціна")]) if "Ціна" in self.column_order else ""
                profit_val = format_number(values[self.column_order.index("Приб")]) if "Приб" in self.column_order else ""
                text = f"{code_val}\t{name_val}\t1\t{price_val}\t{profit_val}"
                QApplication.clipboard().setText(text)
            except Exception:
                QApplication.clipboard().setText("")
        act_copy.triggered.connect(copy_full_row)
        menu.addAction(act_copy)

        menu.exec(QCursor.pos())

    def show_header_menu(self, point: QPoint):
        header = self.tree.header()
        section = header.logicalIndexAt(point)
        if section < 0:
            return
        menu = QMenu(self)
        # Чекбокси видимості
        for i, col in enumerate(self.column_order):
            label = col
            if col == "Арт":
                label = "Арт (Артикул)"
            elif col == "Приб":
                label = "Приб (Прибуток)"
            act = QAction(label, self, checkable=True)
            act.setChecked(self.column_visibility[col])

            def toggler(checked, idx=i, name=col):
                self.column_visibility[name] = checked
                self.tree.setColumnHidden(idx, not checked)
                if checked:
                    self.tree.setColumnWidth(idx, self.default_column_widths.get(name, 100))

            act.toggled.connect(toggler)
            menu.addAction(act)

        menu.exec(QCursor.pos())

    def show_history_menu(self, point: QPoint):
        if self.showing_stocks:
            return
        menu = QMenu(self)
        act_copy = QAction("Копіювати", self)
        act_copy.triggered.connect(self.copy_from_history_selection)
        act_paste = QAction("Вставити", self)
        act_paste.triggered.connect(self.paste_into_search_from_clipboard)
        act_clear = QAction("Очистити історію", self)
        act_clear.triggered.connect(self.clear_history)
        menu.addAction(act_copy)
        menu.addAction(act_paste)
        menu.addAction(act_clear)
        menu.exec(QCursor.pos())

    # ----------------------- ІСТОРІЯ -----------------------

    def log_action(self, message: str):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        self.history.insert(0, f"{timestamp} - {message}")
        # Обмежуємо історію
        self.history = self.history[:100]
        self.update_history_display()

    def update_history_display(self):
        if self.showing_stocks:
            return
        self.history_text.clear()
        cursor = self.history_text.textCursor()
        for entry in self.history:
            parts = entry.split(" - ", 1)
            if len(parts) == 2:
                ts, info = parts
                # Відображаємо тільки час
                time_part = datetime.strptime(ts, "%Y-%m-%d %H:%M").strftime("%H:%M")
                self._append_text(cursor, time_part + " - ", bold=False)
                if "➔" in info:
                    segs = [s.strip() for s in info.split("➔")]
                    if segs:
                        # перший сегмент — пошуковий запит
                        self._append_text(cursor, segs[0], bold=True)
                        # далі з кольорами: 1 — Приб (зелений), 2 — Ціна (жовтий)
                        for idx, s in enumerate(segs[1:], start=1):
                            self._append_text(cursor, " ➔ ", bold=False)
                            if idx == 1:
                                # Приб
                                self._append_text(cursor, s, bold=False, bg="#baf2ba")
                            elif idx == 2:
                                # Ціна
                                self._append_text(cursor, s, bold=False, bg="#fff2a8")
                            else:
                                self._append_text(cursor, s, bold=False)
                else:
                    self._append_text(cursor, info, bold=True)
                self._append_text(cursor, "\n", bold=False)
            else:
                self._append_text(cursor, entry + "\n", bold=False)
        self.history_text.moveCursor(QTextCursor.Start)

    def save_history(self):
        try:
            # Фільтруємо записи, старіші за 2 дні
            two_days_ago = datetime.now() - timedelta(days=2)
            recent_history = [
                entry for entry in self.history
                if datetime.strptime(entry.split(" - ")[0], "%Y-%m-%d %H:%M") > two_days_ago
            ]
            with open(self.history_file_path, "w", encoding="utf-8") as f:
                json.dump(recent_history, f)
        except Exception as e:
            print(f"Error saving history: {e}")

    def load_history(self):
        try:
            if os.path.exists(self.history_file_path):
                with open(self.history_file_path, "r", encoding="utf-8") as f:
                    saved_history = json.load(f)
                    # Фільтруємо записи, старіші за 2 дні
                    two_days_ago = datetime.now() - timedelta(days=2)
                    self.history = [
                        entry for entry in saved_history
                        if datetime.strptime(entry.split(" - ")[0], "%Y-%m-%d %H:%M") > two_days_ago
                    ]
                self.update_history_display()
        except Exception as e:
            print(f"Error loading history: {e}")
            self.history = []
        finally:
            self.log_action("Історію завантажено")

    def _append_text(self, cursor: QTextCursor, text: str, bold: bool, bg: str = None):
        fmt = cursor.charFormat()
        font = self.font_bold if bold else self.font_normal
        fmt.setFont(font)
        if bg:
            fmt.setBackground(QColor(bg))
        else:
            fmt.clearBackground()
        cursor.setCharFormat(fmt)
        cursor.insertText(text)

    def copy_from_history_selection(self):
        cursor = self.history_text.textCursor()
        text = cursor.selectedText()
        if not text:
            return
        # як і в Tk: копіюємо перший сегмент до '➔' без часу
        if " - " in text:
            text = text.split(" - ", 1)[1]
        if "➔" in text:
            text = text.split("➔")[0].strip()
        QApplication.clipboard().setText(text)

    
    
    def copy_from_history_click(self, pos: QPoint):
        """Копіює або код (перший сегмент), або повну назву (останній сегмент),
        залежно від того, де клацнув користувач."""
        # Знаходимо абзац (весь рядок історії)
        cursor = self.history_text.cursorForPosition(pos)
        cursor.movePosition(QTextCursor.StartOfBlock)
        cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
        line = cursor.selectedText().replace("\u2029", " ").strip()
        if not line:
            return

        # Розбиваємо на сегменти
        segments = [s.strip() for s in line.split("➔")]

        # Якщо є код/назва
        if len(segments) > 1:
            word_cursor = self.history_text.cursorForPosition(pos)
            word_cursor.select(QTextCursor.WordUnderCursor)
            word = word_cursor.selectedText().strip()

            if word and word in segments[0]:
                # якщо клікнули по коді → копіюємо код
                chosen = segments[0].split(" - ", 1)[-1].strip()
            else:
                # інакше → повна назва (останній сегмент)
                chosen = segments[-1].strip()
        else:
            # якщо без сегментів → просто прибираємо час
            if " - " in segments[0]:
                chosen = segments[0].split(" - ", 1)[1].strip()
            else:
                chosen = segments[0]

        QApplication.clipboard().setText(chosen)

    def paste_into_search_from_clipboard(self):
        text = QApplication.clipboard().text()
        if text:
            self.search_edit.setText(text)
            self.search_items(record_history=True)

    def clear_history(self):
        self.history.clear()
        self.update_history_display()
        self.save_history() # Очищаємо файл історії
        self.log_action("Історію пошуку очищено")

    # ----------------------- ПОШУК/ВІДОБРАЖЕННЯ -----------------------

    def search_items(self, record_history=False):
        if self.df is None or not self.column_mapping:
            return

        query = self.search_edit.text().strip()
        if query:
            try:
                conds = []
                for col in ["Найменування", "Код", "Арт"]:
                    idx = self.column_mapping.get(col)
                    if idx is not None:
                        series = self.df.iloc[:, idx].astype(str).str.lower()
                        conds.append(series.str.contains(query.lower(), na=False, regex=False))
                if conds:
                    mask = pd.concat(conds, axis=1).any(axis=1)
                    results = self.df[mask]
                else:
                    results = self.df

                if record_history and query != self.last_query:
                    if "Код" in self.column_mapping and len(results) == 1:
                        nm_idx = self.column_mapping.get("Найменування")
                        pr_idx = self.column_mapping.get("Приб")
                        prc_idx = self.column_mapping.get("Ціна")
                        code_idx = self.column_mapping.get("Код")

                        name = str(results.iloc[0, nm_idx]) if nm_idx is not None else ""
                        price = format_number(results.iloc[0, prc_idx]) if prc_idx is not None else ""
                        profit = format_number(results.iloc[0, pr_idx]) if pr_idx is not None else ""
                        code = str(results.iloc[0, code_idx]) if code_idx is not None else ""
                        if code.endswith(".0"):
                            code = code.replace(".0", "")

                        code_cell = str(results.iloc[0, code_idx]) if code_idx is not None else ""
                        if code_cell.replace(".0", "").lower().find(query.lower()) != -1:
                            self.log_action(f"{query} ➔ {profit} ➔ {price} ➔ {name}")
                        else:
                            self.log_action(f"{query} ➔ {profit} ➔ {price} ➔ {code}")
                    else:
                        self.log_action(query)
                self.last_query = query
            except Exception as e:
                QMessageBox.critical(self, "Помилка пошуку", f"Не вдалося виконати пошук:\n{e}")
                results = self.df
        else:
            results = self.df

        self.show_results(results)
        if self.showing_stocks:
            self.update_stocks_display()

    def show_results(self, df: pd.DataFrame):
        self.tree.clear()
        if df is None:
            return

        # Вставка рядків у порядку колонок
        for _, row in df.iterrows():
            values = []

            code_value = ""
            if "Код" in self.column_mapping:
                code_value = str(row.iloc[self.column_mapping["Код"]])
                if code_value.endswith(".0"):
                    code_value = code_value.replace(".0", "")

            # розрахунок "Закуп" як (Ціна - Приб)
            try:
                price = float(row.iloc[self.column_mapping["Ціна"]]) if "Ціна" in self.column_mapping else 0
                profit = float(row.iloc[self.column_mapping["Приб"]]) if "Приб" in self.column_mapping else 0
                purchase = price - profit
            except Exception:
                purchase = ""

            values = [
                str(row.iloc[self.column_mapping["Найменування"]]) if "Найменування" in self.column_mapping else "",
                format_number(purchase) if purchase != "" else "",
                format_number(row.iloc[self.column_mapping["Приб"]]) if "Приб" in self.column_mapping else "",
                format_number(row.iloc[self.column_mapping["Ціна"]]) if "Ціна" in self.column_mapping else "",
                code_value,
                str(row.iloc[self.column_mapping["Арт"]]) if "Арт" in self.column_mapping else "",
            ]

            item = QTreeWidgetItem(values)
            try:
                col_index = self.column_order.index("Ціна")
                item.setBackground(col_index, QBrush(QColor("#fff2a8")))  # Yellow for Ціна
            except Exception:
                pass
            try:
                profit_index = self.column_order.index("Приб")
                item.setBackground(profit_index, QBrush(QColor("#baf2ba")))  # Green for Приб
            except Exception:
                pass
            try:
                zakup_index = self.column_order.index("Закуп")
                item.setBackground(zakup_index, QBrush(QColor("#ffcccc")))  # Red for Закуп
            except Exception:
                pass

            # Вирівнюємо по центру значення у числових колонках
            try:
                if "Закуп" in self.column_order:
                    item.setTextAlignment(self.column_order.index("Закуп"), Qt.AlignCenter)
                if "Приб" in self.column_order:
                    item.setTextAlignment(self.column_order.index("Приб"), Qt.AlignCenter)
                if "Ціна" in self.column_order:
                    item.setTextAlignment(self.column_order.index("Ціна"), Qt.AlignCenter)
            except Exception as e:
                print("Align numeric cols error:", e)
            self.tree.addTopLevelItem(item)

        # Видимість і ширини
        for i, name in enumerate(self.column_order):
            self.tree.setColumnHidden(i, not self.column_visibility.get(name, True))
            if self.column_visibility.get(name, True):
                self.tree.setColumnWidth(i, self.default_column_widths.get(name, 100))

    # ----------------------- ЗАЛИШКИ -----------------------

    def toggle_history_stocks(self):
        self.showing_stocks = not self.showing_stocks
        if self.showing_stocks:
            self.history_label.setText("Залишки:")
            self.toggle_history_btn.setText("Історія")
            self.history_text.setVisible(False)
            self.stocks_panel.setVisible(True)
            self.update_stocks_display()
        else:
            self.history_label.setText("Історія пошуку:")
            self.toggle_history_btn.setText("Залишки")
            self.stocks_panel.setVisible(False)
            self.history_text.setVisible(True)
            self.update_history_display()

    def update_stocks_display(self):
        self.stocks_panel.clear()

        if self.stocks_df is None:
            self.stocks_panel.add_label("Файл залишків не завантажено", self.font_normal)
            return

        selected = self.tree.selectedItems()
        if not selected:
            self.stocks_panel.add_label("Виберіть товар для перегляду залишків", self.font_normal)
            return

        item = selected[0]
        values = [item.text(i) for i in range(self.tree.columnCount())]
        name = values[0] if len(values) > 0 else ""
        article = values[5] if len(values) > 5 else ""

        df = self.stocks_df
        stock_items = df[df['Найменування'].astype(str).str.strip() == str(name).strip()]

        if stock_items.empty and article and article.strip():
            same_articles = df[df['Арт'].astype(str).str.strip() == str(article).strip()]
            if len(same_articles) == 1:
                stock_items = same_articles

        if not stock_items.empty:
            stock_item = stock_items.iloc[0]
            self.display_stock_info(stock_item)
        else:
            self.stocks_panel.add_label("Залишки для товару не знайдено", self.font_normal)

    def display_stock_info(self, stock_item: pd.Series):
        # Назва
        title_font = QFont("Calibri", 14)
        title_font.setBold(True)
        lbl_name = QLabel(str(stock_item['Найменування']), self.stocks_panel._content)
        lbl_name.setFont(title_font)
        lbl_name.setStyleSheet("color: #D20000;")
        lbl_name.setAlignment(Qt.AlignHCenter)
        lbl_name.setContentsMargins(0, 10, 0, 10)
        self.stocks_panel._layout.insertWidget(self.stocks_panel._layout.count() - 1, lbl_name)

        # Доступні магазини
        available = []
        for store in self.stores:
            try:
                qty = int(stock_item[store])
            except Exception:
                qty = 0
            if qty > 0:
                available.append((store, qty))

        if not available:
            self.stocks_panel.add_label("Немає в наявності", self.font_normal, center=False)
            return

        # У 3 колонки
        num_columns = 3
        per_col = (len(available) + num_columns - 1) // num_columns
        cols = [available[i*per_col:(i+1)*per_col] for i in range(num_columns)]

        grid = self.stocks_panel.add_grid(rows=max(len(c) for c in cols), cols=num_columns, hgap=12, vgap=6)
        for c_idx, col in enumerate(cols):
            for r_idx, (store, qty) in enumerate(col):
                lbl = QLabel(f"{store}: {qty}")
                font = lbl.font()
                font.setBold(True)
                lbl.setFont(font)
                grid.addWidget(lbl, r_idx, c_idx)

    # ----------------------- ІМПОРТ/АВТОЗАВАНТАЖЕННЯ -----------------------

    def import_excel(self, _show_mapping: bool = False):
        path, _ = QFileDialog.getOpenFileName(self, "Вибрати файл Облік", "", "Excel files (*.xls *.xlsx *.xlsm)")
        if not path:
            return
        try:
            self.load_file(path)
            self.log_action(f"Імпортовано {len(self.df)} товарів")
        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Не вдалося прочитати файл:\n{e}")

    def import_stocks(self):
        path, _ = QFileDialog.getOpenFileName(self, "Вибрати файл Залишки", "", "Excel files (*.xls *.xlsx *.xlsm)")
        if not path:
            return
        try:
            self.load_stocks_file(path)
            self.log_action(f"Імпортовано залишки для {len(self.stores)} магазинів")
            if self.showing_stocks:
                self.update_stocks_display()
        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Не вдалося прочитати файл залишків:\n{e}")

    def auto_load_files(self):
        self.auto_load_accounting_file()
        self.auto_load_stocks_file()
        if self.df is None and self.stocks_df is None:
            # показати порожню історію
            self.update_history_display()

    def auto_load_accounting_file(self):
        home = os.path.expanduser("~")
        candidates = []
        for base in [os.path.join(home, "Desktop"), home, os.path.join(home, "Documents"), os.path.join(home, "Downloads")]:
            candidates.extend(self.find_accounting_files(base))
        candidates = list(dict.fromkeys(candidates))  # унікальні

        if candidates:
            latest = self.get_latest_file(candidates)
            try:
                self.load_file(latest)
                self.log_action(f"Автоматично завантажено: {os.path.basename(latest)}")
            except Exception as e:
                QMessageBox.critical(self, "Помилка", f"Не вдалося завантажити файл:\n{e}")
                self.log_action(f"Помилка завантаження файлу: {os.path.basename(latest)}")
        else:
            self.log_action("Файли Облік не знайдено")

    def find_accounting_files(self, directory: str):
        patterns = [
            "Облік *.xls", "Облік *.xlsx", "Облік *.xlsm",
            "Облік*.*", "Oblik*.xls", "Oblik*.xlsx", "Oblik*.xlsm",
        ]
        files = []
        for p in patterns:
            files.extend(glob.glob(os.path.join(directory, p)))
        # фільтруємо за наявністю дати у назві
        filtered = []
        date_pattern = re.compile(r'Облік[\s_]*(\d{1,2}[.,]\d{1,2}(?:[.,]\d{2,4})?)', re.IGNORECASE)
        for f in files:
            if date_pattern.search(os.path.basename(f)):
                filtered.append(f)
        return filtered if filtered else files

    def auto_load_stocks_file(self):
        home = os.path.expanduser("~")
        candidates = []
        for base in [os.path.join(home, "Desktop"), home, os.path.join(home, "Documents"), os.path.join(home, "Downloads")]:
            candidates.extend(self.find_stocks_files(base))
        candidates = list(dict.fromkeys(candidates))
        if candidates:
            latest = self.get_latest_file(candidates)
            try:
                self.load_stocks_file(latest)
                self.log_action(f"Автоматично завантажено залишки: {os.path.basename(latest)}")
            except Exception as e:
                self.log_action(f"Помилка завантаження файлу залишків: {e}")

    def find_stocks_files(self, directory: str):
        patterns = [
            "*Залишки*.xls", "*Залишки*.xlsx", "*Залишки*.xlsm",
            "*залишки*.*", "*остатки*.*", "Залишки*.xls", "Залишки*.xlsx",
            "*Zalyshky*.xls", "*Zalyshky*.xlsx", "*Ostatki*.xls", "*Ostatki*.xlsx",
        ]
        files = []
        for p in patterns:
            files.extend(glob.glob(os.path.join(directory, p)))
        return files

    def get_latest_file(self, files):
        files_with_dates = [(f, extract_date_from_filename(os.path.basename(f))) for f in files]
        files_sorted = sorted(files_with_dates, key=lambda x: x[1], reverse=True)
        return files_sorted[0][0] if files_sorted else None
        
    def _find_data_start_column(self, df: pd.DataFrame) -> int:
        """
        Знаходить перший стовпець, який не є порожнім,
        аналізуючи перші 10 рядків.
        """
        for col_idx in range(df.shape[1]):
            # Перевіряємо перші 10 рядків (або менше, якщо їх немає)
            # на наявність не-NaN значень
            has_data = not df.iloc[:10, col_idx].isnull().all()
            if has_data:
                return col_idx
        return 0

    def load_file(self, path: str):
        self.df = read_excel_any(path)
        self.current_file = path
        self.file_label.setText(os.path.basename(path))
        
        # Динамічно визначаємо початковий стовпець для даних
        start_col = self._find_data_start_column(self.df)
        
        # Оновлюємо мапінг колонок, зміщуючи всі індекси
        # відповідно до знайденого початкового стовпця
        self.column_mapping = {
            "Найменування": start_col,
            "Закуп": start_col + 1,
            "Приб": start_col + 4,
            "Ціна": start_col + 5,
            "Код": start_col + 7,
            "Арт": start_col + 6,
        }
        
        self.show_all_items()

    def load_stocks_file(self, path: str):
        df = read_excel_any(path)
        self.current_stocks_file = path
        self.stocks_label.setText(os.path.basename(path))

        # Пошук рядка з назвами магазинів (перші 15 рядків)
        stores_row = None
        for i in range(min(15, len(df))):
            row_vals = df.iloc[i].dropna().values
            if len(row_vals) > 3 and any(isinstance(val, str) and (("арсен" in val.lower()) or ("ааа" in val.lower())) for val in row_vals):
                stores_row = i
                break

        if stores_row is None:
            QMessageBox.warning(self, "Увага", "Не вдалося визначити рядок з назвами магазинів")
            return

        data_start_row = stores_row + 2
        data_row = df.iloc[data_start_row]
        
        # Динамічно визначаємо, з якого стовпця починаються дані про товари
        try:
            art_col_idx = data_row.first_valid_index()
            if art_col_idx is None:
                raise ValueError("Не вдалося визначити початковий стовпець даних.")
        except Exception:
            # Fallback, якщо метод first_valid_index не працює
            art_col_idx = 0
            for col_idx in range(len(data_row)):
                if not pd.isna(data_row[col_idx]):
                    art_col_idx = col_idx
                    break

        name_col_idx = art_col_idx + 1
        stores_start_col = name_col_idx + 1

        stores = []
        for col in range(stores_start_col, len(df.columns)):
            store_name = str(df.iloc[stores_row, col]).strip()
            # Додаємо умову для "ААА"
            if col == stores_start_col and (not store_name or store_name.lower() == 'nan'):
                store_name = "ААА"
            elif not store_name:
                continue
            if store_name.lower() == "итог":
                continue
            stores.append(store_name)

        self.stores = stores

        columns_to_take = [art_col_idx, name_col_idx] + list(range(stores_start_col, stores_start_col + len(stores)))
        stocks_df = df.iloc[data_start_row:, columns_to_take].copy()
        stocks_df.columns = ["Арт", "Найменування"] + stores
        stocks_df = stocks_df.dropna(subset=["Арт", "Найменування"], how='all')

        for store in stores:
            stocks_df[store] = pd.to_numeric(stocks_df[store], errors='coerce').fillna(0).astype(int)

        self.stocks_df = stocks_df
        self.log_action(f"Завантажено залишки для {len(stores)} магазинів")

    # ----------------------- ДОПОМОЖНІ -----------------------

    def show_all_items(self):
        if self.df is not None:
            self.show_results(self.df)

    def show_column_mapping_dialog(self):
        dlg = ColumnMappingDialog(self, self.column_mapping.copy())
        if dlg.exec() == QDialog.Accepted:
            try:
                self.column_mapping = dlg.mapping()
                self.show_all_items()
                self.log_action("Оновлено відображення стовпців")
            except Exception as e:
                QMessageBox.critical(self, "Помилка", f"Невірний номер стовпця:\n{e}")

    def paste_into_search(self, text: str):
        if text:
            self.search_edit.setText(text)
            self.search_items(record_history=True)

    def _load_icon(self):
        try:
            icon_names = ['my_icon.ico', 'icon.ico']
            base_paths = []
            if getattr(sys, 'frozen', False):
                base_paths.append(sys._MEIPASS)
            base_paths.append(os.path.dirname(os.path.abspath(__file__)))
            icon_path = None
            for base in base_paths:
                for name in icon_names:
                    candidate = os.path.join(base, name)
                    if os.path.exists(candidate):
                        icon_path = candidate
                        break
                if icon_path:
                    break
            if icon_path:
                self.setWindowIcon(QIcon(icon_path))
        except Exception:
            pass

    def _setup_window_position(self):
        screen = QGuiApplication.primaryScreen().geometry()
        window_width = screen.width() // 2
        window_height = int(screen.height() * 1.0)
        x = screen.width() - window_width
        y = (screen.height() - window_height) // 2
        self.setGeometry(x, y, window_width, window_height)
        self.setMinimumSize(int(window_width * 0.5), int(window_height * 0.5))


# ============================ СКРОЛ-ПАНЕЛЬ ДЛЯ ЗАЛИШКІВ ============================

class ScrollPane(QScrollArea):
    """
    Легка скрол-панель на базі QScrollArea з API add_label/add_grid/clear.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)
        self._content = QWidget(self)
        self._layout = QVBoxLayout(self._content)
        self._layout.setContentsMargins(8, 8, 8, 8)
        self._layout.setSpacing(6)
        self._layout.addStretch(1)
        self.setWidget(self._content)

    def clear(self):
        # Видаляємо всі віджети, окрім stretch (останній)
        while self._layout.count() > 1:
            item = self._layout.takeAt(0)
            w = item.widget()
            if w:
                w.setParent(None)

    def add_label(self, text: str, font: QFont, center=False, pad=(0, 0, 0, 0)):
        lbl = QLabel(text, self._content)
        lbl.setFont(font)
        if center:
            lbl.setAlignment(Qt.AlignHCenter)
        if pad != (0, 0, 0, 0):
            l, t, r, b = pad
            lbl.setContentsMargins(l, t, r, b)
        self._layout.insertWidget(self._layout.count() - 1, lbl)

    def add_grid(self, rows: int, cols: int, hgap=8, vgap=6):
        cont = QWidget(self._content)
        grid = QGridLayout(cont)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(hgap)
        grid.setVerticalSpacing(vgap)
        self._layout.insertWidget(self._layout.count() - 1, cont)
        return grid


# ============================ MAIN ============================

def main():
    app = QApplication(sys.argv)
    # AppUserModelID для іконки в таскбарі Windows
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('Nomenklatura.App')
        except Exception:
            pass
    w = NomenklaturaApp()
    # Встановлюємо праву половину екрана (рамки враховано, щоб правий край не виходив за межі)
    screen = app.primaryScreen()
    rect = screen.availableGeometry()
    target_w = rect.width() // 2
    target_h = int(rect.height() * 0.96)

    # Спочатку показуємо та задаємо розмір клієнтської області
    w.show()
    app.processEvents()
    w.resize(target_w, target_h)
    app.processEvents()

    # Обчислюємо різницю між frameGeometry та geometry (рамки/заголовок)
    frame = w.frameGeometry()
    extra_w = frame.width() - w.geometry().width()
    extra_h = frame.height() - w.geometry().height()

    # Розміщуємо так, щоб рамка також вкладалась у праву половину
    frame_x = rect.x() + rect.width() - (target_w + extra_w)
    frame_y = rect.y() + (rect.height() - (target_h + extra_h)) // 2
    w.move(frame_x, frame_y)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()