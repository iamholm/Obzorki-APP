"""
app_uii.py — УИИ: Генератор обзорных справок + Управление участковыми
Интерфейс: PySide6 (Qt 6)
"""

import os
import sys
import json
import datetime
import threading

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox,
    QLabel, QPushButton, QLineEdit, QFileDialog,
    QSpinBox, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QProgressBar, QFrame, QTextEdit, QScrollArea,
    QAbstractItemView, QMessageBox, QDialog, QAbstractSpinBox,
    QStyleOptionButton, QStyle, QSizePolicy, QInputDialog,
)
from PySide6.QtCore import Qt, Signal, QObject, QRect
from PySide6.QtGui import QFont, QColor, QPalette, QShortcut, QKeySequence

import extract_uii as uii
import db
import parse_dislocation as pdis
import gen_obzorka as g

# ── Настройки приложения ───────────────────────────────────────────────────────

_SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.json')


def _load_settings() -> dict:
    try:
        with open(_SETTINGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_settings(data: dict):
    try:
        with open(_SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


_DEFAULT_CHAR_TEMPLATES = {
    key: list(value) if isinstance(value, (list, tuple)) else [str(value)]
    for key, value in g.CHAR_TEXTS.items()
    if key in g.CHAR_OPTIONS
}


def _clean_template_text(text: str) -> str:
    # Legacy compatibility: remove obsolete {fio} placeholder from saved templates.
    if not isinstance(text, str):
        return ""
    cleaned = text.replace("{fio}", " ")
    return " ".join(cleaned.split()).strip()


def _norm_char_templates(raw) -> dict:
    templates = {k: list(v) for k, v in _DEFAULT_CHAR_TEMPLATES.items()}
    if not isinstance(raw, dict):
        return templates
    for key in g.CHAR_OPTIONS:
        value = raw.get(key)
        items = []
        if isinstance(value, str):
            text = _clean_template_text(value)
            if text:
                items.append(text)
        elif isinstance(value, (list, tuple)):
            for item in value:
                if isinstance(item, str):
                    text = _clean_template_text(item)
                    if text:
                        items.append(text)
        if items:
            templates[key] = items
    return templates


def _get_char_templates(settings: dict = None) -> dict:
    s = settings if isinstance(settings, dict) else _load_settings()
    return _norm_char_templates(s.get('char_templates'))


def _apply_char_templates(templates: dict):
    normalized = _norm_char_templates(templates)
    for key in g.CHAR_OPTIONS:
        g.CHAR_TEXTS[key] = list(normalized[key])


def _default_template_for(char_type: str) -> str:
    defaults = _DEFAULT_CHAR_TEMPLATES.get(char_type) or []
    return defaults[0] if defaults else "По месту жительства характеризуется удовлетворительно."


# ── Константы ──────────────────────────────────────────────────────────────────
APP_VERSION    = "3.0"
DEFAULT_SOURCE = ""
NO_OFFICER     = "— (не выбран)"
BASE_OUT_DIR   = "Обзорки"
ALL_OFFICERS_FILTER = "Все участковые"
CUSTOM_CHAR_OPTION = "индивидуальная"
BULK_OFFICER_PLACEHOLDER = "Назначить участкового..."

RANK_OPTIONS = [
    'мл. лейтенант полиции',
    'лейтенант полиции',
    'ст. лейтенант полиции',
    'капитан полиции',
    'майор полиции',
    'подполковник полиции',
    'полковник полиции',
]

POSITION_OPTIONS = ['УУП', 'Ст. УУП']

# ── Нормализация звания / должности ────────────────────────────────────────────

_RANK_NORM = {
    'мл. лейтенант полиции':     RANK_OPTIONS[0],
    'мл лейтенант полиции':      RANK_OPTIONS[0],
    'лейтенант полиции':         RANK_OPTIONS[1],
    'ст. лейтенант полиции':     RANK_OPTIONS[2],
    'ст лейтенант полиции':      RANK_OPTIONS[2],
    'старший лейтенант полиции': RANK_OPTIONS[2],
    'капитан полиции':           RANK_OPTIONS[3],
    'майор полиции':             RANK_OPTIONS[4],
    'подполковник полиции':      RANK_OPTIONS[5],
    'полковник полиции':         RANK_OPTIONS[6],
}


def _std_rank(rank: str) -> str:
    return _RANK_NORM.get(rank.strip().lower(), rank.strip())


_POS_NORM = {
    'ууп':         POSITION_OPTIONS[0],
    'ст. ууп':     POSITION_OPTIONS[1],
    'старший ууп': POSITION_OPTIONS[1],
}


def _std_pos(pos: str) -> str:
    return _POS_NORM.get(pos.strip().lower(), pos.strip())


CAT_SHORT = {
    'Условное осуждение':                                   'УС',
    'ПРИНУДИТЕЛЬНОЕ лечение':                               'Принуд.',
    'ШТРАФ С ЛЕЧЕНИЕМ (СТ. 72.1 ук РФ)':                  'Штраф',
    'несовершеннолетние':                                   'НС',
    'ЗЗД':                                                  'ЗЗД',
    'ОБЯЗАТЕЛЬНЫЕ РАБОТЫ':                                  'Обяз.р.',
    'Исправительные работы':                                'Испр.р.',
    'Отсрочка до достижения ребенком 14-летнего возраста':  'Отсрочка',
    'Домашний арест ЕЖЕМЕСЯЧНО':                            'Д/арест',
    'ЗОДЕЖЕМЕСЯЧНО':                                        'ЗОДА',
    'Ограничение свободы':                                  'Огр.св.',
    'УДО':                                                  'УДО',
}

def _apply_palette(app: QApplication, dark: bool):
    if not dark:
        app.setPalette(app.style().standardPalette())
        return

    p = QPalette()
    p.setColor(QPalette.ColorRole.Window, QColor("#1f1f1f"))
    p.setColor(QPalette.ColorRole.WindowText, QColor("#eaeaea"))
    p.setColor(QPalette.ColorRole.Base, QColor("#2a2a2a"))
    p.setColor(QPalette.ColorRole.AlternateBase, QColor("#242424"))
    p.setColor(QPalette.ColorRole.ToolTipBase, QColor("#2a2a2a"))
    p.setColor(QPalette.ColorRole.ToolTipText, QColor("#eaeaea"))
    p.setColor(QPalette.ColorRole.Text, QColor("#eaeaea"))
    p.setColor(QPalette.ColorRole.Button, QColor("#2e2e2e"))
    p.setColor(QPalette.ColorRole.ButtonText, QColor("#eaeaea"))
    p.setColor(QPalette.ColorRole.BrightText, QColor("#ffffff"))
    p.setColor(QPalette.ColorRole.Link, QColor("#7ab8ff"))
    p.setColor(QPalette.ColorRole.Highlight, QColor("#3d7fff"))
    p.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    p.setColor(QPalette.ColorRole.PlaceholderText, QColor("#a0a0a0"))
    app.setPalette(p)


# ── Диалог адресов участкового ────────────────────────────────────────────────

class _AddrDialog(QDialog):
    """Всплывающее окно просмотра и редактирования адресов участкового."""

    def __init__(self, officer: dict, parent=None):
        super().__init__(parent)
        fio  = officer.get('fio', '')
        dist = officer.get('district', '')
        self.setWindowTitle(f"Адреса участка {dist}  —  {fio}")
        self.resize(640, 420)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        hint = QLabel(
            "Каждый адрес — с новой строки. "
            "Изменения сохраняются кнопкой «Сохранить».")
        hint.setObjectName("status")
        layout.addWidget(hint)

        self._edit = QTextEdit()
        self._edit.setPlainText(officer.get('addresses', ''))
        layout.addWidget(self._edit)

        btns = QHBoxLayout()
        btns.addStretch()
        b_cancel = QPushButton("Отмена")
        b_cancel.setProperty("flat", True)
        b_cancel.clicked.connect(self.reject)
        btns.addWidget(b_cancel)
        b_save = QPushButton("  Сохранить  ")
        b_save.clicked.connect(self.accept)
        btns.addWidget(b_save)
        layout.addLayout(btns)

    def get_text(self) -> str:
        return self._edit.toPlainText()


# ── Диалог исправления нераспознанных адресов ──────────────────────────────────

class _FixUnmatchedDialog(QDialog):
    """Исправление адресов подучётных, не привязанных к участковому."""

    def __init__(self, rows_data: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Исправление нераспознанных адресов")
        self.resize(980, 540)
        self._rows_data = rows_data or []

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        hint = QLabel(
            "Исправьте адреса и нажмите «ОК и перепроверить».\n"
            "После этого приложение повторно подберёт участковых по обновлённым адресам.")
        hint.setObjectName("status")
        hint.setWordWrap(True)
        root.addWidget(hint)

        self._tbl = QTableWidget(len(self._rows_data), 4)
        self._tbl.setHorizontalHeaderLabels(['№', 'ФИО', 'Дата рожд.', 'Адрес (исправьте)'])
        self._tbl.setShowGrid(True)
        self._tbl.setAlternatingRowColors(True)
        self._tbl.verticalHeader().setVisible(False)
        self._tbl.verticalHeader().setDefaultSectionSize(30)
        self._tbl.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        hdr = self._tbl.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self._tbl.setColumnWidth(0, 36)
        self._tbl.setColumnWidth(2, 100)

        for i, (row_idx, rec) in enumerate(self._rows_data):
            n = QTableWidgetItem(str(row_idx + 1))
            n.setFlags(Qt.ItemFlag.ItemIsEnabled)
            n.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._tbl.setItem(i, 0, n)

            fio = QTableWidgetItem(rec.get('ФИО', ''))
            fio.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self._tbl.setItem(i, 1, fio)

            dob = QTableWidgetItem(rec.get('Дата рождения', ''))
            dob.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self._tbl.setItem(i, 2, dob)

            addr_src = (rec.get('Место жительства', '') or '').replace('\n', ' ').strip()
            addr = QTableWidgetItem(addr_src)
            addr.setToolTip(rec.get('Место жительства', '') or '')
            addr.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsSelectable)
            self._tbl.setItem(i, 3, addr)

        root.addWidget(self._tbl, stretch=1)

        btns = QHBoxLayout()
        btns.addStretch()
        b_cancel = _flat_btn("Отмена")
        b_cancel.clicked.connect(self.reject)
        btns.addWidget(b_cancel)
        b_ok = QPushButton("  ОК и перепроверить  ")
        b_ok.clicked.connect(self.accept)
        btns.addWidget(b_ok)
        root.addLayout(btns)

    def get_addresses(self) -> list:
        result = []
        for i in range(self._tbl.rowCount()):
            item = self._tbl.item(i, 3)
            result.append((item.text().strip() if item else "").strip())
        return result


# ── Диалог настроек ─────────────────────────────────────────────────────────────

class _SettingsDialog(QDialog):
    """Параметры интерфейса и шаблонов характеристик."""

    def __init__(self, dark_theme: bool, char_templates: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.resize(900, 680)

        self._templates = _norm_char_templates(char_templates)
        self._syncing = False
        self._current_type = g.CHAR_OPTIONS[0]
        self._template_row_widgets = []
        self._template_edits = []

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        root.addWidget(tabs, stretch=1)

        # Вкладка шаблонов (первая)
        tab_tpl = QWidget()
        tabs.addTab(tab_tpl, "Шаблоны характеристик")
        p_layout = QVBoxLayout(tab_tpl)
        p_layout.setContentsMargins(8, 8, 8, 8)
        p_layout.setSpacing(8)

        hint = QLabel(
            "Шаблоны выбираются случайно при генерации. "
            "ФИО в тексте шаблона не используется.")
        hint.setObjectName("status")
        hint.setWordWrap(True)
        p_layout.addWidget(hint)

        type_row = QHBoxLayout()
        type_row.setSpacing(8)
        type_row.addWidget(QLabel("Тип характеристики:"))
        self._char_type_cb = QComboBox()
        self._char_type_cb.addItems(g.CHAR_OPTIONS)
        self._char_type_cb.setFixedWidth(220)
        self._char_type_cb.currentTextChanged.connect(self._on_char_type_changed)
        type_row.addWidget(self._char_type_cb)
        type_row.addStretch()
        p_layout.addLayout(type_row)

        self._templates_scroll = QScrollArea()
        self._templates_scroll.setWidgetResizable(True)
        self._templates_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._templates_host = QWidget()
        self._templates_layout = QVBoxLayout(self._templates_host)
        self._templates_layout.setContentsMargins(0, 0, 0, 0)
        self._templates_layout.setSpacing(6)
        self._templates_scroll.setWidget(self._templates_host)
        p_layout.addWidget(self._templates_scroll, stretch=1)

        tpl_btns = QHBoxLayout()
        b_add = _flat_btn("+", 30)
        b_add.setProperty("compact", True)
        b_add.setToolTip("Добавить шаблон")
        b_add.clicked.connect(self._add_template)
        tpl_btns.addWidget(b_add)
        tpl_btns.addStretch()
        p_layout.addLayout(tpl_btns)

        self._render_templates()

        # Вкладка темы
        tab_theme = QWidget()
        tabs.addTab(tab_theme, "Цветовая схема")
        t_layout = QVBoxLayout(tab_theme)
        t_layout.setContentsMargins(8, 8, 8, 8)
        t_layout.setSpacing(8)

        theme_row = QHBoxLayout()
        theme_row.setSpacing(8)
        theme_row.addWidget(QLabel("Тема:"))
        self._theme_cb = QComboBox()
        self._theme_cb.addItems(["Светлая", "Тёмная"])
        self._theme_cb.setFixedWidth(180)
        self._theme_cb.setCurrentIndex(1 if dark_theme else 0)
        theme_row.addWidget(self._theme_cb)
        theme_row.addStretch()
        t_layout.addLayout(theme_row)
        t_layout.addStretch()

        btns = QHBoxLayout()
        btns.addStretch()
        b_reset = _flat_btn("По умолчанию")
        b_reset.clicked.connect(self._reset_templates)
        btns.addWidget(b_reset)
        b_cancel = _flat_btn("Отмена")
        b_cancel.clicked.connect(self.reject)
        btns.addWidget(b_cancel)
        b_save = QPushButton("  Сохранить  ")
        b_save.clicked.connect(self.accept)
        btns.addWidget(b_save)
        root.addLayout(btns)

    def _templates_for_current_type(self) -> list:
        templates = self._templates.get(self._current_type)
        if not templates:
            templates = [_default_template_for(self._current_type)]
            self._templates[self._current_type] = templates
        return templates

    def _clear_templates_ui(self):
        for w in self._template_row_widgets:
            self._templates_layout.removeWidget(w)
            w.deleteLater()
        self._template_row_widgets.clear()
        self._template_edits.clear()

    def _add_template_row_widget(self, index: int, text: str = ""):
        row_widget = QWidget()
        row = QHBoxLayout(row_widget)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)

        num = QLabel(f"{index + 1}.")
        num.setFixedWidth(28)
        num.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        num.setObjectName("status")
        row.addWidget(num)

        edit = QTextEdit()
        edit.setFixedHeight(74)
        edit.setPlainText(text)
        row.addWidget(edit, stretch=1)

        b_del = _flat_btn("-", 30)
        b_del.setProperty("compact", True)
        b_del.setToolTip("Удалить шаблон")
        b_del.clicked.connect(lambda _=False, w=row_widget: self._delete_template_row(w))
        row.addWidget(b_del)

        self._templates_layout.addWidget(row_widget)
        self._template_row_widgets.append(row_widget)
        self._template_edits.append(edit)

    def _commit_current_rows(self):
        texts = []
        for edit in self._template_edits:
            text = _clean_template_text(edit.toPlainText())
            if text:
                texts.append(text)
        if not texts:
            texts = [_default_template_for(self._current_type)]
        self._templates[self._current_type] = texts

    def _render_templates(self):
        templates = self._templates_for_current_type()
        self._clear_templates_ui()
        for i, text in enumerate(templates):
            self._add_template_row_widget(i, text)

    def _on_char_type_changed(self, char_type: str):
        if char_type not in g.CHAR_OPTIONS or self._syncing:
            return
        self._commit_current_rows()
        self._current_type = char_type
        self._render_templates()

    def _add_template(self):
        self._commit_current_rows()
        templates = self._templates_for_current_type()
        templates.append("")
        self._templates[self._current_type] = templates
        self._render_templates()
        if self._template_edits:
            self._template_edits[-1].setFocus()

    def _delete_template_row(self, row_widget: QWidget):
        if row_widget not in self._template_row_widgets:
            return
        idx = self._template_row_widgets.index(row_widget)
        self._templates_layout.removeWidget(row_widget)
        row_widget.deleteLater()
        self._template_row_widgets.pop(idx)
        self._template_edits.pop(idx)
        self._commit_current_rows()
        self._render_templates()

    def _reset_templates(self):
        self._templates = _norm_char_templates({})
        self._current_type = self._char_type_cb.currentText() or g.CHAR_OPTIONS[0]
        self._render_templates()

    def is_dark_theme(self) -> bool:
        return self._theme_cb.currentIndex() == 1

    def get_templates(self) -> dict:
        self._commit_current_rows()
        return _norm_char_templates(self._templates)


# ── Сигналы для потоков ────────────────────────────────────────────────────────

class _Sig(QObject):
    msg      = Signal(str, str)   # text, css-objectName
    progress = Signal(int, int)   # cur, total
    done     = Signal(str, int, int)   # path/dir, created, total
    error    = Signal(str)


class _ObjSig(QObject):
    done  = Signal(object)
    error = Signal(str)


# ── Заголовок таблицы с мастер-чекбоксом ───────────────────────────────────────

class _CheckHeaderView(QHeaderView):
    toggled = Signal(bool)

    def __init__(self, checkbox_column: int, parent=None):
        super().__init__(Qt.Orientation.Horizontal, parent)
        self._checkbox_column = checkbox_column
        self._check_state = Qt.CheckState.Unchecked
        self.setSectionsClickable(True)

    def set_check_state(self, state: Qt.CheckState):
        if self._check_state != state:
            self._check_state = state
            self.viewport().update()

    def paintSection(self, painter, rect, logical_index):
        super().paintSection(painter, rect, logical_index)
        if logical_index != self._checkbox_column:
            return

        opt = QStyleOptionButton()
        opt.state = QStyle.StateFlag.State_Enabled
        if self._check_state == Qt.CheckState.Checked:
            opt.state |= QStyle.StateFlag.State_On
        elif self._check_state == Qt.CheckState.PartiallyChecked:
            opt.state |= QStyle.StateFlag.State_NoChange
        else:
            opt.state |= QStyle.StateFlag.State_Off

        ind = self.style().subElementRect(QStyle.SubElement.SE_CheckBoxIndicator, opt, self)
        x = rect.x() + (rect.width() - ind.width()) // 2
        y = rect.y() + (rect.height() - ind.height()) // 2
        opt.rect = QRect(x, y, ind.width(), ind.height())
        self.style().drawPrimitive(QStyle.PrimitiveElement.PE_IndicatorCheckBox, opt, painter, self)

    def mousePressEvent(self, event):
        idx = self.logicalIndexAt(event.position().toPoint())
        if idx == self._checkbox_column:
            checked = self._check_state != Qt.CheckState.Checked
            self._check_state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
            self.viewport().update()
            self.toggled.emit(checked)
            return
        super().mousePressEvent(event)


# ── Вспомогательные виджет-функции ────────────────────────────────────────────

def _flat_btn(text: str, w: int = None) -> QPushButton:
    b = QPushButton(text)
    b.setAutoDefault(False)
    b.setDefault(False)
    b.setMinimumHeight(28)
    if w:
        b.setFixedWidth(w)
    return b


def _vsep() -> QFrame:
    f = QFrame()
    f.setFrameShape(QFrame.Shape.VLine)
    f.setObjectName("vsep")
    f.setFixedWidth(1)
    f.setFixedHeight(24)
    return f


def _set_status(label: QLabel, text: str, kind: str = "status"):
    label.setObjectName(kind)
    label.setText(text)
    label.style().unpolish(label)
    label.style().polish(label)


# ── Вкладка «Обзорные справки» ────────────────────────────────────────────────

class ObzorkiTab(QWidget):

    C_CHK  = 0
    C_NUM  = 1
    C_FIO  = 2
    C_CAT  = 3
    C_CHAR_POS = 4
    C_CHAR_NEU = 5
    C_CHAR_NEG = 6
    C_CHAR_CUS = 7
    C_OFF  = 8

    def __init__(self, on_open_settings=None):
        super().__init__()
        self._records: list      = []
        self._officers: list     = []
        self._officers_all: list = []
        self._off_labels: list   = [NO_OFFICER]
        self._off_map: dict      = {}
        self._officers_by_id: dict = {}
        self._officer_replacements: dict = {}
        self._off_filter_updating = False
        self._bulk_officer_updating = False
        self._bulk_assigning = False
        self._auto_assigning = False
        self._suspend_officer_change_handler = False
        self._char_syncing = False
        self._edit_lock_all_officers = False
        self._on_open_settings = on_open_settings
        self._load_in_progress   = False
        self._gen_in_progress    = False
        self._pending_load_path  = ""
        self._last_loaded_source = ""
        self._shortcuts: list = []
        self._last_char_click_row = -1

        db.init_db()
        self._build()
        self._restore_ui_settings()
        self.reload_officers()
        self._autoload_last_source()
        self._update_generate_enabled()

    # ── Построение UI ─────────────────────────────────────────────────────────

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # ─ Панель управления списком ───────────────────────────────────────────
        ctrl = QHBoxLayout()
        ctrl.setSpacing(4)

        self._source_path = ""
        self._browse_btn = _flat_btn("Открыть...", 120)
        self._browse_btn.setToolTip("Выбрать исходный файл списка УИИ")
        self._browse_btn.clicked.connect(self._browse_source)
        ctrl.addWidget(self._browse_btn)

        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText("Поиск: ФИО или адрес")
        self._search_edit.setClearButtonEnabled(True)
        self._search_edit.setMinimumWidth(170)
        self._search_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._search_edit.setFixedHeight(28)
        self._search_edit.setToolTip("Быстрый фильтр по ФИО/адресу (Ctrl+F)")
        self._search_edit.textChanged.connect(lambda _t: self._apply_officer_filter())
        ctrl.addWidget(self._search_edit)

        self._off_filter_cb = QComboBox()
        self._off_filter_cb.setMinimumWidth(140)
        self._off_filter_cb.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._off_filter_cb.setFixedHeight(28)
        self._off_filter_cb.setToolTip("Фильтр участкового")
        self._off_filter_cb.currentTextChanged.connect(lambda _t: self._apply_officer_filter())
        ctrl.addWidget(self._off_filter_cb)

        self._bulk_cb = QComboBox()
        self._bulk_cb.addItems(g.CHAR_OPTIONS)
        self._bulk_cb.setCurrentIndex(1)
        self._bulk_cb.setMinimumWidth(125)
        self._bulk_cb.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._bulk_cb.setFixedHeight(28)
        self._bulk_cb.setToolTip("Характеристика для выделенных строк (или видимых, если не выделено)")
        self._bulk_cb.currentTextChanged.connect(lambda _t: self._apply_bulk())
        ctrl.addWidget(self._bulk_cb)

        self._bulk_officer_cb = QComboBox()
        self._bulk_officer_cb.setMinimumWidth(170)
        self._bulk_officer_cb.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._bulk_officer_cb.setFixedHeight(28)
        self._bulk_officer_cb.addItem(BULK_OFFICER_PLACEHOLDER)
        self._bulk_officer_cb.setToolTip("Массово назначить участкового для выделенных строк (или видимых)")
        self._bulk_officer_cb.currentTextChanged.connect(lambda _t: self._apply_bulk_officer())
        ctrl.addWidget(self._bulk_officer_cb)

        self._auto_btn = _flat_btn("↻ Авто-участк.", 120)
        self._auto_btn.setToolTip("Автоматически назначить участковых")
        self._auto_btn.clicked.connect(lambda: self._auto_assign())
        ctrl.addWidget(self._auto_btn)

        ctrl.addStretch()
        if callable(self._on_open_settings):
            b_settings = _flat_btn("Настройки", 95)
            b_settings.clicked.connect(self._on_open_settings)
            ctrl.addWidget(b_settings)
        root.addLayout(ctrl)

        # ─ Таблица подучётных ─────────────────────────────────────────────────
        self._table = QTableWidget(0, 9)
        self._table.setHorizontalHeaderLabels(
            ['', '№', 'ФИО', 'Кат.', '+', '0', '-', 'Инд.', 'Участковый'])
        chk_header = _CheckHeaderView(self.C_CHK, self._table)
        self._table.setHorizontalHeader(chk_header)
        chk_header.toggled.connect(self._set_all_checks)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setShowGrid(True)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        self._table.verticalHeader().setDefaultSectionSize(30)
        self._table.itemChanged.connect(self._on_records_table_item_changed)
        self._table.itemDoubleClicked.connect(self._on_records_table_item_double_clicked)

        hdr = self._table.horizontalHeader()
        for col in (self.C_CHK, self.C_NUM, self.C_FIO,
                    self.C_CAT, self.C_CHAR_POS, self.C_CHAR_NEU, self.C_CHAR_NEG,
                    self.C_CHAR_CUS, self.C_OFF):
            hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.Interactive)
        hdr.setStretchLastSection(False)
        self._table.setColumnWidth(self.C_CHK,  30)
        self._table.setColumnWidth(self.C_NUM,  38)
        self._table.setColumnWidth(self.C_FIO,  200)
        self._table.setColumnWidth(self.C_CAT,  75)
        self._table.setColumnWidth(self.C_CHAR_POS, 48)
        self._table.setColumnWidth(self.C_CHAR_NEU, 48)
        self._table.setColumnWidth(self.C_CHAR_NEG, 48)
        self._table.setColumnWidth(self.C_CHAR_CUS, 58)
        self._table.setColumnWidth(self.C_OFF,  320)

        root.addWidget(self._table, stretch=1)

        # ─ Статус + генерация ─────────────────────────────────────────────────
        bot = QHBoxLayout()
        self._status = QLabel("Нет данных")
        self._status.setObjectName("status")
        bot.addWidget(self._status)

        bot.addSpacing(8)
        self._unmatched_info = QLabel("")
        self._unmatched_info.setStyleSheet("color: #cf222e;")
        bot.addWidget(self._unmatched_info)

        bot.addSpacing(12)
        self._last_source_info = QLabel("Последний загруженный список: —")
        self._last_source_info.setObjectName("status")
        bot.addWidget(self._last_source_info, stretch=1)

        self._prog = QProgressBar()
        self._prog.setTextVisible(False)
        self._prog.setFixedWidth(160)
        self._prog.setVisible(False)
        bot.addWidget(self._prog)

        bot.addSpacing(4)
        right_ctrl = QHBoxLayout()
        right_ctrl.setSpacing(4)
        right_ctrl.setContentsMargins(0, 0, 0, 0)
        right_ctrl.addWidget(QLabel("Квартал:"))
        self._quarter = QSpinBox()
        self._quarter.setRange(1, 4)
        self._quarter.setValue(1)
        self._quarter.setFixedWidth(56)
        self._quarter.setMinimumHeight(30)
        self._quarter.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._quarter.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.UpDownArrows)
        self._quarter.valueChanged.connect(lambda _v: self._save_ui_settings())
        right_ctrl.addWidget(self._quarter)

        right_ctrl.addSpacing(2)
        right_ctrl.addWidget(QLabel("Год:"))
        self._year = QSpinBox()
        self._year.setRange(2020, 2050)
        self._year.setValue(datetime.date.today().year)
        self._year.setFixedWidth(76)
        self._year.setMinimumHeight(30)
        self._year.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._year.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.UpDownArrows)
        self._year.valueChanged.connect(lambda _v: self._save_ui_settings())
        right_ctrl.addWidget(self._year)

        right_ctrl.addSpacing(4)
        self._gen_btn = QPushButton("  Генерировать справки  ")
        self._gen_btn.setFixedHeight(34)
        self._gen_btn.setEnabled(False)
        self._gen_btn.clicked.connect(self._generate)
        right_ctrl.addWidget(self._gen_btn)

        bot.addLayout(right_ctrl)

        root.addLayout(bot)
        self._init_shortcuts()

    def _restore_ui_settings(self):
        settings = _load_settings()
        quarter = settings.get('last_quarter')
        if isinstance(quarter, int) and 1 <= quarter <= 4:
            self._quarter.setValue(quarter)

        year = settings.get('last_year')
        if isinstance(year, int):
            self._year.setValue(max(self._year.minimum(), min(self._year.maximum(), year)))

        last_source = settings.get('last_source_path') or settings.get('last_source')
        if isinstance(last_source, str):
            self._last_loaded_source = last_source.strip()
        self._refresh_last_source_info()

    def _save_ui_settings(self):
        s = _load_settings()
        s['last_quarter'] = self._quarter.value()
        s['last_year'] = self._year.value()
        s['last_source_path'] = self._last_loaded_source
        _save_settings(s)

    def _refresh_last_source_info(self):
        if not hasattr(self, "_last_source_info"):
            return
        if self._last_loaded_source:
            name = os.path.basename(self._last_loaded_source)
            self._last_source_info.setText(f"Последний загруженный список: {name or self._last_loaded_source}")
            self._last_source_info.setToolTip(self._last_loaded_source)
            return
        self._last_source_info.setText("Последний загруженный список: —")
        self._last_source_info.setToolTip("")

    def _autoload_last_source(self):
        path = (self._last_loaded_source or "").strip()
        if not path:
            return
        self._source_path = path
        if hasattr(self, "_browse_btn"):
            self._browse_btn.setToolTip(path)
        if os.path.exists(path):
            self._load_records(path)
            return
        _set_status(self._status, "Последний список не найден. Выберите файл заново.", "info")

    def _set_unmatched_info(self, count: int):
        if not hasattr(self, "_unmatched_info"):
            return
        if count > 0:
            self._unmatched_info.setText(f"Нераспознано адресов: {count}")
            return
        self._unmatched_info.setText("")

    def _refresh_unmatched_info(self):
        self._set_unmatched_info(len(self._get_unmatched_rows()))

    def _refresh_officer_filter_combo(self):
        if not hasattr(self, "_off_filter_cb"):
            return
        self._off_filter_updating = True
        cur = self._off_filter_cb.currentText()
        self._off_filter_cb.clear()
        self._off_filter_cb.addItem(ALL_OFFICERS_FILTER)
        self._off_filter_cb.addItems(self._off_labels)
        self._off_filter_cb.setCurrentText(cur if cur in [ALL_OFFICERS_FILTER] + self._off_labels
                                           else ALL_OFFICERS_FILTER)
        self._off_filter_updating = False
        self._refresh_bulk_officer_combo()

    def _refresh_bulk_officer_combo(self):
        if not hasattr(self, "_bulk_officer_cb"):
            return
        self._bulk_officer_updating = True
        cur = self._bulk_officer_cb.currentText()
        self._bulk_officer_cb.clear()
        self._bulk_officer_cb.addItem(BULK_OFFICER_PLACEHOLDER)
        self._bulk_officer_cb.addItems(self._off_labels)
        self._bulk_officer_cb.setCurrentText(
            cur if cur in [BULK_OFFICER_PLACEHOLDER] + self._off_labels else BULK_OFFICER_PLACEHOLDER
        )
        self._bulk_officer_updating = False

    def _is_edit_locked(self) -> bool:
        return bool(self._edit_lock_all_officers)

    def _set_row_edit_enabled(self, row: int, enabled: bool):
        for col in self._char_cols():
            it = self._table.item(row, col)
            if it is None:
                continue
            flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsUserCheckable
            if enabled:
                flags |= Qt.ItemFlag.ItemIsEnabled
            it.setFlags(flags)
        off_cb = self._table.cellWidget(row, self.C_OFF)
        if isinstance(off_cb, QComboBox):
            off_cb.setEnabled(enabled)

    def _update_edit_mode_lock(self):
        locked = False
        if hasattr(self, "_off_filter_cb"):
            locked = (self._off_filter_cb.currentText() == ALL_OFFICERS_FILTER)
        self._edit_lock_all_officers = locked

        if hasattr(self, "_bulk_cb"):
            self._bulk_cb.setEnabled(not locked)
        if hasattr(self, "_bulk_officer_cb"):
            self._bulk_officer_cb.setEnabled(not locked)
        if hasattr(self, "_auto_btn"):
            can_auto = bool(self._officers) and bool(self._records) and (not self._load_in_progress) and (not self._gen_in_progress)
            self._auto_btn.setEnabled(can_auto)

        for row in range(self._table.rowCount()):
            self._set_row_edit_enabled(row, not locked)

    def _is_text_input_focused(self) -> bool:
        fw = QApplication.focusWidget()
        return isinstance(fw, (QLineEdit, QTextEdit, QAbstractSpinBox, QComboBox))

    def _init_shortcuts(self):
        self._shortcuts.clear()

        def add_shortcut(seq: str, handler):
            sc = QShortcut(QKeySequence(seq), self)
            sc.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
            sc.activated.connect(handler)
            self._shortcuts.append(sc)

        add_shortcut("Ctrl+F", self._focus_search)
        add_shortcut("Ctrl+A", lambda: self._shortcut_set_all_checks(True))
        add_shortcut("Ctrl+Shift+A", lambda: self._shortcut_set_all_checks(False))
        add_shortcut("1", lambda: self._shortcut_apply_char(0))
        add_shortcut("2", lambda: self._shortcut_apply_char(1))
        add_shortcut("3", lambda: self._shortcut_apply_char(2))
        add_shortcut("4", self._shortcut_apply_bulk_custom)

    def _focus_search(self):
        if not hasattr(self, "_search_edit"):
            return
        self._search_edit.setFocus()
        self._search_edit.selectAll()

    def _shortcut_set_all_checks(self, checked: bool):
        if self._is_text_input_focused():
            return
        self._set_all_checks(checked)

    def _shortcut_apply_char(self, index: int):
        if self._is_text_input_focused() or self._is_edit_locked():
            return
        if index < 0 or index >= len(g.CHAR_OPTIONS):
            return
        val = g.CHAR_OPTIONS[index]
        if self._bulk_cb.currentText() != val:
            self._bulk_cb.setCurrentText(val)
        else:
            self._apply_bulk()

    def _shortcut_apply_bulk_custom(self):
        if self._is_text_input_focused() or self._is_edit_locked():
            return
        text, ok = QInputDialog.getMultiLineText(
            self,
            "Индивидуальная характеристика",
            "Текст для всех видимых строк:",
            "",
        )
        text = (text or "").strip()
        if not ok or not text:
            return
        self._apply_bulk_custom_text(text)

    def _apply_bulk_custom_text(self, custom_text: str):
        if self._is_edit_locked():
            return
        custom_text = (custom_text or "").strip()
        if not custom_text:
            return
        changed = 0
        for row in self._target_rows_for_mass_action():
            cur_type, cur_custom = self._char_type_for_row(row)
            if cur_type == CUSTOM_CHAR_OPTION and cur_custom == custom_text:
                continue
            self._set_row_char_type(row, CUSTOM_CHAR_OPTION, custom_text, persist=True)
            changed += 1
        if changed > 0:
            _set_status(self._status, f"Индивидуальная характеристика применена: {changed}", "ok")

    def _apply_officer_filter(self):
        if not hasattr(self, "_off_filter_cb") or self._off_filter_updating:
            return
        selected = self._off_filter_cb.currentText()
        show_all = (selected == ALL_OFFICERS_FILTER)
        query = ""
        if hasattr(self, "_search_edit"):
            query = (self._search_edit.text() or "").strip().lower()
        visible = 0
        for row in range(self._table.rowCount()):
            off_cb = self._table.cellWidget(row, self.C_OFF)
            label = off_cb.currentText() if isinstance(off_cb, QComboBox) else NO_OFFICER
            match_text = True
            if query and row < len(self._records):
                rec = self._records[row]
                haystack = f"{rec.get('ФИО', '')} {rec.get('Место жительства', '')}".lower()
                match_text = query in haystack
            hidden = ((not show_all and label != selected) or (not match_text))
            self._table.setRowHidden(row, hidden)
            if not hidden:
                visible += 1
        if self._records:
            _set_status(self._status, f"Показано записей: {visible}")
        self._update_edit_mode_lock()
        self._update_master_check_header()
        self._update_generate_enabled()

    def _set_all_checks(self, checked: bool):
        self._table.blockSignals(True)
        for row in range(self._table.rowCount()):
            if self._table.isRowHidden(row):
                continue
            item = self._table.item(row, self.C_CHK)
            if item is not None:
                item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self._table.blockSignals(False)
        self._update_master_check_header()
        self._update_generate_enabled()

    def _update_master_check_header(self):
        hdr = self._table.horizontalHeader()
        if not isinstance(hdr, _CheckHeaderView):
            return
        visible = 0
        checked = 0
        for row in range(self._table.rowCount()):
            if self._table.isRowHidden(row):
                continue
            cell = self._table.item(row, self.C_CHK)
            if cell is None:
                continue
            visible += 1
            if cell.checkState() == Qt.CheckState.Checked:
                checked += 1
        if visible <= 0:
            state = Qt.CheckState.Unchecked
        elif checked == 0:
            state = Qt.CheckState.Unchecked
        elif checked == visible:
            state = Qt.CheckState.Checked
        else:
            state = Qt.CheckState.PartiallyChecked
        hdr.set_check_state(state)

    def _on_records_table_item_changed(self, item: QTableWidgetItem):
        if not item:
            return
        col = item.column()
        if col == self.C_CHK:
            self._update_master_check_header()
            self._update_generate_enabled()
            return
        if col in self._char_cols():
            self._on_row_char_checkbox_changed(item)

    def _on_records_table_item_double_clicked(self, item: QTableWidgetItem):
        if not item:
            return
        if self._is_edit_locked():
            return
        if item.column() != self.C_CHAR_CUS:
            return
        row = item.row()
        current_type, current_text = self._char_type_for_row(row)
        seed_text = current_text if current_type == CUSTOM_CHAR_OPTION else ""
        new_text = self._edit_custom_characteristic(row, seed_text)
        if not new_text:
            return
        self._set_row_char_type(row, CUSTOM_CHAR_OPTION, new_text, persist=True)
        _set_status(self._status, "Индивидуальная характеристика сохранена", "ok")

    def _on_row_char_checkbox_changed(self, item: QTableWidgetItem):
        if self._char_syncing or self._is_edit_locked():
            return
        row = item.row()
        col = item.column()
        if row < 0 or col not in self._char_cols():
            return

        checked = (item.checkState() == Qt.CheckState.Checked)
        if not checked:
            # Запрещаем состояние «не выбрано ничего».
            if all(
                (self._table.item(row, c) is None or self._table.item(row, c).checkState() != Qt.CheckState.Checked)
                for c in self._char_cols()
            ):
                self._char_syncing = True
                item.setCheckState(Qt.CheckState.Checked)
                self._char_syncing = False
            return

        target_rows = [row]
        modifiers = QApplication.keyboardModifiers()
        if modifiers & Qt.KeyboardModifier.ShiftModifier and self._last_char_click_row >= 0:
            lo = min(self._last_char_click_row, row)
            hi = max(self._last_char_click_row, row)
            target_rows = [r for r in range(lo, hi + 1) if not self._table.isRowHidden(r)]
        else:
            selected_rows = self._selected_visible_rows()
            if len(selected_rows) > 1 and row in selected_rows:
                target_rows = selected_rows
        self._last_char_click_row = row

        char_type = self._char_type_for_col(col)
        if char_type == CUSTOM_CHAR_OPTION:
            fallback_type = "нейтральная"
            for c, t in (
                (self.C_CHAR_POS, "положительная"),
                (self.C_CHAR_NEU, "нейтральная"),
                (self.C_CHAR_NEG, "отрицательная"),
            ):
                it = self._table.item(row, c)
                if it is not None and it.checkState() == Qt.CheckState.Checked:
                    fallback_type = t
                    break
            seed = self._custom_text_for_row(row)
            text = self._edit_custom_characteristic(row, seed)
            if not text:
                self._set_row_char_type(row, fallback_type, "", persist=True)
                return
            for r in target_rows:
                self._set_row_char_type(r, CUSTOM_CHAR_OPTION, text, persist=True)
            _set_status(self._status, f"Индивидуальная характеристика применена: {len(target_rows)}", "ok")
            return

        custom_rows = [r for r in target_rows if self._char_type_for_row(r)[0] == CUSTOM_CHAR_OPTION]
        replace_custom = True
        if custom_rows:
            r = QMessageBox.question(
                self,
                "Индивидуальные характеристики",
                (
                    f"В выбранных строках есть индивидуальные: {len(custom_rows)}.\n"
                    f"Заменить их на «{char_type}»?"
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            replace_custom = (r == QMessageBox.StandardButton.Yes)

        changed = 0
        skipped = 0
        for r in target_rows:
            cur_type, _cur_custom = self._char_type_for_row(r)
            if cur_type == CUSTOM_CHAR_OPTION and not replace_custom:
                skipped += 1
                continue
            self._set_row_char_type(r, char_type, "", persist=True)
            changed += 1
        if changed or skipped:
            msg = f"Применено «{char_type}»: {changed}"
            if skipped:
                msg += f" · индивидуальных оставлено: {skipped}"
            _set_status(self._status, msg, "ok")

    def _update_generate_enabled(self):
        if self._load_in_progress or self._gen_in_progress or not self._records:
            self._gen_btn.setEnabled(False)
            return
        has_checked = False
        for row in range(self._table.rowCount()):
            if self._table.isRowHidden(row):
                continue
            item = self._table.item(row, self.C_CHK)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                has_checked = True
                break
        self._gen_btn.setEnabled(has_checked)

    # ── Публичные методы (вызываются из других вкладок) ───────────────────────

    def reload_officers(self):
        """Перечитывает список участковых из БД и обновляет комбобоксы."""
        self._officers_all = db.all_officers()
        self._officers = list(self._officers_all)
        self._officers_by_id = {o['id']: o for o in self._officers}
        raw_replacements = db.all_officer_replacements()
        self._officer_replacements = {
            int(off_id): int(rep_id)
            for off_id, rep_id in raw_replacements.items()
            if off_id in self._officers_by_id and rep_id in self._officers_by_id and off_id != rep_id
        }
        self._off_labels = [NO_OFFICER] + [g._officer_label(o) for o in self._officers]
        self._off_map    = {g._officer_label(o): o for o in self._officers}
        self._refresh_officer_filter_combo()
        self._refresh_officer_combos()

    def _resolve_generation_officer(self, officer: dict):
        """Возвращает фактического подписанта с учётом цепочки замещений."""
        if not isinstance(officer, dict):
            return None
        cur_id = officer.get('id')
        if cur_id is None:
            return officer
        visited = set()
        resolved = officer
        while True:
            if cur_id in visited:
                break
            visited.add(cur_id)
            repl_id = self._officer_replacements.get(cur_id)
            if repl_id is None:
                break
            repl = self._officers_by_id.get(repl_id)
            if repl is None:
                break
            resolved = repl
            cur_id = repl.get('id')
        return resolved

    def get_unmatched_records(self) -> list:
        """Возвращает [(rec, row_idx)] строк без назначенного участкового."""
        result = []
        for row, rec in enumerate(self._records):
            cb = self._table.cellWidget(row, self.C_OFF)
            if isinstance(cb, QComboBox) and cb.currentText() == NO_OFFICER:
                result.append((rec, row))
        return result

    def _get_unmatched_rows(self) -> list:
        rows = []
        for row in range(self._table.rowCount()):
            cb = self._table.cellWidget(row, self.C_OFF)
            if isinstance(cb, QComboBox) and cb.currentText() == NO_OFFICER:
                rows.append(row)
        return rows

    def set_assignment_for_row(self, row: int, label: str):
        """Устанавливает участкового в строке таблицы."""
        cb = self._table.cellWidget(row, self.C_OFF)
        if isinstance(cb, QComboBox) and label in self._off_labels:
            prev = self._suspend_officer_change_handler
            self._suspend_officer_change_handler = True
            cb.setCurrentText(label)
            cb.setProperty('prev_off_label', cb.currentText())
            self._suspend_officer_change_handler = prev

    # ── Загрузка ──────────────────────────────────────────────────────────────

    def _browse_source(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите список УИИ", "",
            "Word документ (*.docx);;Все файлы (*.*)")
        if path:
            self._source_path = path
            self._browse_btn.setToolTip(path)
            self._load_records(path)

    def _load_records(self, path: str):
        path = (path or '').strip()
        if not path or not os.path.exists(path):
            _set_status(self._status, f"Файл не найден: {path}", "error")
            return
        if self._load_in_progress:
            return

        self._save_ui_settings()
        self._pending_load_path = path
        self._load_in_progress = True
        self._browse_btn.setEnabled(False)
        self._gen_btn.setEnabled(False)
        _set_status(self._status, "Загрузка…", "info")

        sig = _ObjSig(self)
        sig.done.connect(self._on_records_loaded_ok)
        sig.error.connect(self._on_records_loaded_error)

        def task():
            try:
                records = uii.parse_document(path)
                sig.done.emit(records)
            except Exception as exc:
                sig.error.emit(str(exc))

        threading.Thread(target=task, daemon=True).start()

    def _current_record_keys(self) -> set:
        keys = set()
        for rec in self._records:
            fio = str(rec.get('ФИО', '')).strip()
            dob = str(rec.get('Дата рождения', '')).strip()
            if fio:
                keys.add((fio, dob))
        return keys

    def _apply_saved_address_fixes(self):
        fixes = db.all_person_address_fixes()
        if not fixes or not self._records:
            return
        for rec in self._records:
            fio = str(rec.get('ФИО', '')).strip()
            dob = str(rec.get('Дата рождения', '')).strip()
            fixed = fixes.get((fio, dob))
            if fixed:
                rec['Место жительства'] = fixed

    def _cleanup_missing_saved_characteristics(self):
        current_keys = self._current_record_keys()
        missing = db.list_missing_person_characteristics(current_keys)
        if not missing:
            return
        preview_lines = []
        for item in missing[:12]:
            fio = item.get('fio', '') or '—'
            dob = item.get('dob', '')
            preview_lines.append(f"- {fio}" + (f" ({dob})" if dob else ""))
        more = "" if len(missing) <= 12 else f"\n... и ещё {len(missing) - 12}"
        msg = (
            f"В сохранённых индивидуальных/ручных характеристиках есть {len(missing)} записей,\n"
            f"которых нет в текущем списке.\n\n"
            f"{chr(10).join(preview_lines)}{more}\n\n"
            f"Удалить эти сохранённые записи?"
        )
        r = QMessageBox.question(
            self,
            "Очистка сохранённых характеристик",
            msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if r == QMessageBox.StandardButton.Yes:
            db.delete_person_characteristics(
                [(m.get('fio', ''), m.get('dob', '')) for m in missing]
            )

    def _on_records_loaded_ok(self, records):
        self._load_in_progress = False
        self._browse_btn.setEnabled(True)
        if self._pending_load_path:
            self._source_path = self._pending_load_path
            self._last_loaded_source = self._pending_load_path
            self._refresh_last_source_info()
            self._save_ui_settings()
        self._pending_load_path = ""
        self._records = records or []
        self._apply_saved_address_fixes()
        self._cleanup_missing_saved_characteristics()
        self._populate_table()
        _set_status(self._status, "Список загружен", "ok")
        self._resolve_unmatched_addresses()
        self._refresh_unmatched_info()
        self._update_generate_enabled()

    def _on_records_loaded_error(self, message: str):
        self._load_in_progress = False
        self._browse_btn.setEnabled(True)
        self._pending_load_path = ""
        QMessageBox.critical(self, "Ошибка загрузки", message)
        _set_status(self._status, "Ошибка загрузки", "error")
        self._update_generate_enabled()

    def _populate_table(self):
        assignments    = db.all_assignments()   # {(fio, dob): officer_id}
        char_prefs     = db.all_person_characteristics()
        officers_by_id = {o['id']: o for o in self._officers}

        self._char_syncing = True
        self._table.blockSignals(True)
        self._table.setRowCount(0)
        self._table.setRowCount(len(self._records))
        self._last_char_click_row = -1

        for i, rec in enumerate(self._records):
            fio_val = rec.get('ФИО', '')
            dob_val = rec.get('Дата рождения', '')

            chk = QTableWidgetItem()
            chk.setCheckState(Qt.CheckState.Checked)
            chk.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            chk.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(i, self.C_CHK, chk)

            num = QTableWidgetItem(str(i + 1))
            num.setFlags(Qt.ItemFlag.ItemIsEnabled)
            num.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(i, self.C_NUM, num)

            fio_item = QTableWidgetItem(fio_val or '—')
            fio_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self._table.setItem(i, self.C_FIO, fio_item)

            cat_full  = rec.get('Категория', '')
            cat_short = CAT_SHORT.get(cat_full, cat_full[:8])
            cat_item  = QTableWidgetItem(cat_short)
            cat_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            cat_item.setToolTip(cat_full)
            cat_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(i, self.C_CAT, cat_item)

            pref = char_prefs.get((fio_val, dob_val))
            pref_type = (pref.get('char_type', '') if isinstance(pref, dict) else '').strip()
            pref_custom = (pref.get('custom_text', '') if isinstance(pref, dict) else '').strip()
            if pref_type not in ("положительная", "нейтральная", "отрицательная", CUSTOM_CHAR_OPTION):
                pref_type = CUSTOM_CHAR_OPTION if pref_custom else 'нейтральная'
            if not pref_type:
                pref_type = 'нейтральная'
            if pref_type == CUSTOM_CHAR_OPTION and not pref_custom:
                pref_type = "нейтральная"

            for col in self._char_cols():
                self._table.setItem(i, col, self._new_char_item())
            self._set_row_char_type(i, pref_type, pref_custom, persist=False)

            off_cb = QComboBox()
            off_cb.addItems(self._off_labels)
            off_cb.setMinimumWidth(290)
            off_cb.setFixedHeight(24)

            # Ручное назначение имеет приоритет над авто
            assigned_id = assignments.get((fio_val, dob_val))
            if assigned_id is not None and assigned_id in officers_by_id:
                label = g._officer_label(officers_by_id[assigned_id])
                if label in self._off_labels:
                    off_cb.setCurrentText(label)
            else:
                matched = g._match_officer(rec.get('Место жительства', ''), self._officers)
                if matched:
                    label = g._officer_label(matched)
                    if label in self._off_labels:
                        off_cb.setCurrentText(label)

            off_cb.setProperty('prev_off_label', off_cb.currentText())
            self._table.setCellWidget(i, self.C_OFF, off_cb)
            off_cb.currentTextChanged.connect(lambda _t, _r=i: self._on_row_officer_changed(_r))

        self._table.blockSignals(False)
        self._char_syncing = False
        self._update_master_check_header()
        self._refresh_officer_filter_combo()
        self._apply_officer_filter()
        self._refresh_unmatched_info()
        self._update_generate_enabled()

    def _refresh_officer_combos(self):
        prev = self._suspend_officer_change_handler
        self._suspend_officer_change_handler = True
        self._table.blockSignals(True)
        for row in range(self._table.rowCount()):
            cb = self._table.cellWidget(row, self.C_OFF)
            if isinstance(cb, QComboBox):
                cur = cb.currentText()
                cb.clear()
                cb.addItems(self._off_labels)
                cb.setCurrentText(cur if cur in self._off_labels else NO_OFFICER)
                cb.setProperty('prev_off_label', cb.currentText())
        self._table.blockSignals(False)
        self._suspend_officer_change_handler = prev
        self._apply_officer_filter()
        self._update_generate_enabled()

    # ── Массовые действия ─────────────────────────────────────────────────────

    def _save_row_assignment(self, row: int):
        if row < 0 or row >= self._table.rowCount():
            return
        fio, dob = self._person_key_for_row(row)
        if not fio:
            return
        cb = self._table.cellWidget(row, self.C_OFF)
        if not isinstance(cb, QComboBox):
            return
        label = (cb.currentText() or '').strip()
        officer = self._off_map.get(label)
        db.set_assignment(fio, dob, officer['id'] if officer else None)

    def _on_row_officer_changed(self, row: int):
        cb = self._table.cellWidget(row, self.C_OFF)
        if not isinstance(cb, QComboBox):
            return
        cur_label = (cb.currentText() or '').strip()
        prev_label = str(cb.property('prev_off_label') or cur_label).strip() or NO_OFFICER

        if self._suspend_officer_change_handler:
            cb.setProperty('prev_off_label', cur_label)
            return

        if (not self._bulk_assigning and not self._auto_assigning and not self._is_edit_locked()
                and prev_label != cur_label):
            fio, _dob = self._person_key_for_row(row)
            r = QMessageBox.question(
                self,
                "Подтверждение смены участкового",
                (
                    f"Изменить участкового для:\n{fio or '—'}\n\n"
                    f"Было: {prev_label}\n"
                    f"Станет: {cur_label}"
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                prev = self._suspend_officer_change_handler
                self._suspend_officer_change_handler = True
                cb.setCurrentText(prev_label if prev_label in self._off_labels else NO_OFFICER)
                self._suspend_officer_change_handler = prev
                return

        cb.setProperty('prev_off_label', cur_label)
        self._save_row_assignment(row)
        if self._bulk_assigning or self._auto_assigning:
            return
        self._apply_officer_filter()
        self._refresh_unmatched_info()

    def _person_key_for_row(self, row: int):
        if row < 0 or row >= len(self._records):
            return "", ""
        rec = self._records[row]
        return (str(rec.get('ФИО', '')).strip(), str(rec.get('Дата рождения', '')).strip())

    def _char_cols(self) -> tuple:
        return (self.C_CHAR_POS, self.C_CHAR_NEU, self.C_CHAR_NEG, self.C_CHAR_CUS)

    def _char_type_for_col(self, col: int) -> str:
        return {
            self.C_CHAR_POS: "положительная",
            self.C_CHAR_NEU: "нейтральная",
            self.C_CHAR_NEG: "отрицательная",
            self.C_CHAR_CUS: CUSTOM_CHAR_OPTION,
        }.get(col, "нейтральная")

    def _char_col_for_type(self, char_type: str) -> int:
        return {
            "положительная": self.C_CHAR_POS,
            "нейтральная": self.C_CHAR_NEU,
            "отрицательная": self.C_CHAR_NEG,
            CUSTOM_CHAR_OPTION: self.C_CHAR_CUS,
        }.get(char_type, self.C_CHAR_NEU)

    def _new_char_item(self) -> QTableWidgetItem:
        it = QTableWidgetItem()
        it.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
        it.setCheckState(Qt.CheckState.Unchecked)
        it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        return it

    def _custom_text_for_row(self, row: int) -> str:
        it = self._table.item(row, self.C_CHAR_CUS)
        if it is None:
            return ""
        return str(it.data(Qt.ItemDataRole.UserRole) or "").strip()

    def _set_custom_text_for_row(self, row: int, text: str):
        it = self._table.item(row, self.C_CHAR_CUS)
        if it is None:
            return
        value = (text or "").strip()
        it.setData(Qt.ItemDataRole.UserRole, value)
        it.setToolTip(value if value else "")

    def _char_type_for_row(self, row: int):
        selected_col = None
        for col in self._char_cols():
            it = self._table.item(row, col)
            if it is not None and it.checkState() == Qt.CheckState.Checked:
                selected_col = col
                break
        if selected_col is None:
            return "нейтральная", ""
        char_type = self._char_type_for_col(selected_col)
        if char_type == CUSTOM_CHAR_OPTION:
            return char_type, self._custom_text_for_row(row)
        return char_type, ""

    def _set_row_char_type(self, row: int, char_type: str, custom_text: str = "", persist: bool = True):
        if row < 0 or row >= self._table.rowCount():
            return
        if char_type not in ("положительная", "нейтральная", "отрицательная", CUSTOM_CHAR_OPTION):
            char_type = "нейтральная"

        if char_type == CUSTOM_CHAR_OPTION:
            custom_text = (custom_text or "").strip()
            if not custom_text:
                custom_text = self._custom_text_for_row(row)
            if not custom_text:
                char_type = "нейтральная"

        target_col = self._char_col_for_type(char_type)
        self._char_syncing = True
        for col in self._char_cols():
            it = self._table.item(row, col)
            if it is None:
                continue
            it.setCheckState(Qt.CheckState.Checked if col == target_col else Qt.CheckState.Unchecked)
        self._set_custom_text_for_row(row, custom_text if char_type == CUSTOM_CHAR_OPTION else "")
        self._char_syncing = False

        if not persist:
            return
        fio, dob = self._person_key_for_row(row)
        if not fio:
            return
        if char_type == CUSTOM_CHAR_OPTION:
            db.set_person_characteristic(fio, dob, CUSTOM_CHAR_OPTION, custom_text)
            return
        db.set_person_characteristic(fio, dob, char_type, "")

    def _selected_visible_rows(self) -> list:
        rows = sorted({idx.row() for idx in self._table.selectionModel().selectedRows()})
        return [r for r in rows if not self._table.isRowHidden(r)]

    def _target_rows_for_mass_action(self) -> list:
        selected_rows = self._selected_visible_rows()
        if selected_rows:
            return selected_rows
        return [r for r in range(self._table.rowCount()) if not self._table.isRowHidden(r)]

    def _edit_custom_characteristic(self, row: int, current_text: str = ""):
        fio, _dob = self._person_key_for_row(row)
        prompt = (
            f"Введите индивидуальный текст характеристики для:\n{fio or '—'}\n\n"
            "Пустой текст не допускается."
        )
        text, ok = QInputDialog.getMultiLineText(
            self, "Индивидуальная характеристика", prompt, current_text or ""
        )
        if not ok:
            return None
        text = (text or "").strip()
        if not text:
            QMessageBox.warning(self, "Пустой текст", "Индивидуальная характеристика не может быть пустой.")
            return None
        return text

    def _apply_bulk(self):
        if self._is_edit_locked():
            return
        text = (self._bulk_cb.currentText() or "").strip()
        if text not in g.CHAR_OPTIONS:
            return
        target_rows = self._target_rows_for_mass_action()
        custom_rows = []
        for row in target_rows:
            if self._char_type_for_row(row)[0] == CUSTOM_CHAR_OPTION:
                custom_rows.append(row)

        replace_custom = False
        if custom_rows:
            r = QMessageBox.question(
                self,
                "Индивидуальные характеристики",
                (
                    f"В видимых строках найдено {len(custom_rows)} индивидуальных характеристик.\n"
                    f"Заменить их на «{text}»?\n\n"
                    "Да — заменить.\n"
                    "Нет — оставить индивидуальные без изменений."
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            replace_custom = (r == QMessageBox.StandardButton.Yes)

        changed = 0
        skipped_custom = 0
        for row in target_rows:
            cur, _custom = self._char_type_for_row(row)
            if cur == CUSTOM_CHAR_OPTION and not replace_custom:
                skipped_custom += 1
                continue
            if cur != text:
                self._set_row_char_type(row, text, "", persist=True)
                changed += 1
        if changed or skipped_custom:
            msg = f"Применено «{text}»: {changed}"
            if skipped_custom:
                msg += f" · индивидуальных оставлено: {skipped_custom}"
            _set_status(self._status, msg, "ok")

    def _apply_bulk_officer(self):
        if self._is_edit_locked():
            return
        if self._bulk_officer_updating or not hasattr(self, "_bulk_officer_cb"):
            return
        label = (self._bulk_officer_cb.currentText() or "").strip()
        if not label or label == BULK_OFFICER_PLACEHOLDER:
            return
        changed = 0
        self._bulk_assigning = True
        try:
            for row in self._target_rows_for_mass_action():
                cb = self._table.cellWidget(row, self.C_OFF)
                if not isinstance(cb, QComboBox):
                    continue
                if (cb.currentText() or "").strip() == label:
                    continue
                cb.setCurrentText(label)
                changed += 1
        finally:
            self._bulk_assigning = False
        self._apply_officer_filter()
        self._refresh_unmatched_info()
        if changed > 0:
            if label == NO_OFFICER:
                _set_status(self._status, f"Снято назначение участкового: {changed}", "ok")
            else:
                _set_status(self._status, f"Массово назначено «{label}»: {changed}", "ok")
        self._bulk_officer_updating = True
        self._bulk_officer_cb.setCurrentText(BULK_OFFICER_PLACEHOLDER)
        self._bulk_officer_updating = False

    def _auto_assign(self, rows: list = None, update_status: bool = True, confirm_changes: bool = True):
        if not self._officers:
            if update_status:
                _set_status(self._status,
                            "Сначала загрузите дислокацию на вкладке «Участковые».", "error")
            return 0
        target_rows = list(rows) if rows is not None else list(range(len(self._records)))
        changes = []
        for row in target_rows:
            if row < 0 or row >= len(self._records):
                continue
            rec = self._records[row]
            cb = self._table.cellWidget(row, self.C_OFF)
            if not isinstance(cb, QComboBox):
                continue
            current_label = (cb.currentText() or "").strip() or NO_OFFICER
            matched = g._match_officer(rec.get('Место жительства', ''), self._officers)
            new_label = g._officer_label(matched) if matched else NO_OFFICER
            if new_label not in self._off_labels:
                new_label = NO_OFFICER
            if new_label != current_label:
                changes.append((row, current_label, new_label))

        if not changes:
            if update_status:
                _set_status(self._status, "Авто-назначение: изменений не найдено", "info")
            return 0

        if confirm_changes:
            preview = []
            for row, old_lbl, new_lbl in changes[:8]:
                fio = str(self._records[row].get('ФИО', '') or '—')
                preview.append(f"• {fio}: {old_lbl} → {new_lbl}")
            more = "" if len(changes) <= 8 else f"\n... и ещё {len(changes) - 8}"
            msg = (
                f"Авто-назначение изменит участкового у {len(changes)} записей.\n\n"
                f"{chr(10).join(preview)}{more}\n\n"
                "Продолжить?"
            )
            r = QMessageBox.question(
                self,
                "Подтверждение авто-назначения",
                msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                if update_status:
                    _set_status(self._status, "Авто-назначение отменено", "info")
                return 0

        count = 0
        self._auto_assigning = True
        try:
            for row, _old_lbl, new_label in changes:
                cb = self._table.cellWidget(row, self.C_OFF)
                if not isinstance(cb, QComboBox):
                    continue
                cb.setCurrentText(new_label)
                self._save_row_assignment(row)
                count += 1
        finally:
            self._auto_assigning = False
        self._apply_officer_filter()
        self._refresh_unmatched_info()
        if update_status:
            total = len(target_rows)
            _set_status(self._status, f"Авто-назначено: {count} из {total}", "ok")
        return count

    def _resolve_unmatched_addresses(self):
        if not self._records or not self._officers:
            return

        while True:
            rows = self._get_unmatched_rows()
            if not rows:
                return

            rows_data = [(row, self._records[row]) for row in rows]
            dlg = _FixUnmatchedDialog(rows_data, self)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                _set_status(self._status, "", "status")
                self._set_unmatched_info(len(rows))
                return

            new_addresses = dlg.get_addresses()
            for idx, row in enumerate(rows):
                if idx < len(new_addresses):
                    addr = (new_addresses[idx] or "").strip()
                    if addr:
                        self._records[row]['Место жительства'] = addr
                        fio, dob = self._person_key_for_row(row)
                        db.set_person_address_fix(fio, dob, addr)

            self._auto_assign(rows=rows, update_status=False, confirm_changes=False)
            remaining = len(self._get_unmatched_rows())
            if remaining <= 0:
                _set_status(self._status, "Все адреса распознаны.", "ok")
                self._set_unmatched_info(0)
                return

            again = QMessageBox.question(
                self, "Остались нераспознанные адреса",
                f"После перепроверки осталось записей без участкового: {remaining}.\n"
                "Открыть окно исправления снова?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if again != QMessageBox.StandardButton.Yes:
                _set_status(self._status, "", "status")
                self._set_unmatched_info(remaining)
                return

    # ── Генерация ─────────────────────────────────────────────────────────────

    def _generate(self):
        if not os.path.exists(g.TEMPLATE_FILE):
            QMessageBox.critical(
                self, "Шаблон не найден",
                f"Файл шаблона «{g.TEMPLATE_FILE}» не найден рядом со скриптом.")
            return

        self._save_ui_settings()
        quarter = self._quarter.value()
        year    = self._year.value()

        selected = []
        for row, rec in enumerate(self._records):
            if self._table.isRowHidden(row):
                continue
            if self._table.item(row, self.C_CHK).checkState() == Qt.CheckState.Checked:
                off_cb  = self._table.cellWidget(row, self.C_OFF)
                char, custom_char = self._char_type_for_row(row)
                off_lbl = off_cb.currentText()  if isinstance(off_cb,  QComboBox) else NO_OFFICER
                officer = self._resolve_generation_officer(self._off_map.get(off_lbl))
                selected.append((rec, char, officer, custom_char))

        if not selected:
            QMessageBox.warning(self, "Нет выбранных", "Не выбрано ни одной записи.")
            return

        self._gen_in_progress = True
        self._gen_btn.setEnabled(False)
        self._prog.setRange(0, len(selected))
        self._prog.setValue(0)
        self._prog.setVisible(True)
        _set_status(self._status, f"Создание: 0 / {len(selected)}…", "info")

        sig = _Sig(self)
        sig.msg.connect(lambda t, k: _set_status(self._status, t, k))
        sig.progress.connect(lambda c, _t: self._prog.setValue(c))
        sig.done.connect(self._on_gen_done)
        sig.error.connect(lambda m: QMessageBox.critical(self, "Ошибки", m))

        def task():
            errors  = []
            created = 0
            for n, (rec, char_type, officer, custom_char) in enumerate(selected, 1):
                try:
                    folder  = g._officer_folder(officer) if officer else 'Без_участкового'
                    out_dir = os.path.join(BASE_OUT_DIR, f'{year}_Q{quarter}', folder)
                    g.generate_one(rec, char_type, quarter, year, out_dir, officer, custom_char_text=custom_char)
                    created += 1
                except Exception as exc:
                    errors.append(f"  {rec.get('ФИО', '?')}: {exc}")
                sig.progress.emit(n, len(selected))
                sig.msg.emit(f"Создание: {n} / {len(selected)}…", "info")
            out_root = os.path.abspath(
                os.path.join(BASE_OUT_DIR, f'{year}_Q{quarter}'))
            sig.done.emit(out_root, created, len(selected))
            if errors:
                sig.error.emit("Ошибки при создании:\n" + "\n".join(errors[:20]))

        threading.Thread(target=task, daemon=True).start()

    def _on_gen_done(self, out_dir: str, count: int, total: int):
        self._gen_in_progress = False
        self._prog.setVisible(False)
        self._update_generate_enabled()
        if count <= 0:
            _set_status(self._status, f"Не создано файлов (0 / {total})", "error")
            return

        _set_status(self._status, f"Создано: {count} файлов  →  {os.path.basename(out_dir)}", "ok")
        if not os.path.isdir(out_dir):
            return

        r = QMessageBox.question(
            self, "Готово",
            f"Создано справок: {count}\nПапка:\n{out_dir}\n\nОткрыть папку?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            try:
                os.startfile(out_dir)
            except Exception as exc:
                QMessageBox.warning(self, "Не удалось открыть папку", str(exc))


# ── Вкладка «Участковые» ──────────────────────────────────────────────────────

class OfficersTab(QWidget):

    C_DIST = 0
    C_VAC  = 1
    C_FIO  = 2
    C_RANK = 3
    C_POS  = 4
    C_ADDR = 5
    C_REPL = 6

    def __init__(self, obzorki_tab: ObzorkiTab, on_open_settings=None):
        super().__init__()
        self._obzorki_tab = obzorki_tab
        self._on_open_settings = on_open_settings
        self._officers: list = []
        self._replacements: dict = {}
        self._on_dislocation_loaded = None   # callback → UnmatchedTab.refresh
        self._on_officers_changed = None
        self._dis_load_in_progress = False
        self._build()
        self._reload_from_db()

    # ── Построение UI ─────────────────────────────────────────────────────────

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # ─ Загрузка дислокации — одна кнопка ─────────────────────────────────
        grp_dis = QGroupBox("Дислокация участковых")
        dis_row = QHBoxLayout(grp_dis)
        dis_row.setContentsMargins(10, 8, 10, 10)
        self._dis_load_btn = QPushButton("📂  Выбрать и загрузить дислокацию…")
        self._dis_load_btn.setFixedHeight(30)
        self._dis_load_btn.clicked.connect(self._load_dislocation)
        dis_row.addWidget(self._dis_load_btn)
        self._dis_status = QLabel()
        self._dis_status.setObjectName("info")
        dis_row.addWidget(self._dis_status)
        dis_row.addStretch()
        if callable(self._on_open_settings):
            b_settings = _flat_btn("Настройки", 110)
            b_settings.clicked.connect(self._on_open_settings)
            dis_row.addWidget(b_settings)
        root.addWidget(grp_dis)

        # ─ Таблица участковых ─────────────────────────────────────────────────
        grp_off = QGroupBox("Список участковых   · двойной клик → редактировать")
        off_vbox = QVBoxLayout(grp_off)
        off_vbox.setContentsMargins(8, 8, 8, 8)
        off_vbox.setSpacing(6)

        self._off_table = QTableWidget(0, 7)
        self._off_table.setHorizontalHeaderLabels(
            ['Участок', 'Вак', 'ФИО', 'Звание', 'Должность', 'Адреса', 'Замещает'])
        self._off_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectItems)
        self._off_table.setSelectionMode(
            QAbstractItemView.SelectionMode.NoSelection)
        self._off_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked)
        self._off_table.setShowGrid(True)
        self._off_table.setAlternatingRowColors(True)
        self._off_table.verticalHeader().setVisible(False)
        self._off_table.verticalHeader().setDefaultSectionSize(28)
        self._off_table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        # Отключаем системную подсветку selected-item, чтобы не было "синей полосы" выше строки.
        self._off_table.setStyleSheet(
            "QTableView::item:selected {"
            "  background: transparent;"
            "  color: palette(text);"
            "  border: none;"
            "}"
            "QTableView::item {"
            "  outline: none;"
            "}"
        )
        self._off_table.itemChanged.connect(self._on_table_item_changed)

        hdr = self._off_table.horizontalHeader()
        hdr.setStretchLastSection(False)
        for col in (self.C_DIST, self.C_VAC, self.C_FIO, self.C_RANK, self.C_POS, self.C_ADDR, self.C_REPL):
            hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.Interactive)
        self._off_table.setColumnWidth(self.C_DIST, 46)
        self._off_table.setColumnWidth(self.C_VAC, 34)
        self._off_table.setColumnWidth(self.C_FIO,  250)
        self._off_table.setColumnWidth(self.C_RANK, 210)
        self._off_table.setColumnWidth(self.C_POS,  90)
        self._off_table.setColumnWidth(self.C_ADDR, 120)
        self._off_table.setColumnWidth(self.C_REPL, 220)
        off_vbox.addWidget(self._off_table)

        root.addWidget(grp_off, stretch=1)

    # ── Загрузка данных ───────────────────────────────────────────────────────

    def _reload_from_db(self):
        self._officers = db.all_officers()
        self._replacements = db.all_officer_replacements()
        self._populate_officers_table()
        cnt = len(self._officers)
        _set_status(self._dis_status,
                    f"  В БД: {cnt} сотрудников" if cnt else "  БД пуста",
                    "info" if cnt else "status")

    def get_active_officers(self) -> list:
        return list(self._officers)

    def _notify_officers_changed(self, refresh_unmatched: bool = True):
        self._obzorki_tab.reload_officers()
        if refresh_unmatched and self._on_officers_changed:
            self._on_officers_changed()

    def _populate_officers_table(self):
        self._off_table.blockSignals(True)
        self._off_table.setRowCount(0)
        self._off_table.setRowCount(len(self._officers))
        officer_labels = {o['id']: g._officer_label(o) for o in self._officers}

        for i, o in enumerate(self._officers):
            is_vac = bool(o.get('is_vacancy'))

            # C_DIST, C_FIO: редактируемые текстовые ячейки
            for col, key in (
                (self.C_DIST, 'district'),
                (self.C_FIO,  'fio'),
            ):
                it = QTableWidgetItem(str(o.get(key, '')))
                it.setFlags(
                    Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
                    | Qt.ItemFlag.ItemIsEditable)
                self._off_table.setItem(i, col, it)

            vac_item = QTableWidgetItem("Вак" if is_vac else "")
            vac_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            vac_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._off_table.setItem(i, self.C_VAC, vac_item)

            oid = o['id']

            # C_RANK: комбобокс звания + свободный ввод
            rank_cb = QComboBox()
            rank_cb.setEditable(True)
            rank_cb.addItems(RANK_OPTIONS)
            rank_cb.setFixedHeight(24)
            rank_cb.blockSignals(True)
            rank_cb.setCurrentText(_std_rank(o.get('rank', '')))
            rank_cb.blockSignals(False)
            rank_cb.currentTextChanged.connect(
                lambda val, _oid=oid, _r=i: self._on_rank_changed(_oid, _r, val))
            self._off_table.setCellWidget(i, self.C_RANK, rank_cb)

            # C_POS: комбобокс должности + свободный ввод
            pos_cb = QComboBox()
            pos_cb.setEditable(True)
            pos_cb.addItems(POSITION_OPTIONS)
            pos_cb.setFixedHeight(24)
            pos_cb.blockSignals(True)
            pos_cb.setCurrentText(_std_pos(o.get('position', '')))
            pos_cb.blockSignals(False)
            pos_cb.currentTextChanged.connect(
                lambda val, _oid=oid, _r=i: self._on_position_changed(_oid, _r, val))
            self._off_table.setCellWidget(i, self.C_POS, pos_cb)

            # C_ADDR: кнопка открытия адресов
            addr_btn = QPushButton("Адреса…")
            addr_btn.setFixedHeight(24)
            addr_btn.clicked.connect(
                lambda checked=False, _r=i: self._open_addr_dialog(_r))
            self._off_table.setCellWidget(i, self.C_ADDR, addr_btn)

            repl_cb = QComboBox()
            repl_cb.setMinimumWidth(170)
            repl_cb.setFixedHeight(24)
            repl_cb.addItem("—", None)
            for repl in self._officers:
                repl_id = repl.get('id')
                if repl_id == o.get('id'):
                    continue
                repl_cb.addItem(officer_labels.get(repl_id, g._officer_label(repl)), repl_id)
            repl_cb.blockSignals(True)
            repl_id = self._replacements.get(o.get('id'))
            if repl_id is not None:
                idx = repl_cb.findData(repl_id)
                if idx >= 0:
                    repl_cb.setCurrentIndex(idx)
                else:
                    repl_cb.setCurrentIndex(0)
            else:
                repl_cb.setCurrentIndex(0)
            repl_cb.blockSignals(False)
            repl_cb.currentIndexChanged.connect(
                lambda _idx, _oid=o.get('id'), _cb=repl_cb: self._on_replacement_changed(_oid, _cb)
            )
            self._off_table.setCellWidget(i, self.C_REPL, repl_cb)

        self._off_table.blockSignals(False)

    # ── Загрузка дислокации ───────────────────────────────────────────────────

    def _load_dislocation(self):
        if self._dis_load_in_progress:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл дислокации", "",
            "Word документ (*.docx);;Все файлы (*.*)")
        if not path:
            return
        self._dis_load_in_progress = True
        self._dis_load_btn.setEnabled(False)
        _set_status(self._dis_status, "  Загрузка…", "info")

        sig = _ObjSig(self)
        sig.done.connect(self._on_dislocation_loaded_ok)
        sig.error.connect(self._on_dislocation_loaded_error)

        def task():
            try:
                officers = pdis.parse_and_save(path)
                sig.done.emit(officers)
            except Exception as exc:
                sig.error.emit(str(exc))

        threading.Thread(target=task, daemon=True).start()

    def _on_dislocation_loaded_ok(self, officers):
        self._dis_load_in_progress = False
        self._dis_load_btn.setEnabled(True)
        self._reload_from_db()
        self._notify_officers_changed()
        if self._on_dislocation_loaded:
            self._on_dislocation_loaded()
        QMessageBox.information(
            self, "Дислокация загружена",
            f"Загружено сотрудников: {len(officers or [])}\nДанные сохранены в БД.")

    def _on_dislocation_loaded_error(self, message: str):
        self._dis_load_in_progress = False
        self._dis_load_btn.setEnabled(True)
        _set_status(self._dis_status, "  Ошибка загрузки", "error")
        QMessageBox.critical(self, "Ошибка загрузки дислокации", message)

    # ── Обработчики изменений в таблице ──────────────────────────────────────

    def _on_table_item_changed(self, item: QTableWidgetItem):
        row = item.row()
        if row < 0 or row >= len(self._officers):
            return
        officer = self._officers[row]
        col     = item.column()

        if col == self.C_DIST:
            new_val = item.text().strip()
            db.update_officer_field(officer['id'], 'district', new_val)
            officer['district'] = new_val
            self._notify_officers_changed()

        elif col == self.C_FIO:
            new_val = item.text().strip()
            db.update_officer_field(officer['id'], 'fio', new_val)
            officer['fio'] = new_val
            self._notify_officers_changed()

    def _on_rank_changed(self, officer_id: int, row: int, value: str):
        if row < len(self._officers):
            db.update_officer_field(officer_id, 'rank', value)
            self._officers[row]['rank'] = value
            self._notify_officers_changed(refresh_unmatched=False)

    def _on_position_changed(self, officer_id: int, row: int, value: str):
        if row < len(self._officers):
            db.update_officer_field(officer_id, 'position', value)
            self._officers[row]['position'] = value
            self._notify_officers_changed(refresh_unmatched=False)

    def _on_replacement_changed(self, officer_id: int, cb: QComboBox):
        if not isinstance(cb, QComboBox):
            return
        repl_id = cb.currentData()
        if repl_id is None:
            self._replacements.pop(officer_id, None)
            db.set_officer_replacement(officer_id, None)
        else:
            self._replacements[officer_id] = int(repl_id)
            db.set_officer_replacement(officer_id, int(repl_id))
        self._notify_officers_changed(refresh_unmatched=False)

    # ── Адреса участкового ────────────────────────────────────────────────────

    def _open_addr_dialog(self, row: int):
        if row >= len(self._officers):
            return
        officer = self._officers[row]
        dlg = _AddrDialog(officer, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        new_addrs = dlg.get_text()

        # Проверка конфликтов: ищем других участковых, у которых есть
        # совпадение «улица + номер дома» из новых адресов
        conflicts = self._addr_conflicts(new_addrs, officer['id'])
        if conflicts:
            names = ', '.join(o.get('fio', '?') for o in conflicts[:5])
            r = QMessageBox.question(
                self, "Возможный конфликт адресов",
                f"Следующие адреса уже числятся за другими участковыми:\n{names}\n\n"
                "Сохранить всё равно?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if r != QMessageBox.StandardButton.Yes:
                return

        db.update_officer_addresses(officer['id'], new_addrs)
        self._officers[row]['addresses'] = new_addrs
        self._notify_officers_changed()

    def _addr_conflicts(self, new_addrs: str, officer_id: int) -> list:
        """Ищет других участковых, у кого совпадают ключевое слово + номер дома."""
        results = []
        lines = [l.strip() for l in new_addrs.split('\n') if len(l.strip()) > 8]
        for line in lines[:10]:
            kws   = g._keywords_from(line)
            house = g._extract_house(line)
            if not kws or not house:
                continue
            kw_n  = g._norm(kws[0])
            for o in self._officers:
                if o['id'] == officer_id or o in results:
                    continue
                addr_n = g._norm(o.get('addresses', ''))
                if kw_n in addr_n and house in addr_n:
                    results.append(o)
        return results


# ── Вкладка «Нераспознанные» ──────────────────────────────────────────────────

class UnmatchedTab(QWidget):
    """Ручное назначение участковых для подучётных с нераспознанным адресом."""

    def __init__(self, obzorki_tab: ObzorkiTab, officers_tab: OfficersTab):
        super().__init__()
        self._obzorki_tab  = obzorki_tab
        self._officers_tab = officers_tab
        self._build()

    # ── Построение UI ─────────────────────────────────────────────────────────

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        hint = QLabel(
            "Подучётные, для которых адрес не совпал ни с одним участком. "
            "Назначьте участкового вручную и нажмите «Сохранить».")
        hint.setObjectName("status")
        hint.setWordWrap(True)
        root.addWidget(hint)

        self._um_table = QTableWidget(0, 4)
        self._um_table.setHorizontalHeaderLabels(
            ['ФИО', 'Дата рожд.', 'Адрес', 'Участковый'])
        self._um_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self._um_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._um_table.setShowGrid(True)
        self._um_table.setAlternatingRowColors(True)
        self._um_table.verticalHeader().setVisible(False)
        self._um_table.verticalHeader().setDefaultSectionSize(26)
        um_hdr = self._um_table.horizontalHeader()
        um_hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        um_hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        um_hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        um_hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self._um_table.setColumnWidth(1, 90)
        self._um_table.setColumnWidth(3, 240)
        root.addWidget(self._um_table, stretch=1)

        ctrl = QHBoxLayout()
        b_refresh = _flat_btn("↻ Обновить список")
        b_refresh.clicked.connect(self.refresh)
        ctrl.addWidget(b_refresh)
        ctrl.addStretch()
        self._save_btn = QPushButton("  Сохранить назначения  ")
        self._save_btn.setFixedHeight(32)
        self._save_btn.setEnabled(False)
        self._save_btn.clicked.connect(self._save_assignments)
        ctrl.addWidget(self._save_btn)
        root.addLayout(ctrl)

    # ── Данные ────────────────────────────────────────────────────────────────

    def refresh(self):
        """Обновляет список нераспознанных записей."""
        officers       = self._officers_tab.get_active_officers()
        off_labels     = [NO_OFFICER] + [g._officer_label(o) for o in officers]
        officers_by_id = {o['id']: o for o in officers}
        unmatched      = self._obzorki_tab.get_unmatched_records()

        self._um_table.blockSignals(True)
        self._um_table.setRowCount(0)
        self._um_table.setRowCount(len(unmatched))

        for i, (rec, row_idx) in enumerate(unmatched):
            fio_val  = rec.get('ФИО', '')
            dob_val  = rec.get('Дата рождения', '')
            addr_val = rec.get('Место жительства', '').replace('\n', ' ')

            for col, val in ((0, fio_val), (1, dob_val), (2, addr_val)):
                item = QTableWidgetItem(val)
                item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                if col == 2:
                    item.setToolTip(rec.get('Место жительства', ''))
                self._um_table.setItem(i, col, item)

            off_cb = QComboBox()
            off_cb.addItems(off_labels)

            assigned_id = db.get_assignment(fio_val, dob_val)
            if assigned_id and assigned_id in officers_by_id:
                label = g._officer_label(officers_by_id[assigned_id])
                if label in off_labels:
                    off_cb.setCurrentText(label)

            off_cb.setProperty('record_fio', fio_val)
            off_cb.setProperty('record_dob', dob_val)
            off_cb.setProperty('table_row',  row_idx)
            self._um_table.setCellWidget(i, 3, off_cb)

        self._um_table.blockSignals(False)
        self._save_btn.setEnabled(self._um_table.rowCount() > 0)

    def _save_assignments(self):
        officers = self._officers_tab.get_active_officers()
        off_map  = {g._officer_label(o): o for o in officers}
        saved    = 0

        for i in range(self._um_table.rowCount()):
            cb = self._um_table.cellWidget(i, 3)
            if not isinstance(cb, QComboBox):
                continue
            fio     = cb.property('record_fio')
            dob     = cb.property('record_dob')
            row_idx = cb.property('table_row')
            label   = cb.currentText()
            officer = off_map.get(label)

            if officer:
                db.set_assignment(fio, dob, officer['id'])
                self._obzorki_tab.set_assignment_for_row(row_idx, label)
                saved += 1
            else:
                db.set_assignment(fio, dob, None)

        QMessageBox.information(self, "Сохранено",
                                f"Сохранено назначений: {saved}")
        self.refresh()


# ── Главное окно ───────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"АвтоОбзорки  v{APP_VERSION}")
        self.resize(1100, 780)
        self.setMinimumSize(860, 600)

        settings = _load_settings()
        self._dark = settings.get('dark_theme', False)
        self._char_templates = _get_char_templates(settings)
        _apply_char_templates(self._char_templates)

        # Вкладки
        self.obzorki_tab = ObzorkiTab(on_open_settings=self._open_settings)
        self.officers_tab = OfficersTab(self.obzorki_tab, on_open_settings=self._open_settings)
        self._tabs = QTabWidget()
        self._tabs.setDocumentMode(True)
        self._tabs.addTab(self.obzorki_tab, "📋  Обзорные справки")
        self._tabs.addTab(self.officers_tab, "👮  Участковые")

        central = QWidget()
        vbox = QVBoxLayout(central)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(0)
        vbox.addWidget(self._tabs)

        self.setCentralWidget(central)

    def _set_theme(self, dark: bool):
        self._dark = dark
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet("")
            _apply_palette(app, self._dark)
        self._save_app_settings()

    def _save_app_settings(self):
        s = _load_settings()
        s['dark_theme'] = self._dark
        s['char_templates'] = _norm_char_templates(self._char_templates)
        _save_settings(s)

    def _open_settings(self):
        dlg = _SettingsDialog(self._dark, self._char_templates, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._char_templates = dlg.get_templates()
        _apply_char_templates(self._char_templates)
        self._set_theme(dlg.is_dark_theme())


# ── Точка входа ───────────────────────────────────────────────────────────────

def main():
    if sys.platform == "win32":
        os.system("chcp 65001 > nul")

    if hasattr(QApplication, "setHighDpiScaleFactorRoundingPolicy"):
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.RoundPreferFloor
        )
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 9))
    settings = _load_settings()
    _apply_char_templates(_get_char_templates(settings))
    app.setStyleSheet("")
    _apply_palette(app, settings.get('dark_theme', False))

    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
