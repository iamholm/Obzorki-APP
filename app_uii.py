"""
app_uii.py — УИИ: Генератор обзорных справок + Управление участковыми
Интерфейс: PySide6 (Qt 6)
"""

import os
import sys
import json
import datetime
import threading
import time as pytime
import ctypes

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox,
    QLabel, QPushButton, QLineEdit, QFileDialog,
    QSpinBox, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QProgressBar, QFrame, QTextEdit, QScrollArea,
    QAbstractItemView, QMessageBox, QDialog, QAbstractSpinBox,
    QStyleOptionButton, QStyle, QSizePolicy, QInputDialog, QSplashScreen,
)
from PySide6.QtCore import Qt, Signal, QObject, QRect, QTimer
from PySide6.QtGui import (
    QFont, QColor, QPalette, QShortcut, QKeySequence,
    QIcon, QPixmap, QPainter, QLinearGradient, QPen, QPainterPath, QRegion,
)

import extract_uii as uii
import db
import parse_dislocation as pdis
import gen_obzorka as g

# ── Настройки приложения ───────────────────────────────────────────────────────

_SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.json')
APP_ICON_CANDIDATES = (
    "autoobzorki.ico",
    "app.ico",
    "icon.ico",
    "favicon.ico",
    "autoobzorki.png",
    "icon.png",
    "favicon-32x32.png",
    "favicon-16x16.png",
)


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


def _app_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _app_search_dirs() -> list[str]:
    dirs: list[str] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if isinstance(meipass, str) and meipass:
        dirs.append(os.path.abspath(meipass))
    base = _app_base_dir()
    dirs.append(base)
    internal = os.path.join(base, "_internal")
    if os.path.isdir(internal):
        dirs.append(internal)
    local = os.path.dirname(os.path.abspath(__file__))
    dirs.append(local)

    out: list[str] = []
    seen: set[str] = set()
    for d in dirs:
        key = os.path.normcase(os.path.abspath(d))
        if key in seen:
            continue
        seen.add(key)
        out.append(os.path.abspath(d))
    return out


def _resolve_app_icon_path() -> str | None:
    for base in _app_search_dirs():
        for name in APP_ICON_CANDIDATES:
            path = os.path.join(base, name)
            if os.path.isfile(path):
                return path
    return None


def _load_app_icon() -> QIcon | None:
    path = _resolve_app_icon_path()
    if not path:
        return None
    icon = QIcon(path)
    if icon.isNull():
        return None
    return icon


def _build_splash_pixmap(progress: float = 0.0) -> QPixmap:
    width, height = 520, 286
    radius = 16
    p = max(0.0, min(1.0, float(progress)))

    pix = QPixmap(width, height)
    pix.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pix)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

    clip = QPainterPath()
    clip.addRoundedRect(0.5, 0.5, width - 1, height - 1, radius, radius)
    painter.setClipPath(clip)

    bg = QLinearGradient(0, 0, width, height)
    bg.setColorAt(0.0, QColor("#1a2f5c"))
    bg.setColorAt(0.45, QColor("#0f1d36"))
    bg.setColorAt(1.0, QColor("#0a1221"))
    painter.fillRect(0, 0, width, height, bg)

    glow = QLinearGradient(0, 0, width, 0)
    glow.setColorAt(0.0, QColor(255, 255, 255, 34))
    glow.setColorAt(1.0, QColor(255, 255, 255, 0))
    painter.fillRect(0, 0, width, 6, glow)

    painter.setClipping(False)
    painter.setPen(QPen(QColor(255, 255, 255, 52), 1))
    painter.setBrush(QColor(255, 255, 255, 16))
    painter.drawRoundedRect(0.5, 0.5, width - 1, height - 1, radius, radius)

    logo_x, logo_y = 24, 22
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#4f86ff"))
    painter.drawRoundedRect(logo_x, logo_y, 66, 66, 14, 14)
    painter.setPen(QColor("white"))
    logo_font = QFont("Segoe UI", 20)
    logo_font.setBold(True)
    painter.setFont(logo_font)
    painter.drawText(logo_x, logo_y, 66, 66, int(Qt.AlignmentFlag.AlignCenter), "АО")

    title_x = logo_x + 82
    painter.setPen(QColor("white"))
    title_font = QFont("Segoe UI", 24)
    title_font.setBold(True)
    painter.setFont(title_font)
    painter.drawText(
        title_x,
        logo_y + 4,
        width - (title_x + 24),
        36,
        int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter),
        "АвтоОбзорки",
    )

    painter.setFont(QFont("Segoe UI", 11))
    painter.setPen(QColor("#d9e6ff"))
    painter.drawText(
        title_x,
        logo_y + 39,
        width - (title_x + 24),
        20,
        int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter),
        f"Версия {APP_VERSION}",
    )

    painter.setFont(QFont("Segoe UI", 10))
    painter.setPen(QColor("#b9cae8"))
    painter.drawText(
        24,
        logo_y + 84,
        width - 48,
        22,
        int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter),
        "Генератор обзорных справок",
    )

    status_y = height - 58
    painter.setFont(QFont("Segoe UI", 9))
    painter.setPen(QColor("#9db4d8"))
    painter.drawText(24, status_y, 220, 20, int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter), "Инициализация")
    painter.drawText(width - 70, status_y, 46, 20, int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter), f"{int(p * 100)}%")

    bar_x, bar_y = 24, height - 32
    bar_w, bar_h = width - 48, 10
    painter.setPen(QPen(QColor(255, 255, 255, 40), 1))
    painter.setBrush(QColor(8, 15, 30, 150))
    painter.drawRoundedRect(bar_x, bar_y, bar_w, bar_h, 5, 5)

    fill_w = int((bar_w - 2) * p)
    if fill_w > 0:
        fill = QLinearGradient(bar_x, 0, bar_x + bar_w, 0)
        fill.setColorAt(0.0, QColor("#72c1ff"))
        fill.setColorAt(1.0, QColor("#5f7cff"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(fill)
        painter.drawRoundedRect(bar_x + 1, bar_y + 1, fill_w, bar_h - 2, 4, 4)

    painter.end()
    return pix


def _apply_splash_shape(splash: QSplashScreen, radius: int = 16) -> None:
    pix = splash.pixmap()
    if pix.isNull():
        return
    path = QPainterPath()
    path.addRoundedRect(0.0, 0.0, float(pix.width()), float(pix.height()), float(radius), float(radius))
    splash.setMask(QRegion(path.toFillPolygon().toPolygon()))


def _show_startup_splash(app: QApplication, app_icon: QIcon | None = None) -> QSplashScreen:
    splash = QSplashScreen(_build_splash_pixmap(0.0))
    splash.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    splash.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
    splash.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
    if app_icon is not None:
        splash.setWindowIcon(app_icon)
    _apply_splash_shape(splash)
    splash.show()

    screen = app.primaryScreen()
    if screen is not None:
        center = screen.availableGeometry().center()
        rect = splash.frameGeometry()
        rect.moveCenter(center)
        splash.move(rect.topLeft())
    app.processEvents()
    return splash


def _finish_startup_splash(
    app: QApplication,
    splash: QSplashScreen,
    started_at: float,
    minimum_seconds: float = 1.0,
) -> None:
    while True:
        elapsed = pytime.monotonic() - started_at
        progress = min(1.0, elapsed / minimum_seconds)
        splash.setPixmap(_build_splash_pixmap(progress))
        _apply_splash_shape(splash)
        app.processEvents()
        if elapsed >= minimum_seconds:
            break
        pytime.sleep(0.03)


def _to_bool(value, default=False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        v = value.strip().lower()
        if v in ("1", "true", "yes", "on", "dark", "темная", "тёмная"):
            return True
        if v in ("0", "false", "no", "off", "light", "светлая"):
            return False
    return bool(default)


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


def _norm_officer_char_templates(raw) -> dict:
    """Нормализует персональные шаблоны: {officer_id: {char_type: [texts]}}."""
    result = {}
    if not isinstance(raw, dict):
        return result
    for officer_id, templates in raw.items():
        try:
            off_id = int(officer_id)
        except Exception:
            continue
        if not isinstance(templates, dict):
            continue
        by_type = {}
        for key in g.CHAR_OPTIONS:
            value = templates.get(key)
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
                by_type[key] = items
        if by_type:
            result[off_id] = by_type
    return result


def _effective_char_templates(base_templates: dict, officer_templates: dict = None) -> dict:
    """Сливает глобальные и персональные шаблоны (персональные имеют приоритет)."""
    merged = _norm_char_templates(base_templates)
    if not isinstance(officer_templates, dict):
        return merged
    for key in g.CHAR_OPTIONS:
        raw = officer_templates.get(key)
        items = []
        if isinstance(raw, str):
            text = _clean_template_text(raw)
            if text:
                items.append(text)
        elif isinstance(raw, (list, tuple)):
            for item in raw:
                if isinstance(item, str):
                    text = _clean_template_text(item)
                    if text:
                        items.append(text)
        if items:
            merged[key] = items
    return merged


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


def _normalize_person_address(text: str) -> str:
    return uii.normalize_address_line(text or "")


def _normalize_inline_text(text: str) -> str:
    return " ".join(str(text or "").replace("\r", " ").replace("\n", " ").split()).strip()

# ── Константы ──────────────────────────────────────────────────────────────────
APP_VERSION    = "3.1"
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
            addr_src = _normalize_person_address(rec.get('Место жительства', '') or '')
            addr = QTableWidgetItem(addr_src)
            addr.setToolTip(addr_src)
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
            result.append(_normalize_person_address(item.text() if item else ""))
        return result


class _PersonCardDialog(QDialog):
    """Редактирование карточки подучётного (поля, идущие в итоговый docx)."""

    def __init__(self, rec: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Карточка подучётного")
        self.resize(860, 860)
        rec = rec or {}

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        hint = QLabel(
            "Изменения применяются только после подтверждения.\n"
            "Редактируются поля, которые используются при генерации справки."
        )
        hint.setObjectName("status")
        hint.setWordWrap(True)
        root.addWidget(hint)

        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        root.addLayout(grid)

        r = 0
        grid.addWidget(QLabel("ФИО:"), r, 0)
        self._fio = QLineEdit(_normalize_inline_text(rec.get("ФИО", "")))
        grid.addWidget(self._fio, r, 1)
        r += 1

        grid.addWidget(QLabel("Дата рождения:"), r, 0)
        self._dob = QLineEdit(_normalize_inline_text(rec.get("Дата рождения", "")))
        self._dob.setPlaceholderText("ДД.ММ.ГГГГ")
        grid.addWidget(self._dob, r, 1)
        r += 1

        grid.addWidget(QLabel("Место жительства:"), r, 0, alignment=Qt.AlignmentFlag.AlignTop)
        self._address = QTextEdit()
        self._address.setFixedHeight(66)
        self._address.setPlainText(_normalize_person_address(rec.get("Место жительства", "")))
        grid.addWidget(self._address, r, 1)
        r += 1

        grid.addWidget(QLabel("Суд (когда, кем):"), r, 0, alignment=Qt.AlignmentFlag.AlignTop)
        self._court = QTextEdit()
        self._court.setFixedHeight(86)
        self._court.setPlainText(_normalize_inline_text(rec.get("Суд (когда, кем)", "")))
        grid.addWidget(self._court, r, 1)
        r += 1

        grid.addWidget(QLabel("Обязанности:"), r, 0, alignment=Qt.AlignmentFlag.AlignTop)
        self._duties = QTextEdit()
        self._duties.setFixedHeight(86)
        self._duties.setPlainText(_normalize_inline_text(rec.get("Обязанности", "")))
        grid.addWidget(self._duties, r, 1)
        r += 1

        grid.addWidget(QLabel("Окончание срока:"), r, 0)
        self._end_date = QLineEdit(_normalize_inline_text(rec.get("Окончание срока", "")))
        self._end_date.setPlaceholderText("ДД.ММ.ГГГГ")
        grid.addWidget(self._end_date, r, 1)
        r += 1

        grid.addWidget(QLabel("Место работы (учебы):"), r, 0)
        self._work_place = QLineEdit(_normalize_inline_text(rec.get("Место работы (учебы)", "")))
        grid.addWidget(self._work_place, r, 1)
        r += 1

        grid.addWidget(QLabel("Телефон:"), r, 0)
        self._phone = QLineEdit(_normalize_inline_text(rec.get("Телефон", "")))
        grid.addWidget(self._phone, r, 1)
        r += 1

        grid.addWidget(QLabel("Характеристика:"), r, 0, alignment=Qt.AlignmentFlag.AlignTop)
        self._char_text = QTextEdit()
        self._char_text.setFixedHeight(96)
        self._char_text.setPlainText(
            _normalize_inline_text(rec.get("Характеристика", rec.get("Характеристика (п.8)", "")))
        )
        grid.addWidget(self._char_text, r, 1)
        r += 1

        grid.addWidget(QLabel("Связи:"), r, 0, alignment=Qt.AlignmentFlag.AlignTop)
        self._links = QTextEdit()
        self._links.setFixedHeight(66)
        self._links.setPlainText(_normalize_inline_text(rec.get("Связи", rec.get("Связи (п.9)", ""))))
        grid.addWidget(self._links, r, 1)
        r += 1

        grid.addWidget(QLabel("Приметы:"), r, 0)
        self._features = QLineEdit(_normalize_inline_text(rec.get("Приметы", rec.get("Приметы (п.10)", ""))))
        grid.addWidget(self._features, r, 1)
        r += 1

        grid.addWidget(QLabel("Сезонная одежда:"), r, 0)
        self._season = QLineEdit(
            _normalize_inline_text(rec.get("Сезонная одежда", rec.get("Сезонная одежда (п.11)", "")))
        )
        grid.addWidget(self._season, r, 1)
        r += 1

        grid.addWidget(QLabel("Нарушения:"), r, 0, alignment=Qt.AlignmentFlag.AlignTop)
        self._violations = QTextEdit()
        self._violations.setFixedHeight(66)
        self._violations.setPlainText(_normalize_inline_text(rec.get("Нарушения", rec.get("Нарушения (п.12)", ""))))
        grid.addWidget(self._violations, r, 1)
        r += 1

        grid.addWidget(QLabel("Проверка ИЦ:"), r, 0)
        self._ic_check = QLineEdit(_normalize_inline_text(rec.get("Проверка ИЦ", rec.get("Проверка ИЦ (п.13)", ""))))
        grid.addWidget(self._ic_check, r, 1)
        r += 1

        grid.setRowStretch(r, 1)
        grid.setColumnStretch(1, 1)

        btns = QHBoxLayout()
        btns.addStretch()
        b_cancel = _flat_btn("Отмена")
        b_cancel.clicked.connect(self.reject)
        btns.addWidget(b_cancel)
        b_save = QPushButton("  Сохранить  ")
        b_save.clicked.connect(self.accept)
        btns.addWidget(b_save)
        root.addLayout(btns)

    def get_data(self) -> dict:
        return {
            "ФИО": _normalize_inline_text(self._fio.text()),
            "Дата рождения": _normalize_inline_text(self._dob.text()),
            "Место жительства": _normalize_person_address(self._address.toPlainText()),
            "Суд (когда, кем)": _normalize_inline_text(self._court.toPlainText()),
            "Обязанности": _normalize_inline_text(self._duties.toPlainText()),
            "Окончание срока": _normalize_inline_text(self._end_date.text()),
            "Место работы (учебы)": _normalize_inline_text(self._work_place.text()),
            "Телефон": _normalize_inline_text(self._phone.text()),
            "Характеристика": _normalize_inline_text(self._char_text.toPlainText()),
            "Связи": _normalize_inline_text(self._links.toPlainText()),
            "Приметы": _normalize_inline_text(self._features.text()),
            "Сезонная одежда": _normalize_inline_text(self._season.text()),
            "Нарушения": _normalize_inline_text(self._violations.toPlainText()),
            "Проверка ИЦ": _normalize_inline_text(self._ic_check.text()),
        }


# ── Диалог настроек ─────────────────────────────────────────────────────────────

class _SettingsDialog(QDialog):
    """Параметры интерфейса и шаблонов характеристик."""

    _GLOBAL_SCOPE_TEXT = "Общие (по умолчанию)"

    def __init__(
        self,
        dark_theme: bool,
        char_templates: dict,
        officers: list = None,
        officer_char_templates: dict = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.resize(900, 680)

        self._global_templates = _norm_char_templates(char_templates)
        self._officers = [
            o for o in (officers or [])
            if isinstance(o, dict) and o.get('id') is not None
        ]
        self._officer_templates = _norm_officer_char_templates(officer_char_templates)
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

        scope_row = QHBoxLayout()
        scope_row.setSpacing(8)
        scope_row.addWidget(QLabel("Набор шаблонов:"))
        self._scope_cb = QComboBox()
        self._scope_cb.setFixedWidth(360)
        self._scope_cb.addItem(self._GLOBAL_SCOPE_TEXT, None)
        for officer in self._officers:
            off_id = int(officer.get('id'))
            self._scope_cb.addItem(g._officer_label(officer), off_id)
        self._scope_cb.currentIndexChanged.connect(lambda _idx: self._on_scope_changed())
        scope_row.addWidget(self._scope_cb)
        scope_row.addStretch()
        p_layout.addLayout(scope_row)

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

    def _current_scope_officer_id(self):
        if not hasattr(self, "_scope_cb"):
            return None
        data = self._scope_cb.currentData()
        if data is None:
            return None
        try:
            return int(data)
        except Exception:
            return None

    def _templates_for_current_type(self) -> list:
        off_id = self._current_scope_officer_id()
        if off_id is None:
            templates = self._global_templates.get(self._current_type)
            if not templates:
                templates = [_default_template_for(self._current_type)]
                self._global_templates[self._current_type] = list(templates)
            return list(templates)

        off_templates = self._officer_templates.get(off_id, {})
        templates = off_templates.get(self._current_type)
        if not templates:
            templates = self._global_templates.get(self._current_type) or [
                _default_template_for(self._current_type)
            ]
        return list(templates)

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

        off_id = self._current_scope_officer_id()
        if off_id is None:
            self._global_templates[self._current_type] = list(texts)
            return

        base = self._global_templates.get(self._current_type) or [_default_template_for(self._current_type)]
        if texts == base:
            by_type = self._officer_templates.get(off_id)
            if by_type:
                by_type.pop(self._current_type, None)
                if not any(by_type.get(k) for k in g.CHAR_OPTIONS):
                    self._officer_templates.pop(off_id, None)
            return

        self._officer_templates.setdefault(off_id, {})[self._current_type] = list(texts)

    def _render_templates(self):
        templates = self._templates_for_current_type()
        self._clear_templates_ui()
        for i, text in enumerate(templates):
            self._add_template_row_widget(i, text)

    def _on_scope_changed(self):
        if self._syncing:
            return
        self._commit_current_rows()
        self._render_templates()

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
        off_id = self._current_scope_officer_id()
        if off_id is None:
            self._global_templates[self._current_type] = list(templates)
        else:
            self._officer_templates.setdefault(off_id, {})[self._current_type] = list(templates)
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
        off_id = self._current_scope_officer_id()
        if off_id is None:
            self._global_templates = _norm_char_templates({})
        else:
            by_type = self._officer_templates.get(off_id)
            if by_type:
                by_type.pop(self._current_type, None)
                if not any(by_type.get(k) for k in g.CHAR_OPTIONS):
                    self._officer_templates.pop(off_id, None)
        self._current_type = self._char_type_cb.currentText() or g.CHAR_OPTIONS[0]
        self._render_templates()

    def is_dark_theme(self) -> bool:
        return self._theme_cb.currentIndex() == 1

    def get_templates(self) -> dict:
        self._commit_current_rows()
        return _norm_char_templates(self._global_templates)

    def get_officer_templates(self) -> dict:
        self._commit_current_rows()
        return _norm_officer_char_templates(self._officer_templates)


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
    C_DOB  = 3
    C_CAT  = 4
    C_CHAR_POS = 5
    C_CHAR_NEU = 6
    C_CHAR_NEG = 7
    C_CHAR_CUS = 8
    C_OFF  = 9
    C_ADDR = 10

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
        self._startup_autoload_done = False

        db.init_db()
        self._build()
        self._restore_ui_settings()
        self.reload_officers()
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
        self._table = QTableWidget(0, 11)
        self._table.setHorizontalHeaderLabels(
            ['', '№', 'ФИО', 'Дата рожд.', 'Кат.', '+', '0', '-', 'Инд.', 'Участковый', 'Адрес'])
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
        for col in (self.C_CHK, self.C_NUM, self.C_FIO, self.C_DOB,
                    self.C_CAT, self.C_CHAR_POS, self.C_CHAR_NEU, self.C_CHAR_NEG,
                    self.C_CHAR_CUS, self.C_OFF, self.C_ADDR):
            hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.Interactive)
        hdr.setStretchLastSection(False)
        self._table.setColumnWidth(self.C_CHK,  30)
        self._table.setColumnWidth(self.C_NUM,  38)
        self._table.setColumnWidth(self.C_FIO,  200)
        self._table.setColumnWidth(self.C_DOB,  98)
        self._table.setColumnWidth(self.C_CAT,  75)
        self._table.setColumnWidth(self.C_CHAR_POS, 48)
        self._table.setColumnWidth(self.C_CHAR_NEU, 48)
        self._table.setColumnWidth(self.C_CHAR_NEG, 48)
        self._table.setColumnWidth(self.C_CHAR_CUS, 58)
        self._table.setColumnWidth(self.C_OFF,  320)
        self._table.setColumnWidth(self.C_ADDR, 90)

        root.addWidget(self._table, stretch=1)

        # ─ Статус + генерация ─────────────────────────────────────────────────
        bot = QHBoxLayout()
        self._status = QLabel("")
        self._status.setObjectName("status")
        self._status.setVisible(False)

        self._unmatched_info = QLabel("")
        self._unmatched_info.setStyleSheet("color: #cf222e;")
        bot.addWidget(self._unmatched_info)

        self._last_source_info = QLabel("")
        self._last_source_info.setObjectName("status")
        bot.addSpacing(12)
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
        self._quarter_lbl = QLabel("Квартал:")
        right_ctrl.addWidget(self._quarter_lbl)
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
        self._year_lbl = QLabel("Год:")
        right_ctrl.addWidget(self._year_lbl)
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

    def run_startup_autoload(self):
        if self._startup_autoload_done:
            return
        self._startup_autoload_done = True
        self._autoload_last_source()

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

    def _warn_edit_locked(self):
        QMessageBox.information(
            self,
            "Редактирование заблокировано",
            "Чтобы изменить данные, выберите конкретного участкового в фильтре,\n"
            "а не режим «Все участковые».",
        )

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
        addr_btn = self._table.cellWidget(row, self.C_ADDR)
        if isinstance(addr_btn, QPushButton):
            addr_btn.setEnabled(enabled)

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
        row = item.row()
        col = item.column()

        if self._is_edit_locked():
            self._warn_edit_locked()
            return

        if col == self.C_CHAR_CUS:
            current_type, current_text = self._char_type_for_row(row)
            seed_text = current_text if current_type == CUSTOM_CHAR_OPTION else ""
            new_text = self._edit_custom_characteristic(row, seed_text)
            if not new_text:
                return
            self._set_row_char_type(row, CUSTOM_CHAR_OPTION, new_text, persist=True)
            _set_status(self._status, "Индивидуальная характеристика сохранена", "ok")
            return

        if col in (self.C_CHK, self.C_CHAR_POS, self.C_CHAR_NEU, self.C_CHAR_NEG):
            return

        self._edit_row_person_card(row)

    def _effective_char_text_for_row(self, row: int) -> str:
        if row < 0 or row >= len(self._records):
            return ""
        rec = self._records[row]
        char_type, custom_text = self._char_type_for_row(row)
        if char_type == CUSTOM_CHAR_OPTION and custom_text:
            return _normalize_inline_text(custom_text)

        # Явно сохранённый текст (legacy/ручной) имеет приоритет.
        manual = _normalize_inline_text(rec.get("Характеристика", rec.get("Характеристика (п.8)", "")))
        if manual:
            return manual

        safe_char = char_type if char_type in g.CHAR_OPTIONS else "нейтральная"
        base_templates = _norm_char_templates(g.CHAR_TEXTS)

        off_cb = self._table.cellWidget(row, self.C_OFF)
        off_label = off_cb.currentText() if isinstance(off_cb, QComboBox) else NO_OFFICER
        officer = self._resolve_generation_officer(self._off_map.get(off_label))
        if isinstance(officer, dict):
            try:
                off_id = int(officer.get("id"))
            except Exception:
                off_id = None
            if off_id is not None:
                per_off = _norm_officer_char_templates(db.all_officer_char_templates()).get(off_id)
                base_templates = _effective_char_templates(base_templates, per_off)

        values = base_templates.get(safe_char) or []
        if values:
            return _normalize_inline_text(values[0])
        return _normalize_inline_text(_default_template_for(safe_char))

    def _effective_links_text_for_row(self, row: int) -> str:
        if row < 0 or row >= len(self._records):
            return ""
        rec = self._records[row]
        manual = _normalize_inline_text(rec.get("Связи", rec.get("Связи (п.9)", "")))
        if manual:
            return manual
        char_type, _custom = self._char_type_for_row(row)
        safe_char = char_type if char_type in g.CHAR_OPTIONS else "нейтральная"
        return _normalize_inline_text(g.CONN_TEXTS.get(safe_char, g.CONN_TEXTS.get("нейтральная", "")))

    def _edit_row_person_card(self, row: int):
        if self._is_edit_locked():
            self._warn_edit_locked()
            return
        if row < 0 or row >= len(self._records):
            return
        rec = self._records[row]
        key_fio, key_dob = self._record_source_key(rec)
        if not key_fio:
            return

        old_char_text = self._effective_char_text_for_row(row)
        old_links_text = self._effective_links_text_for_row(row)

        old_data = {
            "ФИО": _normalize_inline_text(rec.get("ФИО", "")),
            "Дата рождения": _normalize_inline_text(rec.get("Дата рождения", "")),
            "Место жительства": _normalize_person_address(rec.get("Место жительства", "")),
            "Суд (когда, кем)": _normalize_inline_text(rec.get("Суд (когда, кем)", "")),
            "Обязанности": _normalize_inline_text(rec.get("Обязанности", "")),
            "Окончание срока": _normalize_inline_text(rec.get("Окончание срока", "")),
            "Место работы (учебы)": _normalize_inline_text(rec.get("Место работы (учебы)", "")),
            "Телефон": _normalize_inline_text(rec.get("Телефон", "")),
            "Характеристика": old_char_text,
            "Связи": old_links_text,
            "Приметы": _normalize_inline_text(rec.get("Приметы", rec.get("Приметы (п.10)", ""))),
            "Сезонная одежда": _normalize_inline_text(
                rec.get("Сезонная одежда", rec.get("Сезонная одежда (п.11)", ""))
            ),
            "Нарушения": _normalize_inline_text(rec.get("Нарушения", rec.get("Нарушения (п.12)", ""))),
            "Проверка ИЦ": _normalize_inline_text(rec.get("Проверка ИЦ", rec.get("Проверка ИЦ (п.13)", ""))),
        }

        dlg = _PersonCardDialog(old_data, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        new_data = dlg.get_data()
        if not new_data.get("ФИО"):
            QMessageBox.warning(self, "Пустое ФИО", "ФИО не может быть пустым.")
            return

        changed_fields = [
            fld for fld in old_data.keys()
            if _normalize_inline_text(old_data.get(fld, "")) != _normalize_inline_text(new_data.get(fld, ""))
        ]
        if not changed_fields:
            return

        confirm_msg = (
            "Применить изменения в карточке?\n\n"
            f"Изменено полей: {len(changed_fields)}\n"
            f"{', '.join(changed_fields)}"
        )
        r = QMessageBox.question(
            self,
            "Подтверждение изменений",
            confirm_msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return

        # ФИО и адрес сохраняются отдельными override-таблицами.
        db.set_person_fio_override(key_fio, key_dob, new_data.get("ФИО", ""))
        db.set_person_address_fix(key_fio, key_dob, new_data.get("Место жительства", ""))

        source = {
            "dob": _normalize_inline_text(rec.get("_source_dob", key_dob)),
            "court": _normalize_inline_text(rec.get("_source_court", "")),
            "duties": _normalize_inline_text(rec.get("_source_duties", "")),
            "end_date": _normalize_inline_text(rec.get("_source_end_date", "")),
            "work_place": _normalize_inline_text(rec.get("_source_work_place", "")),
            "phone": _normalize_inline_text(rec.get("_source_phone", "")),
            "links": _normalize_inline_text(rec.get("_source_links", "")),
            "features": _normalize_inline_text(rec.get("_source_features", "")),
            "season_clothes": _normalize_inline_text(rec.get("_source_season_clothes", "")),
            "violations": _normalize_inline_text(rec.get("_source_violations", "")),
            "ic_check": _normalize_inline_text(rec.get("_source_ic_check", "")),
        }
        existing_doc_overrides = db.all_person_doc_overrides().get((key_fio, key_dob), {})
        doc_overrides = dict(existing_doc_overrides) if isinstance(existing_doc_overrides, dict) else {}
        changed = set(changed_fields)

        def touch_override(changed_label: str, doc_key: str, new_value: str, source_value: str):
            if changed_label not in changed:
                return
            if _normalize_inline_text(new_value) == _normalize_inline_text(source_value):
                doc_overrides.pop(doc_key, None)
            else:
                doc_overrides[doc_key] = _normalize_inline_text(new_value)

        touch_override("Дата рождения", "dob", new_data.get("Дата рождения", ""), source["dob"])
        touch_override("Суд (когда, кем)", "court", new_data.get("Суд (когда, кем)", ""), source["court"])
        touch_override("Обязанности", "duties", new_data.get("Обязанности", ""), source["duties"])
        touch_override("Окончание срока", "end_date", new_data.get("Окончание срока", ""), source["end_date"])
        touch_override("Место работы (учебы)", "work_place", new_data.get("Место работы (учебы)", ""), source["work_place"])
        touch_override("Телефон", "phone", new_data.get("Телефон", ""), source["phone"])
        touch_override("Связи", "links", new_data.get("Связи", ""), source["links"])
        touch_override("Приметы", "features", new_data.get("Приметы", ""), source["features"])
        touch_override("Сезонная одежда", "season_clothes", new_data.get("Сезонная одежда", ""), source["season_clothes"])
        touch_override("Нарушения", "violations", new_data.get("Нарушения", ""), source["violations"])
        touch_override("Проверка ИЦ", "ic_check", new_data.get("Проверка ИЦ", ""), source["ic_check"])

        db.set_person_doc_overrides(key_fio, key_dob, doc_overrides)

        key = (key_fio, key_dob)
        affected_rows = []
        for idx, item_rec in enumerate(self._records):
            if self._record_source_key(item_rec) != key:
                continue
            item_rec["ФИО"] = new_data.get("ФИО", "")
            item_rec["Дата рождения"] = new_data.get("Дата рождения", "")
            item_rec["Место жительства"] = new_data.get("Место жительства", "")
            item_rec["Суд (когда, кем)"] = new_data.get("Суд (когда, кем)", "")
            item_rec["Обязанности"] = new_data.get("Обязанности", "")
            item_rec["Окончание срока"] = new_data.get("Окончание срока", "")
            item_rec["Место работы (учебы)"] = new_data.get("Место работы (учебы)", "")
            item_rec["Телефон"] = new_data.get("Телефон", "")
            item_rec["Связи"] = new_data.get("Связи", "")
            item_rec["Приметы"] = new_data.get("Приметы", "")
            item_rec["Сезонная одежда"] = new_data.get("Сезонная одежда", "")
            item_rec["Нарушения"] = new_data.get("Нарушения", "")
            item_rec["Проверка ИЦ"] = new_data.get("Проверка ИЦ", "")
            affected_rows.append(idx)

            fio_item = self._table.item(idx, self.C_FIO)
            if fio_item is not None:
                fio_item.setText(item_rec["ФИО"] or "—")
            dob_item = self._table.item(idx, self.C_DOB)
            if dob_item is not None:
                dob_item.setText(item_rec["Дата рождения"] or "—")

        new_char_text = _normalize_inline_text(new_data.get("Характеристика", ""))
        char_or_links_changed = ("Характеристика" in changed) or ("Связи" in changed)
        if char_or_links_changed:
            effective_custom = new_char_text or old_char_text or _default_template_for("нейтральная")
            for r_idx in affected_rows or [row]:
                self._set_row_char_type(r_idx, CUSTOM_CHAR_OPTION, effective_custom, persist=True)
                if 0 <= r_idx < len(self._records):
                    self._records[r_idx]["Характеристика"] = effective_custom
        else:
            for r_idx in affected_rows or [row]:
                if 0 <= r_idx < len(self._records):
                    self._records[r_idx]["Характеристика"] = new_char_text

        if old_data.get("Место жительства", "") != new_data.get("Место жительства", ""):
            self._auto_assign(rows=affected_rows or [row], update_status=False, confirm_changes=False)

        self._apply_officer_filter()
        self._refresh_unmatched_info()
        _set_status(self._status, f"Карточка обновлена: {new_data.get('ФИО', '—')}", "ok")

    def _edit_row_fio(self, row: int):
        if self._is_edit_locked():
            self._warn_edit_locked()
            return
        if row < 0 or row >= len(self._records):
            return
        rec = self._records[row]
        src_fio, src_dob = self._record_source_key(rec)
        old_fio = str(rec.get('ФИО', '') or '').strip()
        if not old_fio and not src_fio:
            return

        new_fio, ok = QInputDialog.getText(
            self,
            "Изменение ФИО",
            "Введите ФИО:",
            QLineEdit.EchoMode.Normal,
            old_fio or src_fio,
        )
        if not ok:
            return
        new_fio = " ".join((new_fio or "").split()).strip()
        if not new_fio:
            QMessageBox.warning(self, "Пустое ФИО", "ФИО не может быть пустым.")
            return
        if new_fio == (old_fio or src_fio):
            return

        r = QMessageBox.question(
            self,
            "Подтверждение изменения ФИО",
            (
                f"Изменить ФИО в выбранной записи?\n\n"
                f"Было: {old_fio}\n"
                f"Станет: {new_fio}"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return

        db.set_person_fio_override(src_fio, src_dob, new_fio)
        old_key = (src_fio, src_dob)
        for idx, r in enumerate(self._records):
            if self._record_source_key(r) != old_key:
                continue
            r['ФИО'] = new_fio
            fio_item = self._table.item(idx, self.C_FIO)
            if fio_item is not None:
                fio_item.setText(new_fio)
        self._apply_officer_filter()
        self._refresh_unmatched_info()

    def _edit_row_address(self, row: int):
        if self._is_edit_locked():
            self._warn_edit_locked()
            return
        if row < 0 or row >= len(self._records):
            return
        rec = self._records[row]
        fio = str(rec.get('ФИО', '') or '').strip()
        key_fio, key_dob = self._record_source_key(rec)
        old_addr = _normalize_person_address(rec.get('Место жительства', '') or '')

        new_addr, ok = QInputDialog.getMultiLineText(
            self,
            "Изменение адреса",
            f"Изменить адрес для:\n{fio or '—'}",
            old_addr,
        )
        if not ok:
            return
        new_addr = _normalize_person_address(new_addr or "")
        if new_addr == old_addr:
            return
        if not new_addr:
            QMessageBox.warning(self, "Пустой адрес", "Адрес не может быть пустым.")
            return

        old_view = old_addr if old_addr else "—"
        new_view = new_addr if new_addr else "—"
        msg = (
            "Применить изменение адреса?\n\n"
            f"Было:\n{old_view}\n\n"
            f"Станет:\n{new_view}"
        )
        r = QMessageBox.question(
            self,
            "Подтверждение изменения адреса",
            msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return

        key = (key_fio, key_dob)
        affected_rows = []
        for r in self._records:
            if self._record_source_key(r) == key:
                r['Место жительства'] = new_addr
        for idx, r in enumerate(self._records):
            if self._record_source_key(r) == key:
                affected_rows.append(idx)
        db.set_person_address_fix(key_fio, key_dob, new_addr)
        self._auto_assign(rows=affected_rows or [row], update_status=False, confirm_changes=False)
        self._apply_officer_filter()
        self._refresh_unmatched_info()

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

    def _record_source_key(self, rec: dict):
        if not isinstance(rec, dict):
            return "", ""
        fio = str(rec.get('_source_fio', rec.get('ФИО', '')) or '').strip()
        dob = str(rec.get('_source_dob', rec.get('Дата рождения', '')) or '').strip()
        return fio, dob

    def _prepare_records_identity(self):
        if not self._records:
            return
        for rec in self._records:
            if not isinstance(rec, dict):
                continue
            rec.setdefault('Место работы (учебы)', '')
            rec.setdefault('Телефон', '')
            rec.setdefault('Связи', rec.get('Связи (п.9)', ''))
            rec.setdefault('Приметы', rec.get('Приметы (п.10)', g.FEATURES_TEXT))
            rec.setdefault('Сезонная одежда', rec.get('Сезонная одежда (п.11)', g.SEASON_CLOTHES_TEXT))
            rec.setdefault('Нарушения', rec.get('Нарушения (п.12)', ''))
            rec.setdefault('Проверка ИЦ', rec.get('Проверка ИЦ (п.13)', getattr(g, "IC_CHECK_TEXT", "См. справку ИБД-Р")))
            src_fio = str(rec.get('ФИО', '') or '').strip()
            src_dob = str(rec.get('Дата рождения', '') or '').strip()
            rec['_source_fio'] = src_fio
            rec['_source_dob'] = src_dob
            rec['_source_court'] = _normalize_inline_text(rec.get('Суд (когда, кем)', ''))
            rec['_source_duties'] = _normalize_inline_text(rec.get('Обязанности', ''))
            rec['_source_end_date'] = _normalize_inline_text(rec.get('Окончание срока', ''))
            rec['_source_work_place'] = _normalize_inline_text(rec.get('Место работы (учебы)', ''))
            rec['_source_phone'] = _normalize_inline_text(rec.get('Телефон', ''))
            rec['_source_links'] = _normalize_inline_text(rec.get('Связи', rec.get('Связи (п.9)', '')))
            rec['_source_features'] = _normalize_inline_text(rec.get('Приметы', rec.get('Приметы (п.10)', g.FEATURES_TEXT)))
            rec['_source_season_clothes'] = _normalize_inline_text(
                rec.get('Сезонная одежда', rec.get('Сезонная одежда (п.11)', g.SEASON_CLOTHES_TEXT))
            )
            rec['_source_violations'] = _normalize_inline_text(rec.get('Нарушения', rec.get('Нарушения (п.12)', '')))
            rec['_source_ic_check'] = _normalize_inline_text(
                rec.get('Проверка ИЦ', rec.get('Проверка ИЦ (п.13)', getattr(g, "IC_CHECK_TEXT", "См. справку ИБД-Р")))
            )

    def _normalize_record_addresses(self):
        if not self._records:
            return
        for rec in self._records:
            if isinstance(rec, dict):
                rec['Место жительства'] = _normalize_person_address(
                    rec.get('Место жительства', '')
                )

    def _apply_saved_fio_overrides(self):
        overrides = db.all_person_fio_overrides()
        if not overrides or not self._records:
            return
        for rec in self._records:
            src_fio, src_dob = self._record_source_key(rec)
            fio_override = (overrides.get((src_fio, src_dob)) or '').strip()
            if fio_override:
                rec['ФИО'] = fio_override

    def _current_record_keys(self) -> set:
        keys = set()
        for rec in self._records:
            fio, dob = self._record_source_key(rec)
            if fio:
                keys.add((fio, dob))
        return keys

    def _apply_saved_address_fixes(self):
        fixes = db.all_person_address_fixes()
        if not fixes or not self._records:
            return
        for rec in self._records:
            fio, dob = self._record_source_key(rec)
            fixed = fixes.get((fio, dob))
            if fixed:
                rec['Место жительства'] = _normalize_person_address(fixed)

    def _apply_saved_doc_overrides(self):
        overrides = db.all_person_doc_overrides()
        if not overrides or not self._records:
            return
        for rec in self._records:
            fio, dob = self._record_source_key(rec)
            data = overrides.get((fio, dob))
            if not isinstance(data, dict):
                continue
            if 'dob' in data:
                rec['Дата рождения'] = _normalize_inline_text(data.get('dob', ''))
            if 'court' in data:
                rec['Суд (когда, кем)'] = _normalize_inline_text(data.get('court', ''))
            if 'duties' in data:
                rec['Обязанности'] = _normalize_inline_text(data.get('duties', ''))
            if 'end_date' in data:
                rec['Окончание срока'] = _normalize_inline_text(data.get('end_date', ''))
            if 'work_place' in data:
                rec['Место работы (учебы)'] = _normalize_inline_text(data.get('work_place', ''))
            if 'phone' in data:
                rec['Телефон'] = _normalize_inline_text(data.get('phone', ''))
            if 'links' in data:
                rec['Связи'] = _normalize_inline_text(data.get('links', ''))
            if 'features' in data:
                rec['Приметы'] = _normalize_inline_text(data.get('features', ''))
            if 'season_clothes' in data:
                rec['Сезонная одежда'] = _normalize_inline_text(data.get('season_clothes', ''))
            if 'violations' in data:
                rec['Нарушения'] = _normalize_inline_text(data.get('violations', ''))
            if 'ic_check' in data:
                rec['Проверка ИЦ'] = _normalize_inline_text(data.get('ic_check', ''))

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
        self._normalize_record_addresses()
        self._prepare_records_identity()
        self._apply_saved_fio_overrides()
        self._apply_saved_address_fixes()
        self._apply_saved_doc_overrides()
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
            key_fio, key_dob = self._record_source_key(rec)

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

            dob_item = QTableWidgetItem(dob_val or '—')
            dob_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            dob_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(i, self.C_DOB, dob_item)

            cat_full  = rec.get('Категория', '')
            cat_short = CAT_SHORT.get(cat_full, cat_full[:8])
            cat_item  = QTableWidgetItem(cat_short)
            cat_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            cat_item.setToolTip(cat_full)
            cat_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self._table.setItem(i, self.C_CAT, cat_item)

            pref = char_prefs.get((key_fio, key_dob))
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
            assigned_id = assignments.get((key_fio, key_dob))
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

            addr_btn = _flat_btn("Адрес...", 78)
            addr_btn.setFixedHeight(24)
            addr_btn.setToolTip("Изменить адрес в этой записи")
            addr_btn.clicked.connect(lambda _=False, _r=i: self._edit_row_address(_r))
            self._table.setCellWidget(i, self.C_ADDR, addr_btn)

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
        return self._record_source_key(rec)

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

        target_rows = self._target_rows_for_mass_action()
        changes = []
        for row in target_rows:
            cb = self._table.cellWidget(row, self.C_OFF)
            if not isinstance(cb, QComboBox):
                continue
            if (cb.currentText() or "").strip() == label:
                continue
            changes.append(row)

        if not changes:
            self._bulk_officer_updating = True
            self._bulk_officer_cb.setCurrentText(BULK_OFFICER_PLACEHOLDER)
            self._bulk_officer_updating = False
            return

        preview = []
        for row in changes[:8]:
            fio = str(self._records[row].get('ФИО', '') or '-')
            preview.append(f"- {fio}")
        more = "" if len(changes) <= 8 else f"\n... и ещё {len(changes) - 8}"
        msg = (
            f"Назначить «{label}» для {len(changes)} записей?\n\n"
            f"{chr(10).join(preview)}{more}\n\n"
            "Продолжить?"
        )
        r = QMessageBox.question(
            self,
            "Подтверждение назначения участкового",
            msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            self._bulk_officer_updating = True
            self._bulk_officer_cb.setCurrentText(BULK_OFFICER_PLACEHOLDER)
            self._bulk_officer_updating = False
            _set_status(self._status, "Массовое назначение участкового отменено", "info")
            return

        changed = 0
        self._bulk_assigning = True
        try:
            for row in changes:
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
                    addr = _normalize_person_address(new_addresses[idx] or "")
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
        base_char_templates = _norm_char_templates(g.CHAR_TEXTS)
        officer_char_overrides = _norm_officer_char_templates(db.all_officer_char_templates())
        char_templates_cache = {}

        def templates_for_officer(officer_obj: dict) -> dict:
            if not isinstance(officer_obj, dict):
                return base_char_templates
            try:
                off_id = int(officer_obj.get('id'))
            except Exception:
                return base_char_templates
            cached = char_templates_cache.get(off_id)
            if cached is not None:
                return cached
            merged = _effective_char_templates(
                base_char_templates,
                officer_char_overrides.get(off_id),
            )
            char_templates_cache[off_id] = merged
            return merged

        selected = []
        for row, rec in enumerate(self._records):
            if self._table.isRowHidden(row):
                continue
            if self._table.item(row, self.C_CHK).checkState() == Qt.CheckState.Checked:
                off_cb  = self._table.cellWidget(row, self.C_OFF)
                char, custom_char = self._char_type_for_row(row)
                off_lbl = off_cb.currentText()  if isinstance(off_cb,  QComboBox) else NO_OFFICER
                officer = self._resolve_generation_officer(self._off_map.get(off_lbl))
                selected.append((rec, char, officer, custom_char, templates_for_officer(officer)))

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
            for n, (rec, char_type, officer, custom_char, char_templates) in enumerate(selected, 1):
                try:
                    folder  = g._officer_folder(officer) if officer else 'Без_участкового'
                    out_dir = os.path.join(BASE_OUT_DIR, f'{year}_Q{quarter}', folder)
                    g.generate_one(
                        rec,
                        char_type,
                        quarter,
                        year,
                        out_dir,
                        officer,
                        custom_char_text=custom_char,
                        char_templates=char_templates,
                    )
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
        self._suspend_off_table_item_handler = False
        self._suspend_officer_combo_handlers = False
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
            "QTableView {"
            "  background: palette(base);"
            "  alternate-background-color: palette(alternate-base);"
            "  color: palette(text);"
            "  gridline-color: #3a3a3a;"
            "}"
            "QHeaderView::section {"
            "  background: palette(button);"
            "  color: palette(button-text);"
            "}"
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
                it.setData(Qt.ItemDataRole.UserRole, str(o.get(key, '')))
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
            rank_cb.setProperty('prev_value', rank_cb.currentText())
            rank_cb.activated.connect(
                lambda _idx, _oid=oid, _r=i, _cb=rank_cb: self._on_rank_changed(_oid, _r, _cb.currentText())
            )
            if rank_cb.lineEdit() is not None:
                rank_cb.lineEdit().editingFinished.connect(
                    lambda _oid=oid, _r=i, _cb=rank_cb: self._on_rank_changed(_oid, _r, _cb.currentText())
                )
            self._off_table.setCellWidget(i, self.C_RANK, rank_cb)

            # C_POS: комбобокс должности + свободный ввод
            pos_cb = QComboBox()
            pos_cb.setEditable(True)
            pos_cb.addItems(POSITION_OPTIONS)
            pos_cb.setFixedHeight(24)
            pos_cb.blockSignals(True)
            pos_cb.setCurrentText(_std_pos(o.get('position', '')))
            pos_cb.blockSignals(False)
            pos_cb.setProperty('prev_value', pos_cb.currentText())
            pos_cb.activated.connect(
                lambda _idx, _oid=oid, _r=i, _cb=pos_cb: self._on_position_changed(_oid, _r, _cb.currentText())
            )
            if pos_cb.lineEdit() is not None:
                pos_cb.lineEdit().editingFinished.connect(
                    lambda _oid=oid, _r=i, _cb=pos_cb: self._on_position_changed(_oid, _r, _cb.currentText())
                )
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
            repl_cb.setProperty('prev_repl_id', repl_cb.currentData())
            repl_cb.currentIndexChanged.connect(
                lambda _idx, _oid=o.get('id'), _cb=repl_cb, _r=i: self._on_replacement_changed(_oid, _cb, _r)
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

    def _confirm_officer_change(
        self,
        title: str,
        officer: dict,
        field_label: str,
        old_value: str,
        new_value: str,
    ) -> bool:
        fio = str((officer or {}).get('fio', '') or '—')
        district = str((officer or {}).get('district', '') or '—')
        old_v = old_value if str(old_value).strip() else "—"
        new_v = new_value if str(new_value).strip() else "—"
        r = QMessageBox.question(
            self,
            title,
            (
                f"Подтвердите изменение:\n"
                f"{fio} (участок {district})\n\n"
                f"{field_label}: {old_v} → {new_v}"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        return r == QMessageBox.StandardButton.Yes

    def _repl_label(self, cb: QComboBox, repl_id) -> str:
        if not isinstance(cb, QComboBox) or repl_id is None:
            return "—"
        idx = cb.findData(repl_id)
        if idx < 0:
            return "—"
        return cb.itemText(idx) or "—"

    def _on_table_item_changed(self, item: QTableWidgetItem):
        if self._suspend_off_table_item_handler:
            return
        row = item.row()
        if row < 0 or row >= len(self._officers):
            return
        officer = self._officers[row]
        col     = item.column()

        if col == self.C_DIST:
            old_val = str(item.data(Qt.ItemDataRole.UserRole) or "").strip()
            new_val = item.text().strip()
            if new_val == old_val:
                return
            if not self._confirm_officer_change(
                "Подтверждение изменения участка",
                officer,
                "Участок",
                old_val,
                new_val,
            ):
                self._suspend_off_table_item_handler = True
                item.setText(old_val)
                self._suspend_off_table_item_handler = False
                return
            db.update_officer_field(officer['id'], 'district', new_val)
            officer['district'] = new_val
            item.setData(Qt.ItemDataRole.UserRole, new_val)
            self._notify_officers_changed()

        elif col == self.C_FIO:
            old_val = str(item.data(Qt.ItemDataRole.UserRole) or "").strip()
            new_val = item.text().strip()
            if new_val == old_val:
                return
            if not new_val:
                QMessageBox.warning(self, "Пустое ФИО", "ФИО не может быть пустым.")
                self._suspend_off_table_item_handler = True
                item.setText(old_val)
                self._suspend_off_table_item_handler = False
                return
            if not self._confirm_officer_change(
                "Подтверждение изменения ФИО",
                officer,
                "ФИО",
                old_val,
                new_val,
            ):
                self._suspend_off_table_item_handler = True
                item.setText(old_val)
                self._suspend_off_table_item_handler = False
                return
            db.update_officer_field(officer['id'], 'fio', new_val)
            officer['fio'] = new_val
            item.setData(Qt.ItemDataRole.UserRole, new_val)
            self._notify_officers_changed()

    def _on_rank_changed(self, officer_id: int, row: int, value: str):
        if self._suspend_officer_combo_handlers or row < 0 or row >= len(self._officers):
            return
        cb = self._off_table.cellWidget(row, self.C_RANK)
        if not isinstance(cb, QComboBox):
            return
        old_val = str(cb.property('prev_value') or "").strip()
        new_val = str(value or "").strip()
        if not new_val:
            self._suspend_officer_combo_handlers = True
            cb.setCurrentText(old_val)
            self._suspend_officer_combo_handlers = False
            return
        if new_val == old_val:
            return
        officer = self._officers[row]
        if not self._confirm_officer_change(
            "Подтверждение изменения звания",
            officer,
            "Звание",
            old_val,
            new_val,
        ):
            self._suspend_officer_combo_handlers = True
            cb.setCurrentText(old_val)
            self._suspend_officer_combo_handlers = False
            return
        db.update_officer_field(officer_id, 'rank', new_val)
        self._officers[row]['rank'] = new_val
        cb.setProperty('prev_value', new_val)
        self._notify_officers_changed(refresh_unmatched=False)

    def _on_position_changed(self, officer_id: int, row: int, value: str):
        if self._suspend_officer_combo_handlers or row < 0 or row >= len(self._officers):
            return
        cb = self._off_table.cellWidget(row, self.C_POS)
        if not isinstance(cb, QComboBox):
            return
        old_val = str(cb.property('prev_value') or "").strip()
        new_val = str(value or "").strip()
        if not new_val:
            self._suspend_officer_combo_handlers = True
            cb.setCurrentText(old_val)
            self._suspend_officer_combo_handlers = False
            return
        if new_val == old_val:
            return
        officer = self._officers[row]
        if not self._confirm_officer_change(
            "Подтверждение изменения должности",
            officer,
            "Должность",
            old_val,
            new_val,
        ):
            self._suspend_officer_combo_handlers = True
            cb.setCurrentText(old_val)
            self._suspend_officer_combo_handlers = False
            return
        db.update_officer_field(officer_id, 'position', new_val)
        self._officers[row]['position'] = new_val
        cb.setProperty('prev_value', new_val)
        self._notify_officers_changed(refresh_unmatched=False)

    def _on_replacement_changed(self, officer_id: int, cb: QComboBox, row: int = None):
        if self._suspend_officer_combo_handlers or not isinstance(cb, QComboBox):
            return
        if row is None or row < 0 or row >= len(self._officers):
            row = next((idx for idx, o in enumerate(self._officers) if o.get('id') == officer_id), -1)
        if row < 0 or row >= len(self._officers):
            return
        officer = self._officers[row]
        old_repl_id = cb.property('prev_repl_id')
        repl_id = cb.currentData()

        old_lbl = self._repl_label(cb, old_repl_id)
        new_lbl = self._repl_label(cb, repl_id)
        if old_lbl == new_lbl:
            return
        if not self._confirm_officer_change(
            "Подтверждение изменения замещения",
            officer,
            "Замещает",
            old_lbl,
            new_lbl,
        ):
            self._suspend_officer_combo_handlers = True
            idx_old = cb.findData(old_repl_id)
            cb.setCurrentIndex(idx_old if idx_old >= 0 else 0)
            self._suspend_officer_combo_handlers = False
            return

        if repl_id is None:
            self._replacements.pop(officer_id, None)
            db.set_officer_replacement(officer_id, None)
        else:
            self._replacements[officer_id] = int(repl_id)
            db.set_officer_replacement(officer_id, int(repl_id))
        cb.setProperty('prev_repl_id', repl_id)
        self._notify_officers_changed(refresh_unmatched=False)

    # ── Адреса участкового ────────────────────────────────────────────────────

    def _open_addr_dialog(self, row: int):
        if row >= len(self._officers):
            return
        officer = self._officers[row]
        dlg = _AddrDialog(officer, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        old_addrs = str(officer.get('addresses', '') or '').strip()
        new_addrs = str(dlg.get_text() or '').strip()
        if new_addrs == old_addrs:
            return

        old_preview = old_addrs.replace('\n', '; ')[:220] if old_addrs else "—"
        new_preview = new_addrs.replace('\n', '; ')[:220] if new_addrs else "—"
        if len(old_addrs) > 220:
            old_preview += "..."
        if len(new_addrs) > 220:
            new_preview += "..."
        fio = str(officer.get('fio', '') or '—')

        r = QMessageBox.question(
            self,
            "Подтверждение изменения адресов",
            (
                f"Изменить адреса для:\n{fio}\n\n"
                f"Было:\n{old_preview}\n\n"
                f"Станет:\n{new_preview}"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return

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
            key_fio  = str(rec.get('_source_fio', fio_val) or '').strip()
            key_dob  = str(rec.get('_source_dob', dob_val) or '').strip()
            addr_val = _normalize_person_address(rec.get('Место жительства', ''))

            for col, val in ((0, fio_val), (1, dob_val), (2, addr_val)):
                item = QTableWidgetItem(val)
                item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                if col == 2:
                    item.setToolTip(addr_val)
                self._um_table.setItem(i, col, item)

            off_cb = QComboBox()
            off_cb.addItems(off_labels)

            assigned_id = db.get_assignment(key_fio, key_dob)
            if assigned_id and assigned_id in officers_by_id:
                label = g._officer_label(officers_by_id[assigned_id])
                if label in off_labels:
                    off_cb.setCurrentText(label)

            off_cb.setProperty('record_fio', key_fio)
            off_cb.setProperty('record_dob', key_dob)
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
        self._dark = _to_bool(settings.get('dark_theme', False), False)
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
        self._dark = bool(dark)
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet("")
            _apply_palette(app, self._dark)
            # Win10/Qt6: локальные styleSheet иногда не пересчитывают palette(...) автоматически.
            # Принудительно переустанавливаем стили и переполировываем виджеты.
            for w in app.allWidgets():
                ss = w.styleSheet()
                if ss:
                    w.setStyleSheet("")
                    w.setStyleSheet(ss)
                style = w.style()
                if style is not None:
                    style.unpolish(w)
                    style.polish(w)
                w.update()
        self._save_app_settings()

    def _save_app_settings(self):
        s = _load_settings()
        s['dark_theme'] = self._dark
        s['char_templates'] = _norm_char_templates(self._char_templates)
        _save_settings(s)

    def _open_settings(self):
        officers = self.officers_tab.get_active_officers() if hasattr(self, "officers_tab") else []
        officer_templates = db.all_officer_char_templates()
        dlg = _SettingsDialog(
            self._dark,
            self._char_templates,
            officers=officers,
            officer_char_templates=officer_templates,
            parent=self,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._char_templates = dlg.get_templates()
        _apply_char_templates(self._char_templates)
        db.replace_officer_char_templates(dlg.get_officer_templates())
        self._set_theme(dlg.is_dark_theme())


# ── Точка входа ───────────────────────────────────────────────────────────────

def main():
    if sys.platform == "win32":
        os.system("chcp 65001 > nul")
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("AutoObzorki.App.3.1")
        except Exception:
            pass

    if hasattr(QApplication, "setHighDpiScaleFactorRoundingPolicy"):
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.RoundPreferFloor
        )
    app = QApplication(sys.argv)
    app_icon = _load_app_icon()
    if app_icon is not None:
        app.setWindowIcon(app_icon)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 9))
    settings = _load_settings()
    _apply_char_templates(_get_char_templates(settings))
    app.setStyleSheet("")
    _apply_palette(app, _to_bool(settings.get('dark_theme', False), False))

    splash = _show_startup_splash(app, app_icon=app_icon)
    started_at = pytime.monotonic()

    win = MainWindow()
    if app_icon is not None:
        win.setWindowIcon(app_icon)
    _finish_startup_splash(app, splash, started_at, minimum_seconds=1.1)
    win.show()
    splash.finish(win)
    if hasattr(win, "obzorki_tab") and hasattr(win.obzorki_tab, "run_startup_autoload"):
        QTimer.singleShot(0, win.obzorki_tab.run_startup_autoload)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
