"""
db.py — SQLite база данных для хранения участковых и назначений.
"""

import sqlite3
import os
import re
import json

_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uii_data.db')


def _conn() -> sqlite3.Connection:
    c = sqlite3.connect(_DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def _normalize_address_line(text: str) -> str:
    if not text:
        return ''
    return re.sub(r'\s+', ' ', str(text).replace('\r', ' ').replace('\n', ' ')).strip()


def _normalize_text_line(text: str) -> str:
    if text is None:
        return ''
    return re.sub(r'\s+', ' ', str(text).replace('\r', ' ').replace('\n', ' ')).strip()


def init_db():
    """Создаёт/обновляет таблицы."""
    with _conn() as c:
        c.execute('''
            CREATE TABLE IF NOT EXISTS officers (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                fio         TEXT    NOT NULL DEFAULT '',
                rank        TEXT    NOT NULL DEFAULT '',
                position    TEXT    NOT NULL DEFAULT '',
                upp         TEXT    NOT NULL DEFAULT '',
                district    TEXT    NOT NULL DEFAULT '',
                addresses   TEXT    NOT NULL DEFAULT '',
                is_vacancy  INTEGER NOT NULL DEFAULT 0,
                source_file TEXT    NOT NULL DEFAULT '',
                do_generate INTEGER NOT NULL DEFAULT 1
            )
        ''')
        # Совместимость: добавить колонку если старая БД
        try:
            c.execute('ALTER TABLE officers ADD COLUMN do_generate INTEGER NOT NULL DEFAULT 1')
        except Exception:
            pass

        c.execute('''
            CREATE TABLE IF NOT EXISTS person_assignments (
                fio        TEXT NOT NULL,
                dob        TEXT NOT NULL,
                officer_id INTEGER,
                PRIMARY KEY (fio, dob)
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS person_characteristics (
                fio         TEXT NOT NULL,
                dob         TEXT NOT NULL,
                char_type   TEXT NOT NULL DEFAULT 'нейтральная',
                custom_text TEXT NOT NULL DEFAULT '',
                PRIMARY KEY (fio, dob)
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS person_address_fixes (
                fio   TEXT NOT NULL,
                dob   TEXT NOT NULL,
                addr  TEXT NOT NULL DEFAULT '',
                PRIMARY KEY (fio, dob)
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS person_overrides (
                fio          TEXT NOT NULL,
                dob          TEXT NOT NULL,
                fio_override TEXT NOT NULL DEFAULT '',
                PRIMARY KEY (fio, dob)
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS person_doc_overrides (
                fio       TEXT NOT NULL,
                dob       TEXT NOT NULL,
                data_json TEXT NOT NULL DEFAULT '{}',
                PRIMARY KEY (fio, dob)
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS officer_replacements (
                officer_id             INTEGER PRIMARY KEY,
                replacement_officer_id INTEGER NOT NULL
            )
        ''')

        c.execute('''
            CREATE TABLE IF NOT EXISTS officer_char_templates (
                officer_id    INTEGER NOT NULL,
                char_type     TEXT    NOT NULL,
                sort_order    INTEGER NOT NULL DEFAULT 0,
                template_text TEXT    NOT NULL DEFAULT '',
                PRIMARY KEY (officer_id, char_type, sort_order)
            )
        ''')


# ── Officers ───────────────────────────────────────────────────────────────────

def save_officers(rows: list):
    """Перезаписывает список сотрудников. Ручные назначения сохраняются через (fio, district)."""
    with _conn() as c:
        # Сохраняем маппинг (fio, district) → old_id
        old_rows = c.execute('SELECT id, fio, district FROM officers').fetchall()
        old_ids = {(r['fio'], r['district']): r['id'] for r in old_rows}
        old_key_by_id = {r['id']: (r['fio'], r['district']) for r in old_rows}
        old_replacements = {
            r['officer_id']: r['replacement_officer_id']
            for r in c.execute(
                'SELECT officer_id, replacement_officer_id FROM officer_replacements'
            ).fetchall()
        }
        old_template_rows = c.execute(
            'SELECT officer_id, char_type, sort_order, template_text '
            'FROM officer_char_templates '
            'ORDER BY officer_id, char_type, sort_order'
        ).fetchall()

        c.execute('DELETE FROM officers')
        c.executemany(
            'INSERT INTO officers '
            '(fio, rank, position, upp, district, addresses, is_vacancy, source_file) '
            'VALUES (:fio, :rank, :position, :upp, :district, :addresses, :is_vacancy, :source_file)',
            rows,
        )

        # Строим маппинг (fio, district) → new_id
        new_ids = {(r['fio'], r['district']): r['id']
                   for r in c.execute('SELECT id, fio, district FROM officers').fetchall()}

        # Перепривязываем person_assignments: old_id → new_id
        for key, old_id in old_ids.items():
            new_id = new_ids.get(key)
            if new_id and new_id != old_id:
                c.execute('UPDATE person_assignments SET officer_id=? WHERE officer_id=?',
                          (new_id, old_id))

        # Удаляем назначения на уже несуществующих сотрудников
        valid_ids = list(new_ids.values())
        if valid_ids:
            c.execute(
                f'DELETE FROM person_assignments '
                f'WHERE officer_id NOT IN ({",".join("?" * len(valid_ids))})',
                valid_ids)
        else:
            c.execute('DELETE FROM person_assignments')

        # Перепривязываем замещения old_id -> new_id
        c.execute('DELETE FROM officer_replacements')
        repl_rows = []
        for old_off_id, old_rep_id in old_replacements.items():
            off_key = old_key_by_id.get(old_off_id)
            rep_key = old_key_by_id.get(old_rep_id)
            if not off_key or not rep_key:
                continue
            new_off_id = new_ids.get(off_key)
            new_rep_id = new_ids.get(rep_key)
            if not new_off_id or not new_rep_id or new_off_id == new_rep_id:
                continue
            repl_rows.append((new_off_id, new_rep_id))
        if repl_rows:
            c.executemany(
                'INSERT OR REPLACE INTO officer_replacements (officer_id, replacement_officer_id) VALUES (?,?)',
                repl_rows
            )

        # Перепривязываем персональные шаблоны old_id -> new_id
        remapped_templates = {}  # {officer_id: {char_type: [text, ...]}}
        for row in old_template_rows:
            off_key = old_key_by_id.get(row['officer_id'])
            if not off_key:
                continue
            new_off_id = new_ids.get(off_key)
            if not new_off_id:
                continue
            char_type = (row['char_type'] or '').strip()
            text = (row['template_text'] or '').strip()
            if not char_type or not text:
                continue
            remapped_templates.setdefault(new_off_id, {}).setdefault(char_type, []).append(text)

        c.execute('DELETE FROM officer_char_templates')
        insert_rows = []
        for off_id, by_type in remapped_templates.items():
            for char_type, texts in by_type.items():
                uniq_texts = []
                seen = set()
                for text in texts:
                    if text in seen:
                        continue
                    seen.add(text)
                    uniq_texts.append(text)
                for idx, text in enumerate(uniq_texts):
                    insert_rows.append((off_id, char_type, idx, text))
        if insert_rows:
            c.executemany(
                'INSERT OR REPLACE INTO officer_char_templates '
                '(officer_id, char_type, sort_order, template_text) VALUES (?,?,?,?)',
                insert_rows,
            )


def all_officers() -> list:
    with _conn() as c:
        # Сортировка по ФИО (фамилии), затем по участку для стабильного порядка.
        return [dict(r) for r in c.execute(
            'SELECT * FROM officers ORDER BY fio COLLATE NOCASE, CAST(district AS INTEGER), district'
        ).fetchall()]


def officers_count() -> int:
    try:
        with _conn() as c:
            return c.execute('SELECT COUNT(*) FROM officers').fetchone()[0]
    except Exception:
        return 0


def set_officer_generate(officer_id: int, generate: bool):
    with _conn() as c:
        c.execute('UPDATE officers SET do_generate=? WHERE id=?',
                  (1 if generate else 0, officer_id))


def all_officer_replacements() -> dict:
    """Возвращает {officer_id: replacement_officer_id}."""
    try:
        with _conn() as c:
            rows = c.execute(
                'SELECT officer_id, replacement_officer_id FROM officer_replacements'
            ).fetchall()
            return {
                int(r['officer_id']): int(r['replacement_officer_id'])
                for r in rows
            }
    except Exception:
        return {}


def set_officer_replacement(officer_id: int, replacement_officer_id):
    """Сохраняет замещение участкового. None удаляет замещение."""
    with _conn() as c:
        if replacement_officer_id is None or replacement_officer_id == officer_id:
            c.execute('DELETE FROM officer_replacements WHERE officer_id=?', (officer_id,))
            return
        c.execute(
            'INSERT OR REPLACE INTO officer_replacements (officer_id, replacement_officer_id) VALUES (?,?)',
            (int(officer_id), int(replacement_officer_id))
        )


# ── Person assignments ─────────────────────────────────────────────────────────

def get_assignment(fio: str, dob: str):
    """officer_id для подучётного или None."""
    try:
        with _conn() as c:
            row = c.execute(
                'SELECT officer_id FROM person_assignments WHERE fio=? AND dob=?',
                (fio, dob)).fetchone()
            return row['officer_id'] if row else None
    except Exception:
        return None


def set_assignment(fio: str, dob: str, officer_id):
    """Сохраняет ручное назначение участкового."""
    with _conn() as c:
        if officer_id is None:
            c.execute('DELETE FROM person_assignments WHERE fio=? AND dob=?', (fio, dob))
        else:
            c.execute(
                'INSERT OR REPLACE INTO person_assignments (fio, dob, officer_id) VALUES (?,?,?)',
                (fio, dob, officer_id))


def all_person_fio_overrides() -> dict:
    """Возвращает {(fio, dob): fio_override}."""
    try:
        with _conn() as c:
            rows = c.execute(
                'SELECT fio, dob, fio_override FROM person_overrides'
            ).fetchall()
            return {
                ((r['fio'] or '').strip(), (r['dob'] or '').strip()): (r['fio_override'] or '').strip()
                for r in rows
                if (r['fio_override'] or '').strip()
            }
    except Exception:
        return {}


def set_person_fio_override(fio: str, dob: str, fio_override: str):
    """Сохраняет пользовательское ФИО (override) для исходного ключа (fio,dob)."""
    fio = (fio or '').strip()
    dob = (dob or '').strip()
    fio_override = " ".join((fio_override or "").split()).strip()
    if not fio:
        return
    with _conn() as c:
        if not fio_override or fio_override == fio:
            c.execute('DELETE FROM person_overrides WHERE fio=? AND dob=?', (fio, dob))
            return
        c.execute(
            'INSERT OR REPLACE INTO person_overrides (fio, dob, fio_override) VALUES (?,?,?)',
            (fio, dob, fio_override),
        )


_DOC_OVERRIDE_KEYS = frozenset({
    'dob',
    'court',
    'duties',
    'end_date',
    'work_place',
    'phone',
    'links',
    'features',
    'season_clothes',
    'violations',
    'ic_check',
})


def all_person_doc_overrides() -> dict:
    """Возвращает {(fio, dob): {key: value, ...}} для персональных правок карточки."""
    try:
        with _conn() as c:
            rows = c.execute(
                'SELECT fio, dob, data_json FROM person_doc_overrides'
            ).fetchall()
            result = {}
            for r in rows:
                key = ((r['fio'] or '').strip(), (r['dob'] or '').strip())
                raw = (r['data_json'] or '').strip()
                if not key[0] or not raw:
                    continue
                try:
                    data = json.loads(raw)
                except Exception:
                    continue
                if not isinstance(data, dict):
                    continue
                cleaned = {}
                for k, v in data.items():
                    if k not in _DOC_OVERRIDE_KEYS:
                        continue
                    cleaned[k] = _normalize_text_line(v)
                if cleaned:
                    result[key] = cleaned
            return result
    except Exception:
        return {}


def set_person_doc_overrides(fio: str, dob: str, data: dict):
    """Сохраняет персональные правки карточки подучётного.
    data: {'dob','court','duties','end_date','work_place','phone'}.
    Пустой dict удаляет запись.
    """
    fio = (fio or '').strip()
    dob = (dob or '').strip()
    if not fio:
        return

    cleaned = {}
    if isinstance(data, dict):
        for k in _DOC_OVERRIDE_KEYS:
            if k in data:
                cleaned[k] = _normalize_text_line(data.get(k))
        cleaned = {k: v for k, v in cleaned.items() if k in _DOC_OVERRIDE_KEYS}

    with _conn() as c:
        if not cleaned:
            c.execute('DELETE FROM person_doc_overrides WHERE fio=? AND dob=?', (fio, dob))
            return
        c.execute(
            'INSERT OR REPLACE INTO person_doc_overrides (fio, dob, data_json) VALUES (?,?,?)',
            (fio, dob, json.dumps(cleaned, ensure_ascii=False)),
        )


def rename_person_key(old_fio: str, old_dob: str, new_fio: str, new_dob: str = None) -> bool:
    """Переименовывает ключ персоны (fio, dob) во всех таблицах персональных данных.
    Возвращает True при успехе, False если изменение невозможно (например, конфликт ключа).
    """
    old_fio = (old_fio or '').strip()
    old_dob = (old_dob or '').strip()
    new_fio = (new_fio or '').strip()
    new_dob = old_dob if new_dob is None else (new_dob or '').strip()

    if not old_fio or not new_fio:
        return False
    if old_fio == new_fio and old_dob == new_dob:
        return True

    tables = ('person_assignments', 'person_characteristics', 'person_address_fixes', 'person_doc_overrides')

    try:
        with _conn() as c:
            for table in tables:
                has_old = c.execute(
                    f'SELECT 1 FROM {table} WHERE fio=? AND dob=?',
                    (old_fio, old_dob),
                ).fetchone()
                if not has_old:
                    continue
                has_new = c.execute(
                    f'SELECT 1 FROM {table} WHERE fio=? AND dob=?',
                    (new_fio, new_dob),
                ).fetchone()
                if has_new:
                    return False

            for table in tables:
                c.execute(
                    f'UPDATE {table} SET fio=?, dob=? WHERE fio=? AND dob=?',
                    (new_fio, new_dob, old_fio, old_dob),
                )
        return True
    except Exception:
        return False


def update_officer_addresses(officer_id: int, addresses: str):
    with _conn() as c:
        c.execute('UPDATE officers SET addresses=? WHERE id=?', (addresses, officer_id))


_ALLOWED_OFFICER_FIELDS = frozenset({'rank', 'position', 'district', 'fio', 'upp'})


def update_officer_field(officer_id: int, field: str, value: str):
    """Обновляет одно поле сотрудника. field должен быть из _ALLOWED_OFFICER_FIELDS."""
    if field not in _ALLOWED_OFFICER_FIELDS:
        raise ValueError(f'Поле {field!r} нельзя обновлять через эту функцию')
    with _conn() as c:
        c.execute(f'UPDATE officers SET {field}=? WHERE id=?', (value, officer_id))


def all_assignments() -> dict:
    """Возвращает {(fio, dob): officer_id}."""
    try:
        with _conn() as c:
            rows = c.execute(
                'SELECT fio, dob, officer_id FROM person_assignments').fetchall()
            return {(r['fio'], r['dob']): r['officer_id'] for r in rows}
    except Exception:
        return {}


# ── Officer characteristic templates ───────────────────────────────────────────

def all_officer_char_templates() -> dict:
    """Возвращает {officer_id: {char_type: [template, ...]}}."""
    try:
        with _conn() as c:
            rows = c.execute(
                'SELECT officer_id, char_type, template_text '
                'FROM officer_char_templates '
                'ORDER BY officer_id, char_type, sort_order'
            ).fetchall()
            result = {}
            for r in rows:
                off_id = int(r['officer_id'])
                char_type = (r['char_type'] or '').strip()
                text = (r['template_text'] or '').strip()
                if not char_type or not text:
                    continue
                result.setdefault(off_id, {}).setdefault(char_type, []).append(text)
            return result
    except Exception:
        return {}


def replace_officer_char_templates(data: dict):
    """Полностью заменяет персональные шаблоны участковых.
    Формат: {officer_id: {char_type: [template, ...]}}.
    """
    rows = []
    if isinstance(data, dict):
        for officer_id, by_type in data.items():
            try:
                off_id = int(officer_id)
            except Exception:
                continue
            if not isinstance(by_type, dict):
                continue
            for char_type, raw in by_type.items():
                ctype = (char_type or '').strip()
                if not ctype:
                    continue
                values = []
                if isinstance(raw, str):
                    values = [raw]
                elif isinstance(raw, (list, tuple)):
                    values = [x for x in raw if isinstance(x, str)]
                cleaned = []
                seen = set()
                for text in values:
                    txt = text.strip()
                    if not txt or txt in seen:
                        continue
                    seen.add(txt)
                    cleaned.append(txt)
                for idx, txt in enumerate(cleaned):
                    rows.append((off_id, ctype, idx, txt))

    with _conn() as c:
        c.execute('DELETE FROM officer_char_templates')
        if rows:
            c.executemany(
                'INSERT OR REPLACE INTO officer_char_templates '
                '(officer_id, char_type, sort_order, template_text) VALUES (?,?,?,?)',
                rows,
            )


# ── Person characteristics ──────────────────────────────────────────────────────

def all_person_characteristics() -> dict:
    """Возвращает {(fio, dob): {'char_type': str, 'custom_text': str}}."""
    try:
        with _conn() as c:
            rows = c.execute(
                'SELECT fio, dob, char_type, custom_text FROM person_characteristics'
            ).fetchall()
            return {
                (r['fio'], r['dob']): {
                    'char_type': r['char_type'] or 'нейтральная',
                    'custom_text': r['custom_text'] or '',
                }
                for r in rows
            }
    except Exception:
        return {}


def set_person_characteristic(fio: str, dob: str, char_type: str, custom_text: str = ''):
    """Сохраняет выбор характеристики для конкретного человека.
    Если выбор дефолтный (нейтральная без custom), запись удаляется.
    """
    fio = (fio or '').strip()
    dob = (dob or '').strip()
    char_type = (char_type or '').strip() or 'нейтральная'
    custom_text = (custom_text or '').strip()
    if not fio:
        return
    with _conn() as c:
        if char_type == 'нейтральная' and not custom_text:
            c.execute('DELETE FROM person_characteristics WHERE fio=? AND dob=?', (fio, dob))
            return
        c.execute(
            'INSERT OR REPLACE INTO person_characteristics (fio, dob, char_type, custom_text) '
            'VALUES (?,?,?,?)',
            (fio, dob, char_type, custom_text)
        )


def get_person_characteristic(fio: str, dob: str):
    """Возвращает {'char_type': ..., 'custom_text': ...} или None."""
    fio = (fio or '').strip()
    dob = (dob or '').strip()
    if not fio:
        return None
    try:
        with _conn() as c:
            row = c.execute(
                'SELECT char_type, custom_text FROM person_characteristics WHERE fio=? AND dob=?',
                (fio, dob)
            ).fetchone()
            if not row:
                return None
            return {
                'char_type': row['char_type'] or 'нейтральная',
                'custom_text': row['custom_text'] or '',
            }
    except Exception:
        return None


def list_missing_person_characteristics(valid_keys: set) -> list:
    """Список сохранённых записей, которых нет в текущем наборе valid_keys={(fio,dob)}."""
    valid = {(str(f or '').strip(), str(d or '').strip()) for (f, d) in (valid_keys or set())}
    try:
        with _conn() as c:
            rows = c.execute(
                'SELECT fio, dob, char_type, custom_text FROM person_characteristics'
            ).fetchall()
            result = []
            for r in rows:
                key = ((r['fio'] or '').strip(), (r['dob'] or '').strip())
                if key not in valid:
                    result.append({
                        'fio': key[0],
                        'dob': key[1],
                        'char_type': r['char_type'] or 'нейтральная',
                        'custom_text': r['custom_text'] or '',
                    })
            return result
    except Exception:
        return []


def delete_person_characteristics(keys: list):
    """Удаляет сохранённые характеристики для списка ключей [(fio,dob), ...]."""
    if not keys:
        return
    with _conn() as c:
        c.executemany(
            'DELETE FROM person_characteristics WHERE fio=? AND dob=?',
            [((fio or '').strip(), (dob or '').strip()) for fio, dob in keys]
        )


# ── Person address fixes ────────────────────────────────────────────────────────

def all_person_address_fixes() -> dict:
    """Возвращает {(fio, dob): addr}."""
    try:
        with _conn() as c:
            rows = c.execute('SELECT fio, dob, addr FROM person_address_fixes').fetchall()
            return {
                ((r['fio'] or '').strip(), (r['dob'] or '').strip()): _normalize_address_line(r['addr'] or '')
                for r in rows
            }
    except Exception:
        return {}


def set_person_address_fix(fio: str, dob: str, addr: str):
    """Сохраняет исправленный адрес для конкретного человека."""
    fio = (fio or '').strip()
    dob = (dob or '').strip()
    addr = _normalize_address_line(addr or '')
    if not fio:
        return
    with _conn() as c:
        if not addr:
            c.execute('DELETE FROM person_address_fixes WHERE fio=? AND dob=?', (fio, dob))
            return
        c.execute(
            'INSERT OR REPLACE INTO person_address_fixes (fio, dob, addr) VALUES (?,?,?)',
            (fio, dob, addr)
        )
