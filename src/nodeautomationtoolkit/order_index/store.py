"""Локальний індекс посад: SQLite-файл у папці, яку обрав користувач.

Повторне індексування пропускає файли, що не змінилися (розмір і час зміни),
і прибирає з індексу файли, яких у папці вже немає.
"""

from __future__ import annotations

import os
import sqlite3
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from .extractor import order_metadata, position_key
from .items import OrderItem, iter_items, person_key
from .subdivisions import chain_key, split_chain
from .position_dictionary import load_position_dictionary
from .unit_split import load_markers, split_position_unit
from .word_reader import WordTextReader, WordUnavailable

INDEX_FILENAME = "order_positions_index.sqlite"
#: Версія структури індексу; старіший файл перебудовується з нуля.
SCHEMA_VERSION = 5

#: Word-формати, які бачить сканер: .docx читається напряму, .doc/.rtf — через Word.
DOCX_SUFFIXES = {".docx", ".docm"}
LEGACY_SUFFIXES = {".doc", ".rtf"}

_SCHEMA = """
CREATE TABLE IF NOT EXISTS files (
    id INTEGER PRIMARY KEY,
    path TEXT NOT NULL UNIQUE,
    size INTEGER NOT NULL,
    mtime REAL NOT NULL,
    number TEXT NOT NULL DEFAULT '',
    date TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL,
    error TEXT NOT NULL DEFAULT '',
    hits INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS positions (
    id INTEGER PRIMARY KEY,
    key TEXT NOT NULL UNIQUE,
    display TEXT NOT NULL,
    known INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY,
    file_id INTEGER NOT NULL REFERENCES files(id) ON DELETE CASCADE,
    section TEXT NOT NULL DEFAULT '',
    label TEXT NOT NULL DEFAULT '',
    paragraph INTEGER NOT NULL DEFAULT 0,
    rank TEXT NOT NULL DEFAULT '',
    surname TEXT NOT NULL DEFAULT '',
    name TEXT NOT NULL DEFAULT '',
    patronymic TEXT NOT NULL DEFAULT '',
    person TEXT NOT NULL DEFAULT '',
    ipn TEXT NOT NULL DEFAULT '',
    current_position TEXT NOT NULL DEFAULT '',
    target_position TEXT NOT NULL DEFAULT '',
    text TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS items_ipn ON items(ipn);
CREATE INDEX IF NOT EXISTS items_person ON items(person);
CREATE INDEX IF NOT EXISTS items_file ON items(file_id);
CREATE TABLE IF NOT EXISTS occurrences (
    position_id INTEGER NOT NULL REFERENCES positions(id),
    file_id INTEGER NOT NULL REFERENCES files(id) ON DELETE CASCADE,
    item_id INTEGER REFERENCES items(id) ON DELETE CASCADE,
    role TEXT NOT NULL,
    unit TEXT NOT NULL DEFAULT '',
    unit_key TEXT NOT NULL DEFAULT '',
    full_text TEXT NOT NULL DEFAULT '',
    section TEXT NOT NULL DEFAULT '',
    paragraph INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS subdivisions (
    id INTEGER PRIMARY KEY,
    key TEXT NOT NULL UNIQUE,
    display TEXT NOT NULL,
    kind TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS item_subdivisions (
    item_id INTEGER NOT NULL REFERENCES items(id) ON DELETE CASCADE,
    subdivision_id INTEGER NOT NULL REFERENCES subdivisions(id),
    file_id INTEGER NOT NULL REFERENCES files(id) ON DELETE CASCADE,
    role TEXT NOT NULL DEFAULT '',
    level INTEGER NOT NULL DEFAULT 0,
    text TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS item_subdivisions_subdivision ON item_subdivisions(subdivision_id);
CREATE INDEX IF NOT EXISTS item_subdivisions_file ON item_subdivisions(file_id);
CREATE INDEX IF NOT EXISTS occurrences_position ON occurrences(position_id);
CREATE INDEX IF NOT EXISTS occurrences_file ON occurrences(file_id);
"""


@dataclass
class IndexSummary:
    scanned: int = 0
    indexed: int = 0
    items: int = 0
    unchanged: int = 0
    removed: int = 0
    skipped_legacy: int = 0
    failed: int = 0
    hits: int = 0
    cancelled: bool = False


def index_path(index_folder: str | Path) -> Path:
    return Path(index_folder) / INDEX_FILENAME


def connect(index_folder: str | Path) -> sqlite3.Connection:
    folder = Path(index_folder)
    folder.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(index_path(folder))
    connection.execute("PRAGMA foreign_keys = ON")
    if connection.execute("PRAGMA user_version").fetchone()[0] != SCHEMA_VERSION:
        connection.executescript(
            "DROP TABLE IF EXISTS occurrences; DROP TABLE IF EXISTS positions; DROP TABLE IF EXISTS files;"
        )
        connection.executescript(_SCHEMA)
        connection.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        connection.commit()
    return connection


def iter_order_files(orders_folder: str | Path, recursive: bool = True):
    root = Path(orders_folder)
    walker = os.walk(root) if recursive else [(str(root), [], os.listdir(root))]
    for folder, _dirs, names in walker:
        for name in sorted(names):
            if name.startswith("~$"):
                continue  # тимчасовий файл відкритого документа Word
            path = Path(folder) / name
            if path.suffix.lower() in DOCX_SUFFIXES | LEGACY_SUFFIXES and path.is_file():
                yield path


def build_index(
    orders_folder: str | Path,
    index_folder: str | Path,
    *,
    recursive: bool = True,
    force: bool = False,
    log: Callable[[str], None] = lambda _message: None,
    progress: Callable[[int, int, str], None] = lambda _done, _total, _name: None,
    cancelled: Callable[[], bool] = lambda: False,
) -> IndexSummary:
    summary = IndexSummary()
    files = list(iter_order_files(orders_folder, recursive))
    total = len(files)
    log(f"Знайдено файлів Word: {total}")
    connection = connect(index_folder)
    context = _IndexContext(load_position_dictionary(), load_markers(), WordTextReader(), log)
    try:
        if force:
            connection.executescript(
                "DELETE FROM occurrences; DELETE FROM items; DELETE FROM files; DELETE FROM positions;"
            )

        seen: set[str] = set()
        for done, path in enumerate(files, start=1):
            if cancelled():
                summary.cancelled = True
                log("⚠ Індексування перервано користувачем")
                break
            progress(done - 1, total, path.name)
            summary.scanned += 1
            key = str(path.resolve())
            seen.add(key)
            _index_one(connection, path, key, summary, context)
            if done % 25 == 0:
                connection.commit()
        progress(total, total, "")

        if not summary.cancelled:
            summary.removed = _remove_missing(connection, seen, Path(orders_folder).resolve(), recursive)
        _drop_orphan_positions(connection)
        connection.commit()
    finally:
        context.word.close()
        connection.close()
    return summary


@dataclass
class _IndexContext:
    dictionary: object
    markers: tuple
    word: WordTextReader
    log: Callable[[str], None]
    word_unavailable: str = ""


def _index_one(
    connection: sqlite3.Connection, path: Path, key: str, summary: IndexSummary, context: _IndexContext
) -> None:
    log = context.log
    stat = path.stat()
    row = connection.execute("SELECT id, size, mtime, status FROM files WHERE path = ?", (key,)).fetchone()
    if row and row[1] == stat.st_size and abs(row[2] - stat.st_mtime) < 1e-6:
        summary.unchanged += 1
        if row[3] == "legacy":
            summary.skipped_legacy += 1
        return
    if row:
        _forget_file(connection, row[0])

    number, date = order_metadata(path.stem)
    if context.word_unavailable and path.suffix.lower() in LEGACY_SUFFIXES:
        summary.skipped_legacy += 1
        connection.execute(
            "INSERT INTO files(path, size, mtime, number, date, status, error) VALUES (?, ?, ?, ?, ?, 'legacy', ?)",
            (key, stat.st_size, stat.st_mtime, number, date, context.word_unavailable),
        )
        return

    try:
        paragraphs = context.word.read_paragraphs(path)
    except WordUnavailable as error:
        context.word_unavailable = str(error)
        log(f"⚠ Увага: .doc/.rtf не читаються — {error}")
        _index_one(connection, path, key, summary, context)
        return
    except Exception as error:  # пошкоджений або захищений паролем файл
        summary.failed += 1
        log(f"❌ Не вдалося прочитати {path.name}: {type(error).__name__}")
        connection.execute(
            "INSERT INTO files(path, size, mtime, number, date, status, error) VALUES (?, ?, ?, ?, ?, 'error', ?)",
            (key, stat.st_size, stat.st_mtime, number, date, f"{type(error).__name__}: {error}"),
        )
        return

    if not (number and date):
        number, date = order_metadata(path.stem, "\n".join(paragraphs[:40]))
    items = iter_items(paragraphs, context.dictionary)
    hits_count = sum(len(item.hits) for item in items)
    cursor = connection.execute(
        "INSERT INTO files(path, size, mtime, number, date, status, hits) VALUES (?, ?, ?, ?, ?, 'ok', ?)",
        (key, stat.st_size, stat.st_mtime, number, date, hits_count),
    )
    file_id = cursor.lastrowid
    for item in items:
        item_id = _save_item(connection, file_id, item)
        for hit in item.hits:
            unit = split_position_unit(hit.text, context.markers)[1]
            position_id = _position_id(connection, hit.position, hit.known)
            connection.execute(
                "INSERT INTO occurrences(position_id, file_id, item_id, role, unit, unit_key, full_text,"
                " section, paragraph) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (position_id, file_id, item_id, hit.role, unit, position_key(unit), hit.text,
                 hit.section, hit.paragraph),
            )
            _save_chain(connection, file_id, item_id, hit)
    summary.indexed += 1
    summary.items += len(items)
    summary.hits += hits_count


def _save_item(connection: sqlite3.Connection, file_id: int, item: OrderItem) -> int:
    """Пункт наказу цілком: особа, посади й текст (лише в локальному індексі)."""
    cursor = connection.execute(
        "INSERT INTO items(file_id, section, label, paragraph, rank, surname, name, patronymic, person, ipn,"
        " current_position, target_position, text) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (file_id, item.section, item.label, item.paragraph, item.rank, item.surname, item.name,
         item.patronymic, person_key(item.surname, item.name, item.patronymic), item.ipn,
         item.current_position, item.target_position, item.text),
    )
    return int(cursor.lastrowid)


def _save_chain(connection: sqlite3.Connection, file_id: int, item_id: int, hit) -> None:
    """Ланцюг підрозділів посади: взвод → рота → пункт → батальйон → полк."""
    # Ланцюг рахується від повного тексту: перша ланка й так починається з виду
    # підрозділу, а обрізання за назвою посади розривало «ударного взводу …».
    for link in split_chain(hit.text):
        key = chain_key(link.text)
        if not key:
            continue
        row = connection.execute("SELECT id FROM subdivisions WHERE key = ?", (key,)).fetchone()
        if row is None:
            cursor = connection.execute(
                "INSERT INTO subdivisions(key, display, kind) VALUES (?, ?, ?)",
                (key, link.text.casefold(), link.kind),
            )
            subdivision_id = int(cursor.lastrowid)
        else:
            subdivision_id = int(row[0])
        connection.execute(
            "INSERT INTO item_subdivisions(item_id, subdivision_id, file_id, role, level, text)"
            " VALUES (?, ?, ?, ?, ?, ?)",
            (item_id, subdivision_id, file_id, hit.role, link.level, link.text),
        )


def _forget_file(connection: sqlite3.Connection, file_id: int) -> None:
    connection.execute("DELETE FROM item_subdivisions WHERE file_id = ?", (file_id,))
    connection.execute("DELETE FROM occurrences WHERE file_id = ?", (file_id,))
    connection.execute("DELETE FROM items WHERE file_id = ?", (file_id,))
    connection.execute("DELETE FROM files WHERE id = ?", (file_id,))


def _position_id(connection: sqlite3.Connection, text: str, known: bool) -> int:
    key = ("+ " if known else "? ") + position_key(text)
    row = connection.execute("SELECT id FROM positions WHERE key = ?", (key,)).fetchone()
    if row is None:
        cursor = connection.execute(
            "INSERT INTO positions(key, display, known) VALUES (?, ?, ?)", (key, text.casefold(), int(known))
        )
        return int(cursor.lastrowid)
    return int(row[0])


def _remove_missing(connection: sqlite3.Connection, seen: set[str], root: Path, recursive: bool) -> int:
    """Прибирає зниклі файли лише з тієї частини папки, яку щойно переглянули."""
    removed = 0
    for file_id, path in connection.execute("SELECT id, path FROM files").fetchall():
        parent = Path(path).parent
        scanned = parent == root or (recursive and root in parent.parents)
        if scanned and path not in seen:
            _forget_file(connection, file_id)
            removed += 1
    return removed


def _drop_orphan_positions(connection: sqlite3.Connection) -> None:
    connection.execute("DELETE FROM positions WHERE id NOT IN (SELECT DISTINCT position_id FROM occurrences)")
    connection.execute(
        "DELETE FROM subdivisions WHERE id NOT IN (SELECT DISTINCT subdivision_id FROM item_subdivisions)"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Читання для вікна
# ─────────────────────────────────────────────────────────────────────────────
@dataclass(frozen=True)
class PositionRow:
    id: int
    display: str
    known: int
    mentions: int
    orders: int
    units: int
    current: int
    target: int
    first_date: str
    last_date: str


@dataclass(frozen=True)
class OccurrenceRow:
    path: str
    number: str
    date: str
    section: str
    role: str
    unit: str
    full_text: str
    paragraph: int


@dataclass(frozen=True)
class SubdivisionRow:
    """Ланка підрозділу в індексі: скільки разів і в скількох наказах трапляється."""

    id: int
    display: str
    kind: str
    mentions: int
    orders: int
    items: int
    first_date: str
    last_date: str


@dataclass(frozen=True)
class PersonItem:
    """Пункт наказу про особу — основа для перенесення в новий наказ."""

    id: int
    rank: str
    surname: str
    name: str
    patronymic: str
    ipn: str
    current_position: str
    target_position: str
    section: str
    label: str
    text: str
    order_number: str
    order_date: str
    path: str

    @property
    def full_name(self) -> str:
        return " ".join(part for part in (self.surname, self.name, self.patronymic) if part)


@dataclass(frozen=True)
class IndexStats:
    files: int
    ok: int
    legacy: int
    errors: int
    positions: int
    mentions: int
    items: int
    people: int


def load_positions(index_folder: str | Path) -> list[PositionRow]:
    if not index_path(index_folder).exists():
        return []
    connection = connect(index_folder)
    try:
        rows = connection.execute(
            """
            SELECT p.id, p.display, p.known,
                   COUNT(*), COUNT(DISTINCT o.file_id), COUNT(DISTINCT NULLIF(o.unit_key, '')),
                   SUM(o.role = 'займана'), SUM(o.role = 'призначення'),
                   COALESCE(MIN(NULLIF(f.date, '')), ''), COALESCE(MAX(NULLIF(f.date, '')), '')
            FROM positions p
            JOIN occurrences o ON o.position_id = p.id
            JOIN files f ON f.id = o.file_id
            GROUP BY p.id
            ORDER BY COUNT(*) DESC, p.key
            """
        ).fetchall()
    finally:
        connection.close()
    return [PositionRow(*row) for row in rows]


def load_occurrences(index_folder: str | Path, position_id: int) -> list[OccurrenceRow]:
    connection = connect(index_folder)
    try:
        rows = connection.execute(
            """
            SELECT f.path, f.number, f.date, o.section, o.role, o.unit, o.full_text, o.paragraph
            FROM occurrences o JOIN files f ON f.id = o.file_id
            WHERE o.position_id = ?
            ORDER BY f.date DESC, f.path, o.paragraph
            """,
            (position_id,),
        ).fetchall()
    finally:
        connection.close()
    return [OccurrenceRow(*row) for row in rows]


def load_stats(index_folder: str | Path) -> IndexStats | None:
    if not index_path(index_folder).exists():
        return None
    connection = connect(index_folder)
    try:
        files, ok, legacy, errors = connection.execute(
            "SELECT COUNT(*), COALESCE(SUM(status='ok'),0), COALESCE(SUM(status='legacy'),0),"
            " COALESCE(SUM(status='error'),0) FROM files"
        ).fetchone()
        positions = connection.execute("SELECT COUNT(*) FROM positions").fetchone()[0]
        mentions = connection.execute("SELECT COUNT(*) FROM occurrences").fetchone()[0]
        items = connection.execute("SELECT COUNT(*) FROM items").fetchone()[0]
        people = connection.execute(
            "SELECT COUNT(*) FROM (SELECT DISTINCT COALESCE(NULLIF(ipn, ''), person) FROM items)"
        ).fetchone()[0]
    finally:
        connection.close()
    return IndexStats(files, ok, legacy, errors, positions, mentions, items, people)


_PERSON_ITEM_QUERY = """
    SELECT i.id, i.rank, i.surname, i.name, i.patronymic, i.ipn, i.current_position, i.target_position,
           i.section, i.label, i.text, f.number, f.date, f.path
    FROM items i JOIN files f ON f.id = i.file_id
    WHERE {condition}
    ORDER BY f.date DESC, f.path DESC, i.paragraph DESC
    LIMIT ?
"""


def load_subdivisions(index_folder: str | Path) -> list[SubdivisionRow]:
    """Усі підрозділи з наказів: «роти безпілотних систем», «901 механізованого батальйону»."""
    if not index_path(index_folder).exists():
        return []
    connection = connect(index_folder)
    try:
        rows = connection.execute(
            """
            SELECT s.id, s.display, s.kind, COUNT(*), COUNT(DISTINCT l.file_id), COUNT(DISTINCT l.item_id),
                   COALESCE(MIN(NULLIF(f.date, '')), ''), COALESCE(MAX(NULLIF(f.date, '')), '')
            FROM subdivisions s
            JOIN item_subdivisions l ON l.subdivision_id = s.id
            JOIN files f ON f.id = l.file_id
            GROUP BY s.id
            ORDER BY COUNT(*) DESC, s.key
            """
        ).fetchall()
    finally:
        connection.close()
    return [SubdivisionRow(*row) for row in rows]


def find_person_items(
    index_folder: str | Path, ipn: str = "", full_name: str = "", limit: int = 20
) -> list[PersonItem]:
    """Пункти наказів про особу — найсвіжіші першими.

    Пошук за РНОКПП (точний) або за ПІБ: прізвище, «прізвище ім'я» чи ПІБ повністю.
    """
    ipn = "".join(character for character in str(ipn or "") if character.isdigit())
    key = person_key(*str(full_name or "").split()[:3]) if full_name else ""
    if not ipn and not key:
        return []
    if not index_path(index_folder).exists():
        return []
    connection = connect(index_folder)
    try:
        if ipn:
            rows = connection.execute(_PERSON_ITEM_QUERY.format(condition="i.ipn = ?"), (ipn, limit)).fetchall()
            if rows:
                return [PersonItem(*row) for row in rows]
        if not key:
            return []
        rows = connection.execute(
            _PERSON_ITEM_QUERY.format(condition="i.person = ? OR i.person LIKE ?"),
            (key, key + " %", limit),
        ).fetchall()
    finally:
        connection.close()
    return [PersonItem(*row) for row in rows]
