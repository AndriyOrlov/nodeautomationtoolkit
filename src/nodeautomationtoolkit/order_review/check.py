"""Правила перевірки наказу: оформлення + історія осіб з індексу попередніх наказів.

Модуль працює лише з текстом (Word тут не потрібен), тож тестується на
вигаданих наказах. Позначену копію документа робить `marking.py`.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from ..builtin_nodes.message_order import find_content_start_line
from ..builtin_nodes.recipient_mapping import map_military_units
from ..builtin_nodes.typography import ORDER_SIGNER_START_RE
from ..order_index.extractor import normalize_space
from ..order_index.items import iter_items, person_key
from ..order_index.position_dictionary import load_position_dictionary, stem
from ..order_index.store import find_person_items, index_path, load_positions, load_stats
from ..personnel.dictionaries import load_keyed_table
from ..personnel.ipn_audit import _split_items as split_ipn_items
from ..personnel.ipn_audit import find_ipn_problems

ERROR = "помилка"
WARNING = "увага"

#: Слово з попередніх наказів вважається «відомим», лише якщо траплялось стільки разів:
#: помилку, зроблену одного разу, програма не має вивчити як норму.
MIN_INDEX_MENTIONS = 2
#: Пословна звірка назв посад («незнайоме слово») вимкнена: на реальних наказах давала
#: забагато жовтого (рішення користувача 24.09.2026). Код лишається — увімкнути тут.
CHECK_POSITION_SPELLING = False

_ITEM_RE = re.compile(r"^\s*(\d{1,3}(?:\.\d{1,3})*)[.)]\s+\S")
_GLUED_ITEM_RE = re.compile(r"^\s*(\d{1,3}(?:\.\d{1,3})*)[.)](?=[^\W\d_]|[«\"“'])")
_WORD_RE = re.compile(r"[^\W\d_]+(?:['’ʼ-][^\W\d_]+)*|\d+")


@dataclass
class Finding:
    """Одна знахідка: що не так, де і як виправити."""

    level: str
    rule: str
    where: str
    what: str
    how: str = ""
    #: Фрагменти тексту наказу, які треба позначити в копії (перший знайдений).
    quotes: tuple[str, ...] = ()


@dataclass
class ReviewResult:
    findings: list[Finding] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    people_checked: int = 0
    people_found: int = 0
    indexed_orders: int = 0

    @property
    def errors(self) -> int:
        return sum(finding.level == ERROR for finding in self.findings)

    @property
    def warnings(self) -> int:
        return sum(finding.level == WARNING for finding in self.findings)


def group_findings(findings: list[Finding], preview: int = 6) -> list[str]:
    """Стислий підсумок за групами (AGENT.md 7.2): скільки і де, без повторів опису."""
    groups: dict[tuple[str, str], list[Finding]] = {}
    for finding in findings:
        groups.setdefault((finding.level, finding.rule), []).append(finding)
    lines = []
    for (level, rule), group in sorted(
        groups.items(), key=lambda pair: (pair[0][0] != ERROR, -len(pair[1]))
    ):
        places = list(dict.fromkeys(finding.where for finding in group))
        shown = ", ".join(places[:preview]) + (
            f" … ще {len(places) - preview}" if len(places) > preview else ""
        )
        lines.append(f"{'✖' if level == ERROR else '⚠'} {rule}: {len(group)} — {shown}")
    return lines


def findings_to_tsv(result: ReviewResult) -> str:
    """Відповідь для макросу Word: той самий рядок знахідки, що дає ядро макросу.

    Рядок знахідки — `рівень, правило, де, що, як, фрагмент1, фрагмент2` через TAB;
    рядок «#note» — нотатка перевірки (індекс, таблиця).
    """

    def clean(value: str) -> str:
        return " ".join(str(value or "").replace("\t", " ").split())

    lines = [f"#note\t{clean(note)}" for note in result.notes]
    for finding in result.findings:
        quotes = (list(finding.quotes) + ["", ""])[:2]
        fields = [finding.level, finding.rule, finding.where, finding.what, finding.how, *quotes]
        lines.append("\t".join(clean(field) for field in fields))
    return "\n".join(lines) + "\n"


# ─────────────────────────────────────────────────────────────────────────────
# Оформлення
# ─────────────────────────────────────────────────────────────────────────────
def check_metadata(order_number: str, order_date: str) -> list[Finding]:
    how = (
        "Назвіть файл за зразком «Наказ № 123 від 15.09.2026.docx» — програма бере їх лише з назви."
    )
    findings = []
    if not order_number:
        findings.append(
            Finding(WARNING, "Реквізити в назві файлу", "Назва файлу", "немає номера наказу", how)
        )
    if not order_date:
        findings.append(
            Finding(WARNING, "Реквізити в назві файлу", "Назва файлу", "немає дати наказу", how)
        )
    return findings


def check_signer(signer: dict | None) -> list[Finding]:
    if signer is None or any(
        str(signer.get(key, "")).strip() for key in ("position", "rank", "name")
    ):
        return []
    return [
        Finding(
            WARNING,
            "Підписант",
            "Кінець наказу",
            "блок підписанта не розпізнано",
            "Після останнього пункту мають стояти посада, звання й ПІБ підписанта.",
        )
    ]


def check_numbering(text: str) -> list[Finding]:
    """Номери пунктів: пропуск, повтор, підпункт не під своїм пунктом, номер без пропуску.

    Нумерація наскрізна через увесь наказ: § і підшапки її не перезапускають
    (користувач, 25.09.2026 — «нумерація йде по всьому наказу спільна»).
    """
    findings: list[Finding] = []
    lines = text.splitlines()
    start = find_content_start_line(text)
    previous_top = 0
    sub_counters: dict[tuple[int, ...], int] = {}
    for line in lines[start:]:
        clean = line.strip()
        if clean.startswith("§"):
            continue
        glued = _GLUED_ITEM_RE.match(clean)
        if glued:
            findings.append(
                Finding(
                    WARNING,
                    "Нумерація пунктів",
                    f"Пункт {glued.group(1)}",
                    "після номера пункту немає пропуску",
                    "Пишіть «1. Капітана …», а не «1.Капітана …».",
                    (clean[:60],),
                )
            )
        match = _ITEM_RE.match(clean) or glued
        if not match:
            continue
        label = match.group(1)
        parts = tuple(int(part) for part in label.split("."))
        quote = (clean[:60],)
        where = f"Пункт {label}"
        if len(parts) == 1:
            number = parts[0]
            if number == previous_top:
                findings.append(
                    Finding(
                        ERROR,
                        "Нумерація пунктів",
                        where,
                        "номер пункту повторюється",
                        f"Після пункту {previous_top} має йти {previous_top + 1}.",
                        quote,
                    )
                )
            elif number != previous_top + 1:
                findings.append(
                    Finding(
                        ERROR,
                        "Нумерація пунктів",
                        where,
                        f"після пункту {previous_top} йде пункт {number}",
                        f"Очікувався пункт {previous_top + 1}. Нумерація наскрізна через увесь наказ.",
                        quote,
                    )
                )
            previous_top = number
            continue
        parent = parts[:-1]
        if parent[0] != previous_top:
            findings.append(
                Finding(
                    ERROR,
                    "Нумерація пунктів",
                    where,
                    f"підпункт стоїть не під пунктом {parent[0]}",
                    f"Підпункт {label} має йти одразу після пункту {parent[0]}.",
                    quote,
                )
            )
        expected = sub_counters.get(parent, 0) + 1
        if parts[-1] != expected:
            findings.append(
                Finding(
                    ERROR,
                    "Нумерація пунктів",
                    where,
                    f"очікувався підпункт {'.'.join(map(str, parent))}.{expected}",
                    "Підпункти нумеруються підряд з 1.",
                    quote,
                )
            )
        sub_counters[parent] = parts[-1]
    return findings


def check_ipn(text: str) -> list[Finding]:
    """Контрольна цифра — червона: це майже завжди описка в номері.

    Розбіжність дати (року) народження чи статі з РНОКПП — жовта: перевірка
    приблизна, і помилка може бути як у номері, так і в даті.
    """
    findings = []
    for problem in find_ipn_problems(text):
        where = f"Пункт {problem.item}" if problem.item else "Текст до першого пункту"
        quotes = (problem.token,) if problem.token else ()
        if problem.kind == "check_digit":
            findings.append(
                Finding(
                    ERROR,
                    "РНОКПП",
                    where,
                    "РНОКПП не проходить контрольну перевірку",
                    "Звірте цифри з документом особи — найчастіше це описка.",
                    quotes,
                )
            )
        elif problem.kind == "sex":
            findings.append(
                Finding(
                    WARNING,
                    "РНОКПП",
                    where,
                    f"стать у РНОКПП не збігається з «{problem.detail}»",
                    "Перевірте РНОКПП і «Народився/Народилася».",
                    quotes,
                )
            )
        else:
            what = "дата" if problem.kind == "birth_date" else "рік"
            findings.append(
                Finding(
                    WARNING,
                    "РНОКПП",
                    where,
                    f"{what} народження {problem.detail} не збігається з РНОКПП",
                    "Перевірка приблизна: звірте дату народження і РНОКПП з документом.",
                    quotes,
                )
            )
    return findings


# ─────────────────────────────────────────────────────────────────────────────
# Запис біографії: рік народження й «у ЗС»
# ─────────────────────────────────────────────────────────────────────────────
#: Будь-яка згадка року народження: «1974 р.н.», «02.03.1988 р.н.», «1974 року народження».
_BIRTH_MENTION_RE = re.compile(
    r"(?<![\w.])(?:(\d{1,2})\.(\d{1,2})\.)?((?:19|20)\d{2})(\s*)(р\.?\s*н\.?|року\s+народження)",
    re.IGNORECASE,
)
#: «у ЗС» / «у ЗСУ» як слово; «Вислуга років у ЗС: …» (звільнення) — не запис служби.
_SERVICE_START_RE = re.compile(r"(?<![\w])[уУ]\s*ЗСУ?(?![\w])(?!\s*:)")
_PERIOD = r"(?:0[1-9]|1[0-2])\.(?:19|20)\d{2}"
_RANGE = rf"{_PERIOD}(?:\s+по\s+{_PERIOD})?"
#: «у ЗС – із 07.1991.»; «у ЗС – із 06.1997 по 06.2010 та з 01.2017.» — періодів скільки завгодно.
_SERVICE_RE = re.compile(
    rf"[уУ]\s+ЗС\s*[–—-]\s*із\s+{_RANGE}(?:(?:\s*,\s*|\s+та\s+)(?:із|з)\s+{_RANGE})*\."
)
#: Після «р.н.»: кома («1978 р.н., освіта …») або — у біографічному блоці — табуляції й
#: РНОКПП із крапкою в кінці («1979 р.н.<TAB><TAB>2900000000.»; користувач, 25.09.2026).
_AFTER_BIRTH_RE = re.compile(r",|[\t  ]+\d{10}\.")
#: Нерозривний пробіл (Ctrl+Shift+Пробіл) — теж правильний пробіл у «1978 р.н.,».
NBSP = "\u00a0"
SERVICE_SAMPLE ="Пишіть «у ЗС – із 07.1991.» або «у ЗС – із 06.1997 по 06.2010 та з 01.2017.»."


def _item_where(label: str) -> str:
    return f"Пункт {label}" if label else "Текст до першого пункту"


def _sentence_from(text: str, start: int) -> str:
    """Від `start` до кінця речення (крапка перед пропуском або кінцем рядка), не довше 60."""
    rest = text[start:].split("\n", 1)[0]
    end = re.search(r"\.(?=\s|$)", rest)
    return rest[: end.end() if end else len(rest)][:60]


def check_bio_format(text: str) -> list[Finding]:
    """Рік народження — лише рік, саме так: «1978 р.н.,»; «у ЗС» — за зразком (жовте).

    Вимоги користувача 25.09.2026. «Народився 11 серпня 1976 року» (накази про
    звільнення, додаток 53) не чіпається: там немає «р.н.».
    """
    birth_findings, service_findings = [], []
    for label, body in split_ipn_items(text):
        where = _item_where(label)
        for match in _BIRTH_MENTION_RE.finditer(body):
            year, spacing, form = match.group(3), match.group(4), match.group(5)
            if match.group(1):
                birth_findings.append(
                    Finding(
                        WARNING,
                        "Рік народження",
                        where,
                        f"дата народження «{match.group(0)}» — пишеться лише рік",
                        f"Пишіть рік без дня й місяця: «{year} р.н.,».",
                        (match.group(0),),
                    )
                )
            elif (
                spacing not in (" ", NBSP)
                or form != "р.н."
                or not _AFTER_BIRTH_RE.match(body, match.end())
            ):
                birth_findings.append(
                    Finding(
                        WARNING,
                        "Рік народження",
                        where,
                        f"«{match.group(0)}» — не за зразком",
                        f"Пишіть «{year} р.н.,».",
                        (match.group(0),),
                    )
                )
        for match in _SERVICE_START_RE.finditer(body):
            if _SERVICE_RE.match(body, match.start()):
                continue
            fragment = _sentence_from(body, match.start())
            service_findings.append(
                Finding(
                    WARNING, "Запис «у ЗС»", where, f"«{fragment}» — не за зразком",
                    SERVICE_SAMPLE, (fragment,),
                )
            )
    return birth_findings + service_findings


# ─────────────────────────────────────────────────────────────────────────────
# Задвоєні знаки й пробіли; два порожні абзаци перед підписантом
# ─────────────────────────────────────────────────────────────────────────────
#: «..» (але не «...»), «,,», «;;», «::», «!!», «??», рівно два пробіли між словами.
#: Три й більше пробілів — вирівнювання (шапка, підписант), не помилка.
_DOUBLE_DOTS_RE = re.compile(r"(?<!\.)\.\.(?!\.)")
_DOUBLE_MARKS_RE = re.compile(r"([,;:!?])\1+")
_DOUBLE_SPACE_RE = re.compile(r"(?<=\S)  (?=\S)")


def _around(line: str, start: int, end: int) -> str:
    """Фрагмент рядка навколо знахідки — щоб знайти її в документі Word."""
    return line[max(0, start - 20) : end + 10]


def check_doubles(text: str) -> list[Finding]:
    """Задвоєні розділові знаки й два пробіли підряд (жовте)."""
    findings = []
    for label, body in split_ipn_items(text):
        where = _item_where(label)
        for line in body.splitlines():
            for match in _DOUBLE_DOTS_RE.finditer(line):
                findings.append(Finding(
                    WARNING, "Задвоєні знаки", where, f"задвоєний знак «{match.group(0)}»",
                    "Приберіть зайвий знак.", (_around(line, match.start(), match.end()),),
                ))
            for match in _DOUBLE_MARKS_RE.finditer(line):
                findings.append(Finding(
                    WARNING, "Задвоєні знаки", where, f"задвоєний знак «{match.group(0)}»",
                    "Приберіть зайвий знак.", (_around(line, match.start(), match.end()),),
                ))
            for match in _DOUBLE_SPACE_RE.finditer(line):
                findings.append(Finding(
                    WARNING, "Задвоєні знаки", where, "два пробіли підряд",
                    "Залиште один пробіл.", (_around(line, match.start(), match.end()),),
                ))
    return findings


def check_signer_spacing(text: str) -> list[Finding]:
    """Перед підписантом наказу — рівно два порожні абзаци (жовте).

    Підписант — перший рядок після останнього пункту, що починається з «Командир»,
    «Начальник», «Командувач», «Т.в.о.», «Тимчасово виконуючий» (як у генераторі).
    Потрібен ПОВНИЙ текст: `text_before_order_signer` порожні рядки зрізає.
    """
    lines = text.splitlines()
    start = find_content_start_line(text)
    last_item = max(
        (index for index in range(start, len(lines)) if _ITEM_RE.match(lines[index].strip())),
        default=None,
    )
    if last_item is None:
        return []
    for index in range(last_item + 1, len(lines)):
        if not ORDER_SIGNER_START_RE.match(lines[index].strip()):
            continue
        blanks = 0
        while index - blanks - 1 > last_item and not lines[index - blanks - 1].strip():
            blanks += 1
        if blanks == 2:
            return []
        return [
            Finding(
                WARNING,
                "Порожні абзаци",
                "Кінець наказу",
                f"порожніх абзаців перед підписантом: {blanks} — має бути 2",
                "Між останнім пунктом і підписантом наказу — рівно два порожні абзаци.",
                (lines[index].strip()[:60],),
            )
        ]
    return []


def check_routing(text: str, mapping: dict | None) -> list[Finding]:
    """Пункти, для яких програма не знайде адресата: витяг за ними не створиться."""
    if not mapping:
        return []
    routes = map_military_units(text=text, mapping=mapping)
    findings = []
    for item in routes.get("unmatched_items", []):
        label = str(item.get("label", "")).rstrip(".") or "Пункт"
        first_line = next(
            (line.strip() for line in str(item.get("text", "")).splitlines() if line.strip()), ""
        )
        findings.append(
            Finding(
                WARNING,
                "Адресат",
                label,
                "не знайдено частини чи ТЦК з таблиці — витяг за цим пунктом не створиться",
                "Перевірте написання назви частини (як у стовпці A таблиці) або додайте рядок у таблицю.",
                (first_line[:60],) if first_line else (),
            )
        )
    return findings


# ─────────────────────────────────────────────────────────────────────────────
# Індекс попередніх наказів
# ─────────────────────────────────────────────────────────────────────────────
def _stems(text: str) -> tuple[str, ...]:
    return tuple(stem(word) for word in _WORD_RE.findall(normalize_space(text).casefold()))


def same_position(first: str, second: str) -> bool:
    """Та сама посада в будь-якому відмінку й регістрі.

    Коротший запис може бути початком довшого: у шапці розділу часто стоїть
    частина, тож у самому пункті посада записана без неї.
    """
    a, b = _stems(first), _stems(second)
    if not a or not b:
        return True
    length = min(len(a), len(b))
    return a[:length] == b[:length]


def _rank_levels() -> list[tuple[tuple[str, ...], int]]:
    levels = []
    for name, row in load_keyed_table("shpk_levels.csv").items():
        try:
            levels.append((_stems(name), int(row.get("Коефіцієнт", ""))))
        except ValueError:
            continue
    return sorted(levels, key=lambda pair: -len(pair[0]))


def rank_level(rank_text: str, levels: list | None = None) -> int | None:
    """«Старшого лейтенанта», «КАПІТАНОМ» → рівень звання (солдат = 1 … генерал = 21)."""
    words = _stems(rank_text)
    for name, level in levels if levels is not None else _rank_levels():
        if name and words[-len(name) :] == name:
            return level
    return None


def _iso_date(order_date: str) -> str:
    for pattern in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(order_date or "").strip(), pattern).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def _human_date(iso: str) -> str:
    try:
        return datetime.strptime(iso, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return iso or "без дати"


def _order_ref(previous) -> str:
    """« №12 від 05.03.2026» — дописується до «наказом», «наказі» тощо."""
    number = f"№{previous.order_number}" if previous.order_number else "без номера"
    return f" {number} від {_human_date(previous.order_date)}"


def order_items(text: str):
    """Пункти з особами — той самий розбір, що й під час індексування."""
    lines = text.splitlines()
    return iter_items(lines[find_content_start_line(text) :], load_position_dictionary())


# ─────────────────────────────────────────────────────────────────────────────
# Номенклатура: до якого звання наш наказ може звільняти й призначати
# ─────────────────────────────────────────────────────────────────────────────
NOMENCLATURE_FILE = "nomenclature.csv"
#: Корабельні звання: у Сухопутних військах їх не буває.
_NAVAL_RANK_RE = re.compile(r"матрос|старшин|\bрангу\b|капітан-лейтенант|коммодор|адмірал", re.I)
_QUOTED_RE = re.compile(r"[«\"“]([^»\"”]{3,40})[»\"”]")
_HEADING_HINT_RE = re.compile(r"відповідно\s+до|згідно\s+з|звільнити|призначити|присвоїти|:\s*$", re.I)


def load_nomenclature() -> dict[str, tuple[int | None, str, bool, str]]:
    """{дія: (рівень найвищого звання, звання, чи допустимі корабельні звання, підстава)}.

    Порожнє «Найвище звання» — звання не обмежене (призначення, переміщення), але
    правило корабельних звань діє.
    """
    limits = {}
    for action, row in load_keyed_table(NOMENCLATURE_FILE).items():
        name = row.get("Найвище звання", "").strip()
        level = rank_level(name) if name else None
        if name and level is None:
            continue  # описка в назві звання — рядок не застосовуємо
        naval = row.get("Корабельні звання", "").strip().casefold() in {"так", "+", "1"}
        limits[action] = (level, name, naval, row.get("Підстава", ""))
    return limits


def _order_action(text: str) -> str:
    clean = " ".join(normalize_space(text).casefold().split())
    if re.search(r"звільнити\s+з\s+військової\s+служби", clean):
        return "звільнення"
    if "призначити" in clean:
        return "призначення"
    return ""


def _heading_context(lines: list[str], paragraph: int) -> str:
    """Шапки над пунктом (до § або початку тексту): з них видно дію наказу."""
    context = []
    for line in reversed(lines[:paragraph]):
        clean = line.strip()
        if clean.startswith("§"):
            break
        if clean and not _ITEM_RE.match(clean) and _HEADING_HINT_RE.search(clean):
            context.append(clean)
    return "\n".join(reversed(context))


def check_nomenclature(text: str, items) -> list[Finding]:
    """Звання особи й шпк у лапках не вище межі номенклатури; лише військові звання.

    Межі — `personnel/dictionaries/nomenclature.csv` (користувач править сам): для
    звільнення — Положення (Указ № 1153/2008), п. 225; для призначення — номенклатура
    посад, яку затверджує Міністр оборони (п. 81).
    """
    limits = load_nomenclature()
    if not limits:
        return []
    levels = _rank_levels()
    lines = text.splitlines()[find_content_start_line(text) :]
    findings = []
    for item in items:
        if not item.surname:
            continue
        action = _order_action(_heading_context(lines, item.paragraph) + "\n" + item.text)
        limit = limits.get(action) or limits.get("усі")
        if not limit:
            continue
        max_level, max_name, naval_allowed, basis = limit
        where, first_line = _item_place(item), _first_line(item)
        what_action = {"звільнення": "звільнення", "призначення": "призначення"}.get(action, "наказу")
        candidates = [("звання особи", item.rank)] + [
            ("шпк / звання в лапках", quoted) for quoted in _QUOTED_RE.findall(item.text)
        ]
        for label, rank in candidates:
            level = rank_level(rank, levels)
            if level is not None and max_level is not None and level > max_level:
                findings.append(
                    Finding(
                        ERROR,
                        "Номенклатура",
                        where,
                        f"{label} «{rank.strip()}» вище за {max_name} — поза номенклатурою {what_action}",
                        basis,
                        (rank.strip(), first_line),
                    )
                )
            if not naval_allowed and rank and _NAVAL_RANK_RE.search(rank) and level is not None:
                findings.append(
                    Finding(
                        WARNING,
                        "Номенклатура",
                        where,
                        f"{label} «{rank.strip()}» — корабельне звання, а не Сухопутних військ",
                        "Перевірте звання: у Сухопутних військах лише військові звання.",
                        (rank.strip(), first_line),
                    )
                )
    return findings


# ─────────────────────────────────────────────────────────────────────────────
# Алфавітний порядок прізвищ у групі
# ─────────────────────────────────────────────────────────────────────────────
#: Українська абетка: звичайне сортування ставить ґ, є, і, ї після «я».
UKRAINIAN_ALPHABET = "абвгґдеєжзиіїйклмнопрстуфхцчшщьюя"


def alphabet_key(word: str) -> tuple[int, ...]:
    """Позиції літер основи слова в українській абетці (апостроф, дефіс — пропускаються).

    Основа (`stem`), а не саме слово: у наказі прізвище у знахідному чи давальному, і
    «КОВАЛЯ» (Коваль) має стояти перед «КОВАЛЕНКА» (Коваленко), хоча «я» > «е».
    """
    return tuple(
        UKRAINIAN_ALPHABET.index(char) + 1 if char in UKRAINIAN_ALPHABET else 100
        for char in stem(word)
        if char.isalpha()
    )


def person_sort_key(item) -> tuple:
    """Прізвище, ім'я, по батькові. 0 між словами — менше за будь-яку літеру: коротше
    прізвище, яке є початком довшого, стоїть першим (так само рахує макрос)."""
    key: list[int] = []
    for word in (item.surname, item.name, item.patronymic):
        key.extend(alphabet_key(word))
        key.append(0)
    return tuple(key)


def _starts_new_group(lines: list[str], previous: int, current: int) -> bool:
    """Чи є між двома пунктами шапка: §, «Відповідно до …:», «У ЗАПАС ЗА …», «МАЙОР»."""
    for line in lines[previous + 1 : current]:
        clean = line.strip()
        if not clean:
            continue
        letters = [char for char in clean if char.isalpha()]
        if (
            clean.startswith("§")
            or _HEADING_HINT_RE.search(clean)
            or (letters and not any(char.islower() for char in letters))
        ):
            return True
    return False


def check_alphabet(text: str, items) -> list[Finding]:
    """Пункти однієї групи — за прізвищами в алфавітному порядку (жовте).

    Група — пункти під однією шапкою; нова шапка, § чи підзаголовок («МАЙОР», «У ЗАПАС
    ЗА ПІДПУНКТОМ …») починає нову. Підпункти («1.1.») не порівнюються.
    """
    lines = text.splitlines()[find_content_start_line(text) :]
    findings = []
    previous = None
    for item in items:
        if not item.surname or "." in item.label:
            continue
        if (
            previous is not None
            and previous.section == item.section
            and not _starts_new_group(lines, previous.paragraph, item.paragraph)
            and person_sort_key(previous) > person_sort_key(item)
        ):
            findings.append(
                Finding(
                    WARNING,
                    "Алфавітний порядок",
                    _item_place(item),
                    f"«{item.surname}» стоїть після «{previous.surname}» — порушено абетку",
                    "Пункти однієї групи (під однією шапкою) йдуть за прізвищами за абеткою.",
                    (_first_line(item),),
                )
            )
        previous = item
    return findings


def check_history(
    items,
    index_folder: str | Path,
    order_number: str = "",
    order_date: str = "",
    result: ReviewResult | None = None,
) -> list[Finding]:
    """Звірка осіб наказу з їхніми попередніми наказами в індексі."""
    result = result if result is not None else ReviewResult()
    stats = load_stats(index_folder)
    result.indexed_orders = stats.ok if stats else 0
    levels = _rank_levels()
    current_iso = _iso_date(order_date)

    findings: list[Finding] = []
    for item in items:
        if not item.surname:
            continue
        result.people_checked += 1
        previous_items = [
            previous
            for previous in find_person_items(
                index_folder, ipn=item.ipn, full_name=item.full_name, limit=10
            )
            if not (previous.order_number == order_number and previous.order_date == current_iso)
            and (not current_iso or not previous.order_date or previous.order_date < current_iso)
        ]
        if previous_items:
            result.people_found += 1
            findings.extend(
                _person_findings(
                    item, previous_items[0], _item_place(item), _first_line(item), levels
                )
            )
    return findings


def _item_place(item) -> str:
    return f"Пункт {item.label}" if item.label else "Пункт без номера"


def _first_line(item) -> str:
    return item.text.split("\n", 1)[0][:60]


def position_vocabulary(index_folder: str | Path | None = None) -> tuple[set[str], set[str]]:
    """(основи слів, слова) з довідника посад і з попередніх наказів."""
    dictionary = load_position_dictionary()
    words = {
        word
        for nominative in dictionary.phrases.values()
        for word in _WORD_RE.findall(nominative.casefold())
    }
    if index_folder and index_path(index_folder).exists():
        for row in load_positions(index_folder):
            if row.mentions >= MIN_INDEX_MENTIONS:
                words.update(_WORD_RE.findall(row.display.casefold()))
    stems = {part for key in dictionary.phrases for part in key}
    stems.update(stem(word) for word in words)
    return stems, {word for word in words if not word.isdigit()}


def check_position_spelling(items, index_folder: str | Path | None = None) -> list[Finding]:
    """Слово в назві посади, якого немає ні в довіднику, ні в попередніх наказах.

    Довідник упізнає посаду за першими словами («заступник …», «начальник …») і
    дописує наступний іменник без перевірки, тож «ЗАСТУПНИКОМ КОМАНДИРРА» вважалося
    відомою посадою. Тому звіряється КОЖНЕ слово назви.
    """
    stems, words = position_vocabulary(index_folder)
    ordered_words = sorted(words)
    findings = []
    for item in items:
        for hit in item.hits:
            for word in _WORD_RE.findall(hit.position.casefold()):
                if word.isdigit() or len(word) < 3 or stem(word) in stems:
                    continue
                similar = difflib.get_close_matches(word, ordered_words, n=1, cutoff=0.75)
                findings.append(
                    Finding(
                        WARNING,
                        "Написання посади",
                        _item_place(item),
                        f"у посаді «{hit.position}» незнайоме слово «{word}»",
                        (
                            f"Можливо, «{similar[0]}»."
                            if similar
                            else "Такого слова немає ні в довіднику "
                            "посад, ні в попередніх наказах — перевірте за штатом."
                        ),
                        (hit.text[:60], _first_line(item)),
                    )
                )
    return findings


def _person_findings(item, previous, where: str, first_line: str, levels) -> list[Finding]:
    findings = []
    reference = _order_ref(previous)
    same_ipn = bool(item.ipn and previous.ipn and item.ipn == previous.ipn)
    same_person = person_key(item.surname, item.name, item.patronymic) == person_key(
        previous.surname, previous.name, previous.patronymic
    )
    if same_ipn and not same_person:
        findings.append(
            Finding(
                ERROR,
                "Історія особи",
                where,
                f"цей РНОКПП у наказі{reference} належить іншій особі",
                f"У наказі{reference}: {previous.full_name}. Звірте ПІБ і РНОКПП.",
                (item.ipn, first_line),
            )
        )
        return findings
    if not same_ipn and item.ipn and previous.ipn:
        findings.append(
            Finding(
                WARNING,
                "Історія особи",
                where,
                f"РНОКПП відрізняється від указаного в наказі{reference}",
                f"У наказі{reference} для цієї особи: {previous.ipn}. Можлива описка або тезка.",
                (item.ipn, first_line),
            )
        )

    # Лише увага, ніколи не помилка: особу могли призначити наказом іншого органу,
    # якого в індексі немає, тож розбіжність з НАШИМ попереднім наказом нічого не доводить.
    expected = previous.target_position or previous.current_position
    if item.current_position and expected and not same_position(item.current_position, expected):
        state = "призначена" if previous.target_position else "займала посаду"
        findings.append(
            Finding(
                WARNING,
                "Історія особи",
                where,
                f"займана посада відрізняється від наказу{reference}",
                f"За наказом{reference} особа {state}: «{expected}». "
                f"У пункті: «{item.current_position}». Якщо після того особу призначали іншим "
                "наказом (зокрема не нашим) — усе гаразд; інакше перевірте посаду.",
                (item.current_position[:60], first_line),
            )
        )

    current_rank, previous_rank = rank_level(item.rank, levels), rank_level(previous.rank, levels)
    if current_rank is not None and previous_rank is not None and current_rank < previous_rank:
        findings.append(
            Finding(
                WARNING,
                "Історія особи",
                where,
                f"звання нижче, ніж у наказі{reference}",
                f"У наказі{reference}: «{previous.rank}», у пункті: «{item.rank}».",
                (first_line,),
            )
        )
    return findings


# ─────────────────────────────────────────────────────────────────────────────
def review_order(
    text: str,
    *,
    order_number: str = "",
    order_date: str = "",
    signer: dict | None = None,
    mapping: dict | None = None,
    index_folder: str | Path | None = None,
    check_spelling: bool = CHECK_POSITION_SPELLING,
    full_text: str | None = None,
) -> ReviewResult:
    """Усі перевірки наказу. `text` — уже без підписного хвоста.

    `full_text` — увесь документ разом із підписантом (для порожніх абзаців перед ним);
    якщо не передано — `text`.
    """
    result = ReviewResult()
    # № і дата в назві файлу НЕ перевіряються: проєкт наказу, який ще набирається, їх не
    # має — це нормально (користувач, 25.09.2026). `check_metadata` лишився для інших місць.
    result.findings += check_signer(signer)
    result.findings += check_numbering(text)
    result.findings += check_ipn(text)
    result.findings += check_bio_format(text)
    result.findings += check_doubles(text)
    result.findings += check_signer_spacing(full_text if full_text is not None else text)
    result.findings += check_routing(text, mapping)
    items = order_items(text)
    result.findings += check_nomenclature(text, items)
    result.findings += check_alphabet(text, items)
    has_index = bool(index_folder) and index_path(index_folder).exists()
    if has_index:
        result.findings += check_history(items, index_folder, order_number, order_date, result)
        result.notes.append(
            f"Індекс: {result.indexed_orders} наказів; осіб у наказі: {result.people_checked}, "
            f"знайдено в попередніх наказах: {result.people_found}."
        )
    else:
        result.notes.append(
            "Індекс попередніх наказів не знайдено — історію осіб не перевірено. "
            "Вкажіть папку індексу у вікні «🔒 Накази» або в індексаторі посад."
        )
    if check_spelling:
        result.findings += check_position_spelling(items, index_folder if has_index else None)
    return result


def default_index_folder(configured: str = "") -> str:
    """Папка індексу: з вікна «Накази», інакше — з окремого індексатора посад."""
    candidates = [configured]
    try:
        from PySide6.QtCore import QSettings

        candidates.append(
            str(
                QSettings("NodeAutomationToolkit", "OrderPositionIndexer").value("index_folder", "")
            )
        )
    except Exception:  # без Qt (тести, консоль) — лише налаштування вікна
        pass
    for folder in candidates:
        if folder and index_path(folder).exists():
            return folder
    return ""
