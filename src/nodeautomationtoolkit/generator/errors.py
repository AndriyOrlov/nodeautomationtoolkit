"""Зрозумілі користувачу помилки та знеособлення журналу.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import os
import re



class UserError(Exception):
    """Помилка, яку показуємо користувачу простими словами.

    `what` — що саме сталося, `todo` — що з цим робити. Текст складається
    так, щоб його зрозуміла людина, яка ніколи не відкривала журнал програми:
    без кодів, без англійських слів, з назвою конкретного файлу.
    """

    def __init__(self, what: str, todo: str = "", detail: str = ""):
        self.what = what.strip()
        self.todo = todo.strip()
        self.detail = detail.strip()
        super().__init__(self.what)

    def __str__(self) -> str:
        parts = [f"Що сталося: {self.what}"]
        if self.todo:
            parts.append(f"Що зробити: {self.todo}")
        if self.detail:
            parts.append(f"(технічні подробиці: {self.detail})")
        return "\n".join(parts)


# Коди Windows, які трапляються в цій програмі найчастіше. Windows часто
# повідомляє їх БЕЗ назви файлу («[WinError 2] The system cannot find the
# file specified»), тому саме тут вони перетворюються на зрозумілий текст.
_WINDOWS_DRIVE_ERRORS = (21, 53, 67, 1231, 1265)  # диск/мережева тека недоступні


_WINDOWS_BUSY_ERRORS = (32, 33)                   # файл тримає інша програма


def _error_file_name(error: BaseException) -> str:
    """Назва файлу з помилки, якщо ОС її повідомила."""
    for attribute in ("filename", "filename2"):
        value = getattr(error, attribute, None)
        if value:
            return str(value)
    return ""


def explain_error(error: BaseException) -> str:
    """Перекладає будь-який збій на просту мову: що сталося і що робити.

    Показується і у вікні, і в журналі. Технічний текст не викидається —
    він лишається в кінці, щоб було з чим прийти по допомогу.
    """
    if isinstance(error, UserError):
        return str(error)

    technical = f"{type(error).__name__}: {error}".strip()
    winerror = getattr(error, "winerror", None)
    name = _error_file_name(error)
    named = f"«{os.path.basename(name)}»" if name else "потрібний файл"
    where = f"\n{name}" if name else ""

    if isinstance(error, FileNotFoundError) or winerror in (2, 3):
        return str(UserError(
            f"комп'ютер не знайшов {named} — такого файлу на диску немає." + where,
            "Файл перейменували, перемістили або видалили. Відкрийте «Зразки та "
            "реквізити» й виберіть його заново кнопкою «Вибрати», а наказ — "
            "кнопкою вибору наказу в головному вікні.",
            technical,
        ))

    if isinstance(error, PermissionError) or winerror in _WINDOWS_BUSY_ERRORS:
        return str(UserError(
            f"файл {named} зараз зайнятий іншою програмою або закритий для запису." + where,
            "Закрийте цей файл у Word чи Excel (і перевірте, чи не відкритий він "
            "у вікні попереднього перегляду) і натисніть кнопку ще раз.",
            technical,
        ))

    if winerror in _WINDOWS_DRIVE_ERRORS:
        return str(UserError(
            "диск або мережева тека, де лежить потрібний файл, зараз недоступні." + where,
            "Перевірте, чи підключений диск (наприклад, E:) і чи є мережа, "
            "потім повторіть.",
            technical,
        ))

    if type(error).__name__ == "com_error":
        return str(UserError(
            "Word не виконав дію: він або зайнятий, або показує власне вікно "
            "(наприклад, запит на збереження), або його закрили під час роботи.",
            "Закрийте всі вікна Word, переконайтесь, що жоден документ не чекає "
            "відповіді, і повторіть. Якщо не допомогло — перезавантажте комп'ютер.",
            technical,
        ))

    return str(UserError(
        "сталася несподівана помилка, і програма зупинила саме цей файл.",
        "Спробуйте ще раз. Якщо повториться — збережіть текст із журналу "
        "(нижнє поле вікна) і покажіть його розробнику.",
        technical,
    ))


#: Слова, які у журналі стоять поруч із числом, але частинами НЕ є.
#: Без цього переліку знеособлення зʼїдало б корисну статистику
#: («2 стор.», «13 пунктів», «+39 ентерів») разом із назвами частин.
_LOG_MEASURE_WORDS = (
    "стор", "сторінк", "сторінка", "сторінки", "сторінок", "односторінк", "багатосторінк",
    "арк", "аркуш", "аркуша", "аркушів", "пункт", "пункту", "пунктів", "пунктами",
    "витяг", "витягів", "витягах", "витяги", "абзац", "абзаців", "ентер", "ентерів",
    "шт", "пт", "с", "сек", "секунд", "байт", "записів", "скорочень", "прим",
    "рядк", "рядків", "разів", "файл", "файлів", "з", "із", "та", "і", "на", "від",
)


#: Одне правило знеособлення на всі випадки: перший збіг виграє, і вже
#: підставлена мітка під наступні правила не потрапляє.
_LOG_NAME_WORD = r"[А-ЯІЇЄҐ][а-яіїєґ]+(?:['’\-][А-ЯІЇЄҐа-яіїєґ]+)*"


_LOG_SURNAME_UPPER = r"[А-ЯІЇЄҐ]{2,}(?:['’\-][А-ЯІЇЄҐ]+)*"


_REDACT_RE = re.compile(
    # Шлях закінчується перед « · »: далі в журналі стоять версія модуля
    # маршрутизації, дата зміни словника й розмір. Вони нічого не розкривають,
    # а без них зі знеособленого журналу не видно, чи програму перезапущено.
    r"(?P<path>[A-Za-zА-ЯІЇЄҐ]:[\\/][^\n·]*[^\n·\s])"
    r"|(?P<otck>(?i:[А-ЯІЇЄҐ][а-яіїєґ'’\-]+(?:ськ|цьк|зьк)[а-яіїєґ]*\s+"
    r"(?:обласн\w+\s+|районн\w+\s+|міськ\w+\s+)?(?:О?М?Р?ТЦК|територіальн\w+)"
    r"(?:[^\n,;.]{0,80}?підтримки)?(?:\s+та\s+СП)?))"
    rf"|(?P<name>\b(?:{_LOG_SURNAME_UPPER}[ \t]+{_LOG_NAME_WORD}(?:[ \t]+{_LOG_NAME_WORD})?"
    rf"|{_LOG_NAME_WORD}(?:[ \t]+{_LOG_NAME_WORD})?[ \t]+{_LOG_SURNAME_UPPER}"
    rf"|{_LOG_NAME_WORD}[ \t]+{_LOG_NAME_WORD}(?:[ \t]+{_LOG_NAME_WORD})?)\b)"
    r"|(?P<cipher>\b[АA]\s?\d{3,4}\b)"
    r"|(?P<unit>\b\d{1,3}\s?(?P<unit_word>[А-Яа-яІЇЄҐіїєґ]{1,12})\b)"
    r"|(?P<upper>\b[А-ЯІЇЄҐ]{3,}\b)",
    re.UNICODE,
)


#: ВЕЛИКІ слова, які в журналі є службовими, а не прізвищами.
_LOG_UPPER_KEEP = {
    "УВАГА", "ГЕНЕРАЦІЯ", "ВИТЯГІВ", "ВИТЯГ", "РОЗРАХУНОК", "РОЗСИЛКИ", "СТАТИСТИКА",
    "ТА", "ДРУКУ", "ОТЦК", "ТЦК", "РТЦК", "МТЦК", "ОМТЦК", "СП", "АК", "ВЧ", "ЗС",
    "ПОМИЛКА", "ЗБІЙ", "ГОТОВО", "DOCX", "XLSX", "COM", "ПІБ", "КОМУ", "КУДИ",
}


def redact_sensitive_text(text: str) -> str:
    """Знеособлює журнал: назви частин, шифри, ПІБ і шляхи — на стабільні мітки.

    Потрібно, щоб журнал збою можна було комусь показати. Структура (скільки
    витягів, скільки сторінок, які попередження, скільки часу) лишається
    незмінною — саме вона й потрібна для розбору. Мітки стабільні в межах
    одного копіювання: та сама частина всюди буде тим самим «Ч-1».
    """
    counters: dict[str, dict[str, str]] = {}

    def token(prefix: str, value: str) -> str:
        bucket = counters.setdefault(prefix, {})
        key = value.casefold().strip()
        if key not in bucket:
            bucket[key] = f"{prefix}-{len(bucket) + 1}"
        return bucket[key]

    def replace(match: re.Match) -> str:
        if match.lastgroup == "path":
            return "<шлях>"
        if match.lastgroup == "otck":
            return token("ОТЦК", match.group(0))
        if match.lastgroup == "cipher":
            return token("Ш", match.group(0))
        if match.lastgroup == "name":
            return token("ПІБ", match.group(0))
        if match.lastgroup == "unit":
            word = match.group("unit_word").casefold().rstrip(".")
            if word in _LOG_MEASURE_WORDS:
                return match.group(0)
            return token("Ч", match.group(0))
        word = match.group(0)
        return word if word in _LOG_UPPER_KEEP else token("ПІБ", word)

    # ОДИН прохід: інакше мітка, яка сама містить число («ОТЦК-1»), потрапляла
    # під наступне правило й перетворювалась на «ОТЦК-Ч-2».
    text = _REDACT_RE.sub(replace, text)
    # Ім'я з великої літери поруч із міткою ПІБ («полковник Петро ПІБ-1»).
    text = re.sub(r"\b[А-ЯІЇЄҐ][а-яіїєґ'’\-]+(?=\s+ПІБ-\d)", "<імʼя>", text)
    return text
