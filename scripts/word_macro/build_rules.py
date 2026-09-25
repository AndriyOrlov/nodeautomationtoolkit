"""Збирає `word_macro/nat_rules.txt` — правила для макросу Word з довідників програми.

Макрос (`word_macro/NATCheckCore.bas`) не має Python і не читає CSV програми, тому все,
що йому треба, лежить в одному UTF-8 файлі «ключ=значення»: звання з рівнями (вже в
основах слів), межі номенклатури, закінчення для основ слів, словник посад і тексти
повідомлень. Українські тексти — тут, а не в .bas: .bas Word імпортує в кодуванні
ANSI, і кирилиця в ньому ламається.

Запуск (після зміни довідників у `personnel/dictionaries`):
    python scripts/word_macro/build_rules.py
Тест `tests/test_word_macro.py` стежить, щоб файл відповідав довідникам.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from nodeautomationtoolkit.order_index import position_dictionary  # noqa: E402
from nodeautomationtoolkit.order_review import check  # noqa: E402

OUTPUT = ROOT / "word_macro" / "nat_rules.txt"

#: Тексти знахідок — ДОСЛІВНО як у `order_review/check.py` (тест звіряє результат).
MESSAGES = {
    "level.error": check.ERROR,
    "level.warning": check.WARNING,
    "where.item": "Пункт {0}",
    "where.item_none": "Пункт без номера",
    "where.before": "Текст до першого пункту",
    "num.rule": "Нумерація пунктів",
    "num.glued.what": "після номера пункту немає пропуску",
    "num.glued.how": "Пишіть «1. Капітана …», а не «1.Капітана …».",
    "num.dup.what": "номер пункту повторюється",
    "num.dup.how": "Після пункту {0} має йти {1}.",
    "num.gap.what": "після пункту {0} йде пункт {1}",
    "num.gap.how": "Очікувався пункт {0}. Нумерація наскрізна через увесь наказ.",
    "num.sub_parent.what": "підпункт стоїть не під пунктом {0}",
    "num.sub_parent.how": "Підпункт {0} має йти одразу після пункту {1}.",
    "num.sub_seq.what": "очікувався підпункт {0}.{1}",
    "num.sub_seq.how": "Підпункти нумеруються підряд з 1.",
    "ipn.rule": "РНОКПП",
    "ipn.digit.what": "РНОКПП не проходить контрольну перевірку",
    "ipn.digit.how": "Звірте цифри з документом особи — найчастіше це описка.",
    "ipn.sex.what": "стать у РНОКПП не збігається з «{0}»",
    "ipn.sex.how": "Перевірте РНОКПП і «Народився/Народилася».",
    "ipn.date.what": "дата народження {0} не збігається з РНОКПП",
    "ipn.year.what": "рік народження {0} не збігається з РНОКПП",
    "ipn.date.how": "Перевірка приблизна: звірте дату народження і РНОКПП з документом.",
    "nom.rule": "Номенклатура",
    "nom.over.what": "{0} «{1}» вище за {2} — поза номенклатурою {3}",
    "nom.naval.what": "{0} «{1}» — корабельне звання, а не Сухопутних військ",
    "nom.naval.how": "Перевірте звання: у Сухопутних військах лише військові звання.",
    "nom.label.person": "звання особи",
    "nom.label.quoted": "шпк / звання в лапках",
    "nom.action.звільнення": "звільнення",
    "nom.action.призначення": "призначення",
    "nom.action.": "наказу",
    "birth.rule": "Рік народження",
    "birth.good_form": "р.н.",
    "birth.full.what": "дата народження «{0}» — пишеться лише рік",
    "birth.full.how": "Пишіть рік без дня й місяця: «{0} р.н.,».",
    "birth.form.what": "«{0}» — не за зразком",
    "birth.form.how": "Пишіть «{0} р.н.,».",
    "service.rule": "Запис «у ЗС»",
    "service.what": "«{0}» — не за зразком",
    "service.how": check.SERVICE_SAMPLE,
    "doubles.rule": "Задвоєні знаки",
    "doubles.mark.what": "задвоєний знак «{0}»",
    "doubles.mark.how": "Приберіть зайвий знак.",
    "doubles.space.what": "два пробіли підряд",
    "doubles.space.how": "Залиште один пробіл.",
    "signer.rule": "Порожні абзаци",
    "signer.where": "Кінець наказу",
    "signer.what": "порожніх абзаців перед підписантом: {0} — має бути 2",
    "signer.how": "Між останнім пунктом і підписантом наказу — рівно два порожні абзаци.",
    "lineend.rule": "Кінець рядка",
    "lineend.where": "Сторінка {0}",
    "lineend.word.what": "«{0}» у кінці рядка",
    "lineend.number.what": "число «{0}» у кінці рядка",
    "lineend.how": "Поставте нерозривний пробіл (Ctrl+Shift+Пробіл) після «{0}» — тоді воно перейде на "
    "наступний рядок разом із наступним словом. На початку рядка — можна.",
    "alpha.rule": "Алфавітний порядок",
    "alpha.what": "«{0}» стоїть після «{1}» — порушено абетку",
    "alpha.how": "Пункти однієї групи (під однією шапкою) йдуть за прізвищами за абеткою.",
    "spell.rule": "Написання посади",
    "spell.what": "у посаді «{0}» незнайоме слово «{1}»",
    "spell.how_suggest": "Можливо, «{0}».",
    "spell.how_none": "Такого слова немає в довіднику посад — перевірте за штатом.",
    # Службові слова для розбору (не повідомлення, але теж українською — тому тут).
    "born.male": "народився",
    "action.dismissal": "звільнення",
    "action.appointment": "призначення",
    "action.appoint_verb": "призначити",
    "action.any": "усі",
    "verb.ending": "ТИ",
    "not_position_start": "ДО У В НА ІЗ З ЗА НАЧАЛЬНИКУ",
    "done.title": "Перевірка наказу",
    "done.summary": "Помилок (червоне): {0}\nЗауважень (жовте): {1}",
    "done.clean": "Нічого не знайдено.",
    "done.note": "Примітки й виділення додано в документ. Документ НЕ збережено — збережіть або закрийте без збереження.",
    "err.rules": "Не знайдено файл правил nat_rules.txt. Вкажіть його.",
    "err.program": "Не знайдено програму генератора в теці: {0}",
    "err.run": "Перевірка програмою не вдалася (код {0}).",
    "ask.program": "Оберіть теку програми генератора (де лежить generate_extracts_qt.py або GeneratorVytyagivQt.exe).",
}

#: Регулярні вирази — підмножина, яку розуміє VBScript.RegExp (без lookbehind і \b біля кирилиці).
LETTERS = "A-Za-zА-ЯЁІЇЄҐа-яёіїєґ"
UP, LO = "А-ЯІЇЄҐ", "а-яіїєґ"
APOS = "'’ʼ`"
NBSP = "\u00a0"  # у VBScript \s нерозривного пробілу не бачить, у Python — бачить
#: Короткі слова, які не лишаються в кінці рядка (макрос; користувач 25.09.2026).
#: «про», «та», «як» — винятки: їм можна стояти в кінці рядка.
SHORT_WORDS = [
    "а", "в", "у", "і", "й", "о", "з", "зі", "із", "до", "за", "на", "не", "ні", "по",
    "чи", "що", "від", "для", "під", "над", "без", "при", "№",
]
PATTERNS = {
    "word": f"[{LETTERS}]+(?:['’ʼ-][{LETTERS}]+)*|[0-9]+",
    # (^|не-літера) + ПРІЗВИЩЕ Ім'я По-батькові,
    "person": f"(^|[^0-9_{LETTERS}{APOS}-])([{UP}][{UP}{APOS}-]+\\s+[{UP}][{LO}{APOS}-]+\\s+[{UP}][{LO}{APOS}-]+\\s*,\\s*)",
    "item_label": "^(?:Пункт\\s+)?([0-9]{1,3}(?:\\.[0-9]{1,3})*)[.)]\\s+",
    "item": "^\\s*([0-9]{1,3}(?:\\.[0-9]{1,3})*)[.)]\\s+\\S",
    "glued": f"^\\s*([0-9]{{1,3}}(?:\\.[0-9]{{1,3}})*)[.)](?=[{LETTERS}]|[«\"“'])",
    "ipn_item": "^\\s*([0-9]{1,3})[.)]\\s",
    "ipn_token": "(^|[^0-9.,/+-])([1-9][0-9]{9})(?![0-9])(?![.,/-][0-9])",
    "birth_year": f"(^|[^0-9_{LETTERS}])(19[0-9]{{2}}|20[0-9]{{2}})\\s*р\\.?\\s*н\\.?",
    "birth_numeric": f"(^|[^0-9_{LETTERS}])([0-9]{{1,2}})\\.([0-9]{{1,2}})\\.((?:19|20)[0-9]{{2}})\\s*(?:р\\.?\\s*н\\.?|року\\s+народження)",
    "born": f"(^|[^0-9_{LETTERS}])(народився|народилася)\\s+(?:([0-9]{{1,2}})\\s+([{LO}]+)\\s+((?:19|20)[0-9]{{2}})|([0-9]{{1,2}})\\.([0-9]{{1,2}})\\.((?:19|20)[0-9]{{2}}))",
    "heading": f"^(?:наказ|затвердж|додаток|відповідно\\s+до|у\\s+зв['’]язку|про\\s+[{LO}])",
    "hint": "відповідно\\s+до|згідно\\s+з|звільнити|призначити|присвоїти|:\\s*$",
    "dismissal": "звільнити\\s+з\\s+військової\\s+служби",
    "quoted": "[«\"“]([^»\"”]{3,40})[»\"”]",
    "section": "^\\s*§\\s*[0-9]+",
    # Звання перед ПІБ, що закінчується одним Словом з великої — лишається саме воно
    # (як `order_index.items._person_parts`).
    "last_cap_word": f"(^|[^{LO}])([{UP}][{LO}]+)\\s*$",
    # Рік народження: 1 префікс, 2 «дд.мм.», 3 дд, 4 мм, 5 рік, 6 пробіл, 7 «р.н.» (на LCase).
    "birth_mention": f"(^|[^0-9_{LETTERS}.])(([0-9]{{1,2}})\\.([0-9]{{1,2}})\\.)?((?:19|20)[0-9]{{2}})"
    f"([\\s{NBSP}]*)(р\\.?\\s*н\\.?|року\\s+народження)",
    # Після «р.н.»: кома або табуляції + РНОКПП + крапка (як check._AFTER_BIRTH_RE).
    "after_birth": f",|[\\t {NBSP}]+[0-9]{{10}}\\.",
    # «у ЗС» / «у ЗСУ» як слово, але не «Вислуга років у ЗС: …».
    "service_start": f"(^|[^0-9_{LETTERS}])([уУ][\\s{NBSP}]*ЗСУ?)(?![0-9_{LETTERS}])(?![\\s{NBSP}]*:)",
    "service_ok": check._SERVICE_RE.pattern.replace("\\s", f"[\\s{NBSP}]"),
    "sentence_end": f"\\.(?=[\\s{NBSP}]|$)",
    # Підписант наказу (як typography.ORDER_SIGNER_START_RE), на LCase рядка.
    "signer": f"^\\s*(?:(?:т\\s*\\.\\s*)?в\\s*\\.\\s*о\\s*\\.?(?=\\s)|тимчасово\\s+виконуюч(?:ий|а)?(?![0-9_{LETTERS}])"
    f"|(?:командувач|командир|начальник|заступник\\s+командувача)(?![0-9_{LETTERS}]))",
    "caps_word":f"(^|[^0-9_{LETTERS}{APOS}-])([{UP}][{UP}{APOS}-]{{4,}})(?![0-9_{LETTERS}{APOS}-])",
}
CONTENT_MARKERS = ["НАКАЗУЮ", "ПРИЗНАЧИТИ", "НАПРАВИТИ", "ВІДРЯДИТИ", "ЗВІЛЬНИТИ", "ВІЙСЬКОВОСЛУЖБОВЦІВ"]
NAVAL = ["матрос", "старшин", "рангу", "капітан-лейтенант", "коммодор", "адмірал"]
MONTHS = ["січня", "лютого", "березня", "квітня", "травня", "червня",
          "липня", "серпня", "вересня", "жовтня", "листопада", "грудня"]


def build() -> str:
    lines = [
        "# Правила для макросу Word «Перевірка наказу» (NAT). Згенеровано з довідників програми —",
        "# не правте вручну: змініть personnel/dictionaries і запустіть scripts/word_macro/build_rules.py.",
        f"endings={'|'.join(sorted(position_dictionary._ENDINGS, key=lambda e: (-len(e), e)))}",
        # Описки в назвах посад: 1 — перевіряти (як CHECK_POSITION_SPELLING у програмі).
        f"check.spelling={int(check.CHECK_POSITION_SPELLING)}",
        # Корабельні звання: у макросі не перевіряються (лише Сухопутні війська, 25.09.2026);
        # у програмі — лишаються жовтою позначкою.
        "check.naval=0",
        f"alphabet={check.UKRAINIAN_ALPHABET}",
        f"short_words={'|'.join(SHORT_WORDS)}",
        f"content_markers={'|'.join(CONTENT_MARKERS)}",
        f"naval={'|'.join(NAVAL)}",
    ]
    lines += [f"month={name}|{number}" for number, name in enumerate(MONTHS, start=1)]
    lines += [f"pattern.{key}={value}" for key, value in PATTERNS.items()]
    lines += [f"msg.{key}={value.replace(chr(10), '\\n')}" for key, value in MESSAGES.items()]
    for stems, level in check._rank_levels():
        lines.append(f"rank={' '.join(stems)}|{level}")
    for action, (level, name, naval, basis) in check.load_nomenclature().items():
        lines.append(f"limit={action}|{'' if level is None else level}|{name}|{int(naval)}|{basis}")
    stems, words = check.position_vocabulary(None)
    lines += [f"v={item}" for item in sorted(stems) if item]
    lines += [f"w={item}" for item in sorted(words) if item]
    return "\n".join(lines) + "\n"


def main() -> int:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(build(), encoding="utf-8", newline="\n")
    print(f"{OUTPUT} — {OUTPUT.stat().st_size} байт")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
