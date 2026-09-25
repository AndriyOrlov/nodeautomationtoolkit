"""Макрос Word «Перевірка наказу»: ядро VBA дає те саме, що й програма.

Ядро `word_macro/NATCheckCore.bas` (між мітками NAT CORE) написане спільною
підмножиною VBA і VBScript, тож тут воно виконується через `cscript` на вигаданих
наказах, а результат звіряється з `order_review.review_order`. Word не потрібен.

Усі особи й РНОКПП вигадані.
"""

import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
MACRO_DIR = PROJECT_ROOT / "word_macro"
sys.path.insert(0, str(PROJECT_ROOT / "scripts" / "word_macro"))

from nodeautomationtoolkit.order_review import review_order  # noqa: E402

CSCRIPT = shutil.which("cscript")
needs_cscript = pytest.mark.skipif(CSCRIPT is None, reason="cscript є лише у Windows")

#: Правила, які макрос повторює за програмою один в один.
SHARED_RULES = {
    "Нумерація пунктів", "РНОКПП", "Номенклатура", "Алфавітний порядок",
    "Рік народження", "Запис «у ЗС»", "Задвоєні знаки", "Порожні абзаци",
}

HARNESS = """
Dim gRules, gText, gFindings
Set gRules = NAT_LoadRules(WScript.Arguments(0))
gText = NAT_ReadUtf8(WScript.Arguments(1))
Set gFindings = NAT_Review(gRules, gText)
NAT_WriteUtf8 WScript.Arguments(2), NAT_FindingsText(gFindings)
"""


def _ipn(birth: date, serial: int = 123, male: bool = True) -> str:
    days = (birth - date(1899, 12, 31)).days
    digits = f"{days:05d}{serial:03d}{1 if male else 2}"
    weights = (-1, 5, 7, 9, 4, 6, 10, 5, 7)
    return digits + str(sum(int(d) * w for d, w in zip(digits, weights, strict=True)) % 11 % 10)


IPN = _ipn(date(1988, 3, 1))
IPN_WOMAN = _ipn(date(1990, 5, 17), 456, male=False)
BAD_IPN = IPN[:-1] + str((int(IPN[-1]) + 1) % 10)


def _core() -> str:
    text = (MACRO_DIR / "NATCheckCore.bas").read_text(encoding="ascii")
    start = text.index("'=== NAT CORE BEGIN ===")
    end = text.index("'=== NAT CORE END ===")
    return text[start:end]


def _run_macro(tmp_path: Path, order_text: str, spelling: bool = False) -> list[tuple]:
    script = tmp_path / "core.vbs"
    script.write_text(_core() + HARNESS, encoding="ascii")
    rules = tmp_path / "nat_rules.txt"
    rules_text = (MACRO_DIR / "nat_rules.txt").read_text(encoding="utf-8")
    if spelling:
        rules_text = rules_text.replace("check.spelling=0", "check.spelling=1")
    rules.write_text(rules_text, encoding="utf-8")
    source = tmp_path / "in.txt"
    source.write_text(order_text, encoding="utf-8")
    output = tmp_path / "out.txt"
    completed = subprocess.run(
        [CSCRIPT, "//nologo", "//E:vbscript", str(script), str(rules), str(source), str(output)],
        capture_output=True, text=True, timeout=120,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr
    lines = output.read_text(encoding="utf-8-sig").splitlines()
    return [tuple(line.split("\t")) for line in lines if line.strip()]


def _program(order_text: str) -> list[tuple]:
    result = review_order(order_text, order_number="30", order_date="01.09.2026", signer=None)
    return [
        (f.level, f.rule, f.where, f.what, f.how, (f.quotes or ("",))[0])
        for f in result.findings
        # Корабельні звання макрос не перевіряє (check.naval=0) — лише програма.
        if f.rule in SHARED_RULES and "корабельне звання" not in f.what
    ]


def _shared(macro: list[tuple]) -> list[tuple]:
    return [row[:6] for row in macro if row[1] in SHARED_RULES]


HEAD = [
    "НАКАЗ",
    "§ 1",
    "Відповідно до пункту 1 Положення нижчепойменованих офіцерів ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ:",
]
DISMISSAL = [
    "§ 2",
    "Відповідно до підпункту «б» пункту 2 частини п'ятої статті 26 Закону України "
    "нижчепойменованих офіцерів ЗВІЛЬНИТИ З ВІЙСЬКОВОЇ СЛУЖБИ:",
]

CASES = {
    "numbering": HEAD + [
        "1. Капітана АЛЬФЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "1.1. Капітана БЕТЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "1.3. Капітана ГАМЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "2. Капітана ДЕЛЬТЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "2. Капітана ЕПСЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "4. Капітана ЗЕТЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "5.Капітана ЕТЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "§ 2",
        "Відповідно до пункту 2 Положення ПРИЗНАЧИТИ:",
        "1. Капітана ТЕТЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "3. Капітана ЙОТЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
    ],
    "ipn": HEAD + [
        "1. Капітана АЛЬФЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "02.03.1988 р.н., освіта: вища.",
        f"{IPN}.",
        "2. Капітана БЕТЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "1989 р.н.",
        f"{IPN}.",
        "3. Капітана ГАМЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "Народився 1 березня 1988 року.",
        f"{BAD_IPN}.",
        "4. Капітана ДЕЛЬТЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "Народилася 17 травня 1990 року.",
        f"{IPN_WOMAN}.",
        "5. Капітана ЕПСЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "Народилася 1 березня 1988 року.",
        f"{IPN}.",
        "Виконавець тел. 0501234567",
    ],
    "nomenclature": HEAD + [
        "1. Підполковника АЛЬФЕНКА Івана Івановича, заступника командира бригади - КОМАНДИРОМ БРИГАДИ.",
        "Призначається на вищу посаду з шпк «підполковник» на шпк «полковник».",
        "2. Полковника БЕТЕНКА Івана Івановича, командира бригади - КОМАНДИРОМ КОРПУСУ.",
        "3. Капітана 2 рангу ГАМЕНКА Івана Івановича, командира корабля - КОМАНДИРОМ ДИВІЗІОНУ.",
        "4. Старшого лейтенанта ДЕЛЬТЕНКА Івана Івановича, командира взводу - КОМАНДИРОМ РОТИ.",
    ] + DISMISSAL + [
        "5. Бригадного генерала ЕПСЕНКА Івана Івановича, командира корпусу.",
        "6. Майора ЗЕТЕНКА Івана Івановича, начальника штабу батальйону.",
        "7. Генерал-майора ЕТЕНКА Івана Івановича, начальника управління.",
    ],
    "alphabet": HEAD + [
        "1. Капітана БОНДАРЯ Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "1985 р.н.",
        "2. Капітана АНДРІЄНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "3. Капітана ҐУДЗЯ Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "4. Капітана ЄВЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "5. Капітана КОВАЛЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "6. Капітана КОВАЛЯ Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "7. Капітана КОВАЛЯ Андрія Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "7.1. Капітана АБРАМЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "8. Капітана ЯРЕМЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "9. Капітана ІВАНЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
    ] + DISMISSAL + [
        "У ЗАПАС ЗА ПІДПУНКТОМ «а»:",
        "10. Майора ЯЦЕНКА Івана Івановича, начальника штабу батальйону.",
        "У ЗАПАС ЗА ПІДПУНКТОМ «б»:",
        "11. Майора АБРАМЕНКА Івана Івановича, начальника штабу батальйону.",
    ],
    "bio": HEAD + [
        "1. Капітана АЛЬФЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "1978 р.н., освіта: вища, у ЗС – із 07.1991.",
        "2. Капітана БЕТЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "02.03.1988 р.н., освіта: вища, у ЗС - із 06.1997 по 06.2010 та з 01.2017.",
        "3. Капітана ГАМЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "1979 р.н.\t\t\t\t\t2900000000.",
        "4. Капітана ДЕЛЬТЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "1974р.н., у ЗС з 07.1991.",
        "5. Капітана ЕПСЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "1974 р. н., у ЗСУ – із 1991.",
        "6. Капітана ЗЕТЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "1975 р.н.",
        "1976\u00a0р.н., у\u00a0ЗС –\u00a0із 06.1997 по 06.2010, з 03.2012 по 04.2014 та з 01.2017.",
        "7. Капітана ЕТЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "Народився 11 серпня 1976 року. Вислуга років у ЗС: календарна - 20 років.",
    ],
    "doubles and signer": HEAD + [
        "1. Капітана АЛЬФЕНКА Івана Івановича,, командира роти.. -  КОМАНДИРОМ БАТАЛЬЙОНУ...",
        "1978 р.н., освіта: вища;; у ЗС – із 07.1991!!",
        "",
        "Командир військової частини А0001",
        "полковник                    Петро ТЕСТОВИЙ",
    ],
    "broken paragraph": DISMISSAL + [
        "1. Полковника АЛЬФЕНКА Івана Івановича, командира",
        "бригади 901 механізованого батальйону.",
    ],
}


@needs_cscript
@pytest.mark.parametrize("case", sorted(CASES))
def test_macro_matches_the_program(tmp_path, case):
    text = "\n".join(CASES[case])
    expected = _program(text)
    assert expected, "кейс має давати знахідки, інакше порівнювати нічого"
    assert _shared(_run_macro(tmp_path, text)) == expected


@needs_cscript
def test_macro_finds_misspelled_positions_and_keeps_correct_ones_clean(tmp_path):
    text = "\n".join(HEAD + [
        "1. Капітана АЛЬФЕНКА Івана Івановича, командира механізованої роти - "
        "ЗАСТУПНИКОМ КОМАНДИРРА 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ.",
        "2. Капітана БЕТЕНКА Івана Івановича, командира механізованої роти - "
        "НАЧАЛЬНИКОМ ШТАБУУ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ.",
        "3. Майора ГАМЕНКА Івана Івановича, старшого офіцера відділу кадрів - "
        "СТАРШИМ ОФІЦЕРОМ ВІДДІЛУ КАДРІВ.",
        "4. Капітана ДЕЛЬТЕНКА Івана Івановича, начальника служби - ЗАСТУПНИКОМ КОМАНДИРА БАТАЛЬЙОНУ.",
    ])
    spelling = [row for row in _run_macro(tmp_path, text, spelling=True) if row[1] == "Написання посади"]
    assert [(row[2], row[3].rsplit(" ", 1)[-1], row[4]) for row in spelling] == [
        ("Пункт 1", "«командирра»", "Можливо, «командира»."),
        ("Пункт 2", "«штабуу»", "Можливо, «штабу»."),
    ]
    # Вимкнено за замовчуванням (check.spelling=0): на реальних наказах забагато жовтого.
    assert [row for row in _run_macro(tmp_path, text) if row[1] == "Написання посади"] == []


def test_macro_source_is_ascii_only():
    """Word імпортує .bas у кодуванні ANSI — кирилиця там ламається."""
    for path in MACRO_DIR.glob("*.bas"):
        path.read_bytes().decode("ascii")


def test_rules_file_is_up_to_date():
    import build_rules

    assert (MACRO_DIR / "nat_rules.txt").read_text(encoding="utf-8") == build_rules.build(), (
        "Довідники змінились — запустіть scripts/word_macro/build_rules.py"
    )


def test_program_answers_the_macro_like_the_check_button(tmp_path):
    """Макрос «Перевірити програмою» запускає generate_extracts_qt.py --review-text."""
    source = tmp_path / "nat_review_in.txt"
    source.write_text("\n".join(CASES["numbering"]), encoding="utf-8-sig")  # ADODB пише з BOM
    output = tmp_path / "nat_review_out.tsv"
    config = tmp_path / "config.json"
    config.write_text("{}", encoding="utf-8")
    completed = subprocess.run(
        [
            sys.executable, str(PROJECT_ROOT / "generate_extracts_qt.py"),
            "--review-text", str(source), "--order-name", "Наказ № 30 від 01.09.2026.docx",
            "--out", str(output), "--config", str(config), "--index", "",
        ],
        capture_output=True, text=True, timeout=180, cwd=tmp_path,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr

    lines = output.read_text(encoding="utf-8").splitlines()
    notes = [line for line in lines if line.startswith("#note\t")]
    findings = [tuple(line.split("\t")) for line in lines if line and not line.startswith("#")]
    assert any("Таблицю відповідностей не вибрано" in note for note in notes)
    assert all(len(row) == 7 for row in findings)
    # Той самий результат, що дає ядро макросу для правил, спільних з ним.
    assert _shared(findings) == _program("\n".join(CASES["numbering"]))


@needs_cscript
def test_macro_does_not_flag_naval_ranks(tmp_path):
    """У макросі корабельні звання не перевіряються — лише Сухопутні війська (25.09.2026)."""
    text = "\n".join(DISMISSAL + ["1. Капітана 2 рангу АЛЬФЕНКА Івана Івановича, командира корабля."])
    assert _shared(_run_macro(tmp_path, text)) == []


@needs_cscript
@pytest.mark.parametrize("blank_lines", [1, 2, 3])
def test_macro_counts_empty_paragraphs_before_signer(tmp_path, blank_lines):
    text = "\n".join(HEAD + [
        "1. Капітана АЛЬФЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "1978 р.н., освіта: вища.",
    ] + [""] * blank_lines + ["Командир військової частини А0001", "полковник   Петро ТЕСТОВИЙ"])
    expected = _program(text)
    assert _shared(_run_macro(tmp_path, text)) == expected
    assert bool(expected) == (blank_lines != 2)
