"""Перевірка наказу, складеного людиною: оформлення + історія осіб з індексу.

Усі особи, частини й РНОКПП вигадані (публічний репозиторій — без реальних даних).
"""

from datetime import date

import pytest

from nodeautomationtoolkit.order_index import store as index_store
from nodeautomationtoolkit.order_index.word_reader import WordTextReader
from nodeautomationtoolkit.order_review import ERROR, WARNING, review_order
from nodeautomationtoolkit.order_review import check as review_check


class _NoWordReader(WordTextReader):
    def _read_with_word(self, copy):
        raise OSError("Word у тестах не запускається")


def _ipn(birth: date, serial: int = 123, male: bool = True) -> str:
    """Вигаданий РНОКПП із правильною контрольною цифрою."""
    days = (birth - date(1899, 12, 31)).days
    digits = f"{days:05d}{serial:03d}{1 if male else 2}"
    weights = (-1, 5, 7, 9, 4, 6, 10, 5, 7)
    control = sum(int(d) * w for d, w in zip(digits, weights, strict=True)) % 11 % 10
    return digits + str(control)


IPN_A = _ipn(date(1988, 3, 1))
IPN_B = _ipn(date(1985, 6, 2), 456)

PREVIOUS_ORDER = [
    "НАКАЗ",
    "§ 1",
    "Відповідно до пункту 1 Положення нижчепойменованих офіцерів ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ:",
    "1. Капітана ТЕСТЕНКА Олега Васильовича, командира механізованого взводу 901 механізованого "
    "батальйону - КОМАНДИРОМ МЕХАНІЗОВАНОЇ РОТИ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ, ВОС - 0210003.",
    "1988 р.н., освіта: вища.",
    f"{IPN_A}.",
    "2. Майора ПРИКЛАДЕНКА Івана Петровича, начальника служби 901 механізованого батальйону - "
    "ЗАСТУПНИКОМ КОМАНДИРА 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ, ВОС - 0210003.",
    "1985 р.н., освіта: вища.",
    f"{IPN_B}.",
]


def _write_docx(path, paragraphs):
    from docx import Document

    document = Document()
    for text in paragraphs:
        document.add_paragraph(text)
    document.save(path)


@pytest.fixture
def index(tmp_path, monkeypatch):
    monkeypatch.setattr(index_store, "WordTextReader", _NoWordReader)
    orders, folder = tmp_path / "orders", tmp_path / "index"
    orders.mkdir()
    _write_docx(orders / "Наказ №12 від 05.03.2026.docx", PREVIOUS_ORDER)
    index_store.build_index(orders, folder)
    return folder


def _order(*items: str) -> str:
    return "\n".join(
        [
            "НАКАЗ",
            "§ 1",
            "Відповідно до пункту 1 Положення нижчепойменованих офіцерів ЗВІЛЬНИТИ з займаних "
            "посад і ПРИЗНАЧИТИ:",
            *items,
        ]
    )


def _review(text, **kwargs):
    kwargs.setdefault("order_number", "30")
    kwargs.setdefault("order_date", "01.09.2026")
    kwargs.setdefault(
        "signer", {"position": "Командир", "rank": "полковник", "name": "Петро ТЕСТОВИЙ"}
    )
    return review_order(text, **kwargs)


def _by_rule(result, rule):
    return [(f.level, f.where, f.what) for f in result.findings if f.rule == rule]


# ── Історія особи ───────────────────────────────────────────────────────────
def test_person_moved_from_the_position_given_last_time_is_fine(index):
    result = _review(
        _order(
            "1. Капітана ТЕСТЕНКА Олега Васильовича, командира механізованої роти 901 механізованого "
            "батальйону - НАЧАЛЬНИКОМ ШТАБУ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ, ВОС - 0210003.",
            "1988 р.н., освіта: вища.",
            f"{IPN_A}.",
        ),
        index_folder=index,
    )
    assert _by_rule(result, "Історія особи") == []
    assert (result.people_checked, result.people_found) == (1, 1)


def test_wrong_current_position_and_lower_rank_are_only_warnings(index):
    """Особу могли призначити не нашим наказом — тому лише жовте, не червоне."""
    result = _review(
        _order(
            "1. Капітана ПРИКЛАДЕНКА Івана Петровича, начальника служби 901 механізованого батальйону - "
            "НАЧАЛЬНИКОМ ШТАБУ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ, ВОС - 0210003.",
            "1985 р.н., освіта: вища.",
            f"{IPN_B}.",
        ),
        index_folder=index,
    )
    history = _by_rule(result, "Історія особи")
    assert history == [
        (WARNING, "Пункт 1", "займана посада відрізняється від наказу №12 від 05.03.2026"),
        (WARNING, "Пункт 1", "звання нижче, ніж у наказі №12 від 05.03.2026"),
    ]
    finding = next(f for f in result.findings if f.what.startswith("займана посада"))
    assert "ЗАСТУПНИКОМ КОМАНДИРА" in finding.how
    assert "не нашим" in finding.how


def test_ipn_of_another_person_is_an_error(index):
    result = _review(
        _order(
            "1. Лейтенанта ЗРАЗКОВА Петра Івановича, командира взводу 901 механізованого батальйону - "
            "КОМАНДИРОМ РОТИ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ.",
            "1988 р.н.",
            f"{IPN_A}.",
        ),
        index_folder=index,
    )
    assert (
        ERROR,
        "Пункт 1",
        "цей РНОКПП у наказі №12 від 05.03.2026 належить іншій особі",
    ) in _by_rule(result, "Історія особи")


def test_later_or_same_orders_are_not_history(index):
    """Наказ, датований пізніше за перевірюваний, не є «попереднім»."""
    result = _review(
        _order(
            "1. Капітана ПРИКЛАДЕНКА Івана Петровича, начальника служби 901 механізованого батальйону - "
            "НАЧАЛЬНИКОМ ШТАБУ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ.",
            f"{IPN_B}.",
        ),
        index_folder=index,
        order_date="01.02.2026",
    )
    assert _by_rule(result, "Історія особи") == []
    assert result.people_found == 0


@pytest.mark.parametrize(
    ("target", "word", "suggestion"),
    [
        ("ЗАСТУПНИКОМ КОМАНДИРРА 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ", "командирра", "командир"),
        ("НАЧАЛЬНИКОМ ШТАБУУ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ", "штабуу", "штабу"),
    ],
)
def test_misspelled_word_in_position_is_reported(target, word, suggestion):
    """Без індексу — звірка з довідником посад; кожне слово назви окремо."""
    text = _order(
        "1. Капітана ТЕСТЕНКА Олега Васильовича, командира механізованої роти 901 механізованого "
        f"батальйону - {target}.",
    )
    spelling = [f for f in _review(text, check_spelling=True).findings if f.rule == "Написання посади"]
    assert [f.what.rsplit(" ", 1)[-1] for f in spelling] == [f"«{word}»"]
    assert suggestion in spelling[0].how
    # Вимкнено за замовчуванням: на реальних наказах забагато жовтого (24.09.2026).
    assert [f for f in _review(text).findings if f.rule == "Написання посади"] == []


def test_correct_positions_are_not_reported(index):
    result = _review(
        _order(
            "1. Капітана ТЕСТЕНКА Олега Васильовича, командира механізованої роти 901 механізованого "
            "батальйону - ЗАСТУПНИКОМ КОМАНДИРА 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ.",
            "2. Майора ПРИКЛАДЕНКА Івана Петровича, начальника служби 901 механізованого батальйону - "
            "СТАРШИМ ОФІЦЕРОМ ВІДДІЛУ КАДРІВ.",
        ),
        index_folder=index,
        check_spelling=True,
    )
    assert [f for f in result.findings if f.rule == "Написання посади"] == []


def test_word_seen_once_in_index_is_not_learned(tmp_path, monkeypatch):
    """Помилка, зроблена в одному наказі, не стає нормою; повторена — стає."""
    from nodeautomationtoolkit.order_review.check import position_vocabulary

    monkeypatch.setattr(index_store, "WordTextReader", _NoWordReader)
    orders, folder = tmp_path / "orders", tmp_path / "index"
    orders.mkdir()
    item = "1. Капітана ТЕСТЕНКА Олега Васильовича, командира роти - ЗЕЛЕНОПОЛЬНИКОМ."
    learned = review_check.stem("зеленопольник")
    _write_docx(orders / "Наказ №1 від 01.01.2026.docx", ["§ 1", item])
    index_store.build_index(orders, folder)
    assert learned not in position_vocabulary(folder)[0]

    _write_docx(orders / "Наказ №2 від 02.01.2026.docx", ["§ 1", item])
    index_store.build_index(orders, folder)
    assert learned in position_vocabulary(folder)[0]


def test_without_index_only_a_note_is_added(tmp_path):
    result = _review(
        _order("1. Капітана ТЕСТЕНКА Олега Васильовича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ."),
        index_folder=tmp_path / "немає",
    )
    assert _by_rule(result, "Історія особи") == []
    assert any("Індекс попередніх наказів не знайдено" in note for note in result.notes)


# ── Оформлення ─────────────────────────────────────────────────────────────
def test_numbering_gaps_duplicates_subitems_and_glued_numbers():
    text = _order(
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
    )
    numbering = _by_rule(_review(text), "Нумерація пунктів")
    assert numbering == [
        (ERROR, "Пункт 1.3", "очікувався підпункт 1.2"),
        (ERROR, "Пункт 2", "номер пункту повторюється"),
        (ERROR, "Пункт 4", "після пункту 2 йде пункт 4"),
        (WARNING, "Пункт 5", "після номера пункту немає пропуску"),
        # Нумерація наскрізна: у § 2 знову «1» — помилка (25.09.2026).
        (ERROR, "Пункт 1", "після пункту 5 йде пункт 1"),
    ]


def test_continuous_numbering_across_sections_is_fine():
    text = _order(
        "1. Капітана АЛЬФЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "§ 2",
        "Відповідно до пункту 2 Положення ПРИЗНАЧИТИ:",
        "2. Капітана БЕТЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ.",
        "1988 р.н., 01.09.2026 року.",
    )
    assert _by_rule(_review(text), "Нумерація пунктів") == []


def test_filename_signer_ipn_and_routing():
    mapping = {
        "55 окремий батальйон": {
            "open_name": "55 окремий батальйон",
            "cipher": "А0055",
            "abbreviation": "55 об",
            "corps": "",
            "recipient_to": "Командиру військової частини А0055",
            "destination_where": "м. Тестове",
        }
    }
    bad_ipn = IPN_A[:-1] + str((int(IPN_A[-1]) + 1) % 10)
    text = _order(
        "1. Капітана АЛЬФЕНКА Івана Івановича, офіцера 55 окремого батальйону - ОФІЦЕРОМ 55 ОКРЕМОГО БАТАЛЬЙОНУ.",
        "1988 р.н.",
        f"{bad_ipn}.",
        "2. Капітана БЕТЕНКА Івана Івановича, офіцера відділу кадрів - СТАРШИМ ОФІЦЕРОМ ВІДДІЛУ КАДРІВ.",
    )
    result = review_order(text, order_number="", order_date="", signer={}, mapping=mapping)
    # Проєкт наказу ще без № і дати в назві файлу — це нормально, не знахідка.
    assert _by_rule(result, "Реквізити в назві файлу") == []
    assert _by_rule(result, "Підписант") == [
        (WARNING, "Кінець наказу", "блок підписанта не розпізнано")
    ]
    assert _by_rule(result, "РНОКПП") == [
        (ERROR, "Пункт 1", "РНОКПП не проходить контрольну перевірку")
    ]
    assert [where for _level, where, _what in _by_rule(result, "Адресат")] == ["Пункт 2"]


def test_grouped_summary_is_short():
    text = _order(
        *(
            f"{n}. Капітана АЛЬФЕНКА Івана Івановича, офіцера - ОФІЦЕРОМ."
            for n in (1, 3, 5, 7, 9, 11, 13, 15)
        )
    )
    lines = review_check.group_findings(_review(text).findings)
    assert len(lines) == 1
    assert lines[0].startswith("✖ Нумерація пунктів: 7 — Пункт 3, Пункт 5")
    assert lines[0].endswith("… ще 1")


@pytest.mark.parametrize(
    ("rank", "level"),
    [
        ("Капітана", 14),
        ("Старшого лейтенанта", 13),
        ("ПІДПОЛКОВНИКОМ", 16),
        ("головного сержанта", 6),
        ("", None),
    ],
)
def test_rank_level_in_any_case(rank, level):
    assert review_check.rank_level(rank) == level


# ── РНОКПП: дата народження й стать — жовті, контрольна цифра — червона ─────
@pytest.mark.parametrize(
    ("bio", "expected"),
    [
        ("01.03.1988 р.н., освіта: вища.", []),
        ("Народився 1 березня 1988 року.", []),
        (
            "02.03.1988 р.н., освіта: вища.",
            [(WARNING, "Пункт 1", "дата народження 02.03.1988 не збігається з РНОКПП")],
        ),
        (
            "Народився 2 березня 1988 року.",
            [(WARNING, "Пункт 1", "дата народження 02.03.1988 не збігається з РНОКПП")],
        ),
        (
            "1989 р.н., освіта: вища.",
            [(WARNING, "Пункт 1", "рік народження 1989 не збігається з РНОКПП")],
        ),
        (
            "Народилася 1 березня 1988 року.",
            [(WARNING, "Пункт 1", "стать у РНОКПП не збігається з «Народилася»")],
        ),
    ],
)
def test_birth_date_and_sex_against_ipn_are_warnings(bio, expected):
    text = _order(
        "1. Капітана ТЕСТЕНКА Олега Васильовича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        bio,
        f"{IPN_A}.",
    )
    assert _by_rule(_review(text), "РНОКПП") == expected


# ── Номенклатура: до підполковника включно, лише військові звання ───────────
def _dismissal(*items: str) -> str:
    return "\n".join(
        [
            "НАКАЗ",
            "§ 1",
            "Відповідно до підпункту «б» пункту 2 частини п'ятої статті 26 Закону України "
            "нижчепойменованих офіцерів ЗВІЛЬНИТИ З ВІЙСЬКОВОЇ СЛУЖБИ:",
            *items,
        ]
    )


@pytest.mark.parametrize("rank", ["Підполковника", "Майора", "Старшого лейтенанта"])
def test_up_to_lieutenant_colonel_is_within_nomenclature(rank):
    text = _dismissal(f"1. {rank} ТЕСТЕНКА Олега Васильовича, начальника штабу батальйону.")
    assert _by_rule(_review(text), "Номенклатура") == []


@pytest.mark.parametrize("rank", ["Полковника", "Генерал-майора", "Бригадного генерала"])
def test_dismissal_above_lieutenant_colonel_is_an_error(rank):
    text = _dismissal(f"1. {rank} ТЕСТЕНКА Олега Васильовича, командира бригади.")
    findings = [f for f in _review(text).findings if f.rule == "Номенклатура"]
    assert [(f.level, f.where) for f in findings] == [(ERROR, "Пункт 1")]
    assert f"«{rank}» вище за підполковник — поза номенклатурою звільнення" in findings[0].what
    assert "1153/2008" in findings[0].how


@pytest.mark.parametrize("rank", ["Полковника", "Бригадного генерала"])
def test_appointment_and_transfer_have_no_rank_limit(rank):
    """Призначати й переміщати можна будь-кого — межа лише для звільнення (25.09.2026)."""
    appointment = _order(
        f"1. {rank} ТЕСТЕНКА Олега Васильовича, командира бригади - КОМАНДИРОМ КОРПУСУ.",
        "Призначається на вищу посаду з шпк «полковник» на шпк «генерал-майор».",
    )
    transfer = "\n".join([
        "НАКАЗ", "§ 1", "Нижчепойменованих офіцерів ПЕРЕВЕСТИ:",
        f"1. {rank} ТЕСТЕНКА Олега Васильовича, командира бригади.",
    ])
    assert _by_rule(_review(appointment), "Номенклатура") == []
    assert _by_rule(_review(transfer), "Номенклатура") == []


def test_naval_rank_is_a_warning():
    text = _dismissal("1. Капітана 2 рангу ТЕСТЕНКА Олега Васильовича, командира корабля.")
    assert _by_rule(_review(text), "Номенклатура") == [
        (
            WARNING,
            "Пункт 1",
            "звання особи «Капітана 2 рангу» — корабельне звання, а не Сухопутних військ",
        ),
    ]


def test_nomenclature_limits_are_read_from_the_dictionary():
    limits = review_check.load_nomenclature()
    assert {action: limit[1] for action, limit in limits.items()} == {
        "звільнення": "підполковник",
        "призначення": "",
        "усі": "",
    }
    assert limits["призначення"][0] is None and limits["усі"][0] is None
    assert review_check.rank_level("підполковник") == limits["звільнення"][0] == 16


# ── Алфавітний порядок прізвищ у групі ──────────────────────────────────────
def _alphabet(*surnames: str, head=None) -> list:
    items = [
        f"{number}. Капітана {surname} Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ."
        for number, surname in enumerate(surnames, start=1)
    ]
    text = (head or _order)(*items)
    return _by_rule(_review(text), "Алфавітний порядок")


def test_alphabetical_order_is_fine():
    assert _alphabet("АНДРІЄНКА", "БОНДАРЯ", "ГОНЧАРА", "ҐУДЗЯ", "ЄВЕНКА", "ІВАНЕНКА", "ЇЖАКА") == []


def test_wrong_order_is_a_warning():
    assert _alphabet("БОНДАРЯ", "АНДРІЄНКА", "ГОНЧАРА") == [
        (WARNING, "Пункт 2", "«АНДРІЄНКА» стоїть після «БОНДАРЯ» — порушено абетку"),
    ]


def test_ukrainian_letters_follow_the_ukrainian_alphabet():
    """Звичайне сортування ставить ґ, є, і, ї після «я»; українська абетка — ні."""
    assert _alphabet("ЯРЕМЕНКА", "ІВАНЕНКА") != []
    assert _alphabet("ГОНЧАРА", "ҐУДЗЯ", "ДАНИЛЕНКА") == []
    assert _alphabet("ИВАХНА", "ІВАНЕНКА", "ЇЖАКА", "ЙОСИПЕНКА") == []


def test_case_ending_does_not_break_the_order():
    """Коваль → «КОВАЛЯ», Коваленко → «КОВАЛЕНКА»: Коваль іде першим."""
    assert _alphabet("КОВАЛЯ", "КОВАЛЕНКА") == []
    assert _alphabet("КОВАЛЕНКА", "КОВАЛЯ") != []


def test_same_surname_is_ordered_by_name():
    text = _order(
        "1. Капітана ТЕСТЕНКА Петра Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
        "2. Капітана ТЕСТЕНКА Андрія Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
    )
    assert [where for _l, where, _w in _by_rule(_review(text), "Алфавітний порядок")] == ["Пункт 2"]


def test_new_heading_starts_a_new_group():
    text = "\n".join([
        "НАКАЗ", "§ 1", "Відповідно до пункту 2 ЗВІЛЬНИТИ З ВІЙСЬКОВОЇ СЛУЖБИ:",
        "У ЗАПАС ЗА ПІДПУНКТОМ «а»:",
        "1. Капітана ЯРЕМЕНКА Івана Івановича, командира роти.",
        "1985 р.н.",
        "У ЗАПАС ЗА ПІДПУНКТОМ «б»:",
        "2. Капітана АНДРІЄНКА Івана Івановича, командира роти.",
        "§ 2", "Відповідно до пункту 1 ПРИЗНАЧИТИ:",
        "3. Капітана АБРАМЕНКА Івана Івановича, командира роти - КОМАНДИРОМ БАТАЛЬЙОНУ.",
    ])
    assert _by_rule(_review(text), "Алфавітний порядок") == []
