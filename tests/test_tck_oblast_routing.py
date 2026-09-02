"""Районний/міський ТЦК ЗАВЖДИ прямує до свого ОБЛАСНОГО ТЦК.

Власного витягу районний ТЦК не отримує ніколи. У наказі область звичайно
названо прямо («… ТЦК та СП Волинської області»), і саме вона визначає
адресата; «Кому»/«Куди» беруться з рядка області.
"""

import sys
from pathlib import Path

import pytest

SRC_PATH = Path(__file__).resolve().parents[1] / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units


def _entry(name, abbr, city=""):
    return {
        "open_name": name,
        "cipher": "",
        "abbreviation": abbr,
        "corps": "",
        "recipient_to": f"Начальнику {abbr}" if city else "",
        "destination_where": city,
    }


OBLAST = "Волинський обласний територіальний центр комплектування та соціальної підтримки"
RAYON = "Ковельський районний територіальний центр комплектування та соціальної підтримки"

ONLY_OBLAST = {OBLAST: _entry(OBLAST, "Волинський ОТЦК та СП", "м. Луцьк")}
BOTH = dict(ONLY_OBLAST, **{RAYON: _entry(RAYON, "Ковельський РТЦК та СП")})


def _order(unit_phrase):
    return (
        "§ 1\n"
        "Відповідно до пункту 1 ПРИЗНАЧИТИ:\n"
        f"1. Капітана ТЕСТЕНКА, офіцера {unit_phrase}, офіцером цього самого центру.\n"
        "\nКомандир військової частини А0001\nполковник   Петро ТЕСТОВИЙ\n"
    )


RAYON_IN_ORDER = (
    "Ковельського районного територіального центру комплектування "
    "та соціальної підтримки Волинської області"
)


@pytest.mark.parametrize("mapping", [ONLY_OBLAST, BOTH], ids=["лише-обласний", "обласний+районний"])
def test_rayon_tck_always_routes_to_its_oblast(mapping):
    """Навіть якщо районний ТЦК є ОКРЕМИМ РЯДКОМ таблиці, витяг один — на область.

    Без переспрямування районний рядок збігався як звичайна частина і
    отримував власний витяг.
    """
    routes = map_military_units(text=_order(RAYON_IN_ORDER), mapping=mapping)
    units = routes["unit_paragraphs"]

    assert list(units) == ["Волинський ОТЦК та СП"]
    entry = units["Волинський ОТЦК та СП"]
    assert entry["recipient_to"] == "Начальнику Волинський ОТЦК та СП"
    assert entry["destination_where"] == "м. Луцьк"


def test_oblast_named_in_the_order_wins_over_an_unknown_rayon():
    """Область названо прямо, тож навіть незнайомий район маршрутизується."""
    unknown = (
        "Вигаданського районного територіального центру комплектування "
        "та соціальної підтримки Волинської області"
    )
    routes = map_military_units(text=_order(unknown), mapping=ONLY_OBLAST)

    assert list(routes["unit_paragraphs"]) == ["Волинський ОТЦК та СП"]


def test_city_tck_without_an_oblast_is_reported_not_guessed():
    """Виняток — міський обʼєднаний ТЦК без області: адресата не вигадуємо.

    Такий пункт має бути видно в «Контроль пропущених пунктів», а не мовчки
    піти до випадкового ТЦК.
    """
    city = (
        "Криворізького міського обʼєднаного територіального центру "
        "комплектування та соціальної підтримки"
    )
    routes = map_military_units(text=_order(city), mapping=ONLY_OBLAST)

    assert routes["unit_paragraphs"] == {}
    assert len(routes["unmatched_items"]) == 1


def test_explicit_city_wins_over_compound_district_name():
    lviv_oblast = (
        "Львівський обласний територіальний центр комплектування "
        "та соціальної підтримки"
    )
    mapping = {
        lviv_oblast: _entry(lviv_oblast, "Львівський ОТЦК та СП", "м. Львів")
    }
    district_in_lviv = (
        "Галицько-Франківського об’єднаного районного у місті Львові "
        "територіального центру комплектування та соціальної підтримки"
    )

    routes = map_military_units(text=_order(district_in_lviv), mapping=mapping)

    assert list(routes["unit_paragraphs"]) == ["Львівський ОТЦК та СП"]
    entry = routes["unit_paragraphs"]["Львівський ОТЦК та СП"]
    assert entry["destination_where"] == "м. Львів"


def test_one_item_with_two_explicit_oblasts_routes_to_both_oblast_tcks():
    ternopil_oblast = (
        "Тернопільський обласний територіальний центр комплектування "
        "та соціальної підтримки"
    )
    lviv_oblast = (
        "Львівський обласний територіальний центр комплектування "
        "та соціальної підтримки"
    )
    mapping = {
        ternopil_oblast: _entry(ternopil_oblast, "Тернопільський ОТЦК та СП", "м. Тернопіль"),
        lviv_oblast: _entry(lviv_oblast, "Львівський ОТЦК та СП", "м. Львів"),
    }
    item = (
        "Чортківського районного територіального центру комплектування та "
        "соціальної підтримки Тернопільської області, заступника начальника "
        "Самбірського районного територіального центру комплектування та "
        "соціальної підтримки Львівської області СКАСУВАТИ як нереалізований"
    )

    routes = map_military_units(text=_order(item), mapping=mapping)

    assert set(routes["unit_paragraphs"]) == {
        "Тернопільський ОТЦК та СП",
        "Львівський ОТЦК та СП",
    }


# ── Районний ТЦК У МІСТІ + внутрішнє переміщення ВЕЛИКИМИ літерами ───────────
CITY_DISTRICT = (
    "Сихівський районний у місті Львові територіальний центр комплектування "
    "та соціальної підтримки"
)
LVIV_OBLAST = (
    "Львівський обласний територіальний центр комплектування та соціальної підтримки"
)

CITY_DISTRICT_ITEM = (
    "4. Підполковника ТЕСТЕНКА Тараса Тарасовича, начальника Сихівського районного "
    "у місті Львові територіального центру комплектування та соціальної підтримки – "
    "ЗАСТУПНИКОМ НАЧАЛЬНИКА ТЕРИТОРІАЛЬНОГО ЦЕНТРУ КОМПЛЕКТУВАННЯ ТА СОЦІАЛЬНОЇ "
    "ПІДТРИМКИ – НАЧАЛЬНИКОМ МОБІЛІЗАЦІЙНОГО ВІДДІЛЕННЯ ЦЬОГО САМОГО ЦЕНТРУ."
)

_LVIV_ONLY = {LVIV_OBLAST: _entry(LVIV_OBLAST, "Львівський ОТЦК та СП", "м. Львів")}
_LVIV_AND_DISTRICT = dict(
    _LVIV_ONLY, **{CITY_DISTRICT: _entry(CITY_DISTRICT, "Сихівський РТЦК та СП")}
)


@pytest.mark.parametrize(
    "mapping",
    [_LVIV_ONLY, _LVIV_AND_DISTRICT],
    ids=["лише-обласний", "обласний+районний"],
)
def test_city_district_tck_routes_to_the_oblast_once(mapping):
    """Районний ТЦК У МІСТІ дає ОДИН витяг — на обласний.

    Область береться з назви міста («у місті Львові»), а «ЦЬОГО САМОГО ЦЕНТРУ»
    великими літерами має лишатися внутрішнім переміщенням, тобто не додавати
    другого адресата.
    """
    order = (
        "§ 1\nВідповідно до пункту 1 ПРИЗНАЧИТИ:\n\n"
        + CITY_DISTRICT_ITEM
        + "\n\nКомандир військової частини А0001\nполковник   Петро ТЕСТОВИЙ\n"
    )
    routes = map_military_units(text=order, mapping=mapping)

    assert list(routes["unit_paragraphs"]) == ["Львівський ОТЦК та СП"]
    assert routes["unmatched_items"] == []


# ── У наказі пропущено слово «територіального» ───────────────────────────────
KHMEL_OBLAST = (
    "Хмельницький обласний територіальний центр комплектування та соціальної підтримки"
)
_KHMEL_MAPPING = {
    KHMEL_OBLAST: _entry(KHMEL_OBLAST, "Хмельницький ОТЦК та СП", "м. Хмельницький")
}

CONTRACT_ITEM = (
    "1. Відповідно до статей 20 та 23 Закону України “Про військовий обовʼязок і "
    "військову службу” з лейтенантом запасу ТЕСТЕНКОМ Тарасом Тарасовичем УКЛАСТИ "
    "контракт про проходження військової служби строком на 3 (три) роки, ПРИЙНЯТИ її "
    "на військову службу за контрактом, ЗАРАХУВАТИ до списків особового складу "
    "Збройних Сил України та ПРИЗНАЧИТИ ОФІЦЕРОМ ВІДДІЛЕННЯ ОБЛІКУ ОФІЦЕРІВ ЗАПАСУ "
    "ВІДДІЛУ ВІЙСЬКОВОГО ОБЛІКУ ТА БРОНЮВАННЯ СЕКТОРУ МОБІЛІЗАЦІЙНО-ОБОРОННОЇ РОБОТИ "
    "{name}."
)


@pytest.mark.parametrize(
    "name_in_order",
    [
        "ХМЕЛЬНИЦЬКОГО ОБЛАСНОГО ЦЕНТРУ КОМПЛЕКТУВАННЯ ТА СОЦІАЛЬНОЇ ПІДТРИМКИ",
        "ХМЕЛЬНИЦЬКОГО ОБЛАСНОГО ТЕРИТОРІАЛЬНОГО ЦЕНТРУ КОМПЛЕКТУВАННЯ ТА "
        "СОЦІАЛЬНОЇ ПІДТРИМКИ",
    ],
    ids=["без-слова-територіального", "повна-назва"],
)
def test_item_is_not_skipped_when_the_order_omits_territorial(name_in_order):
    """У наказі часто пропускають слово «територіального».

    Шаблон зі стовпця A вимагає ВСІ значущі слова, тож пункт не знаходив
    адресата зовсім і потрапляв у «Пропущені». Ядром назви лишається
    «центр комплектування», а «територіальний» став необовʼязковим.
    """
    order = (
        "§ 1\n"
        + CONTRACT_ITEM.format(name=name_in_order)
        + "\n\nКомандир військової частини А0001\nполковник   Петро ТЕСТОВИЙ\n"
    )
    routes = map_military_units(text=order, mapping=_KHMEL_MAPPING)

    assert routes["unmatched_items"] == []
    assert list(routes["unit_paragraphs"]) == ["Хмельницький ОТЦК та СП"]


def test_a_plain_training_centre_is_still_not_a_tck():
    """«Навчальний центр» не має ставати ТЦК через послаблену вимогу."""
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _extract_tck_sender

    assert _extract_tck_sender("навчального центру Сухопутних військ") is None
