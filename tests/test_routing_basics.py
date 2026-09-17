"""Базові випадки маршрутизації пунктів (`map_military_units`).

Перенесено з тестів видаленого редактора нод (`tests/test_new_features.py`): там ці
перевірки йшли разом зі створенням витягів старим модулем `word_batch`, якого
більше немає. Маршрутизацію використовує генератор, тож перевірки лишаються.
"""

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units


def test_items_are_grouped_by_unit():
    order_text = """НАКАЗ
15 січня 2026 року № 77

1. 100 окрема бригада направляє 5 бійців.
2. 200 артилерійська бригада забезпечує техніку.
3. 100 окрема бригада надає додаткове постачання.
"""
    mapping = {
        "100 окрема бригада": "військова частина А1000",
        "200 артилерійська бригада": "військова частина А2000",
    }

    result = map_military_units(text=order_text, mapping=mapping)

    assert len(result["units_list"]) == 2
    assert "військова частина А1000" in result["unit_paragraphs"]
    assert len(result["unit_paragraphs"]["військова частина А1000"]["items"]) == 2


def test_section_heading_travels_with_the_item():
    order_text = """НАКАЗ
командира військової частини А0000
15 січня 2026 року № 88

§ 1

Відповідно до пунктів 82, 83 та 257 Положення про проходження громадянами України військової служби у Збройних Силах України нижчепойменованих осіб офіцерського складу ЗВІЛЬНИТИ з займаних посад і ПРИЗНАЧИТИ:

1.1. Майора Петренка П.П., 300 окрема механізована бригада, на посаду командира батальйону.
"""
    mapping = {"300 окрема механізована бригада": "військова частина А3000"}

    result = map_military_units(text=order_text, mapping=mapping)
    parent_heading = result["unit_paragraphs"]["військова частина А3000"]["items"][0]["parent_heading"]

    assert "§ 1" in parent_heading
    assert "Відповідно до пунктів 82, 83 та 257" in parent_heading
    assert "ПРИЗНАЧИТИ:" in parent_heading


def test_multiline_item_keeps_biography_and_basis_and_open_names():
    order_text = """НАКАЗ
15 червня 2026 року № 308

§ 2

10. Пункт 2 наказу про призначення на посаду начальника групи персоналу штабу 167 окремої механізованої бригади молодшого лейтенанта ІВАНОВА Івана Івановича, СКАСУВАТИ як нереалізований.

1997 р.н. 0000000000.
Підстава: клопотання начальника Волинського ОТЦК та СП від 30.07.2026 №ХХХ/ХХ.
"""
    mapping = {"167 окремої механізованої бригади": "військова частина А1670"}

    result = map_military_units(text=order_text, mapping=mapping)
    item = result["unit_paragraphs"]["військова частина А1670"]["items"][0]

    assert "10. Пункт 2 наказу" in item["text"]
    assert "1997 р.н. 0000000000." in item["text"]
    assert "Підстава: клопотання начальника" in item["text"]
    # Правило 4.1: item["text"] — відкриті назви, item["text_cipher"] — шифри.
    assert "167 окремої механізованої бригади" in item["text"]
    assert "військова частина А1670" in item["text_cipher"]
