"""Позначка «$» у рядку таблиці: частина в повідомленні НЕ шифрується (AGENT.md 9.5.11).

Рішення користувача 17.09.2026: «$ — ігноруємо той рядок і не закриваємо», діє
«тільки в шифрах» — витяги рядок бачать як звичайний. Усі дані вигадані.
"""

import openpyxl

import generate_extracts as generator
from nodeautomationtoolkit.builtin_nodes.message_order import (
    cipher_unit_names,
    generate_decision_order,
)
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import read_recipient_mapping

HEADER = ["Відкрите найменування", "Шифр", "Скорочення", "Корпус", "Кому", "Куди"]
TEXT = "командира взводу 77 окремої тестової бригади та 88 окремого тестового полку"


def _table(tmp_path, rows):
    path = tmp_path / "словник.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(HEADER)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return read_recipient_mapping(path=str(path))


def _rows(mark_in_a=True):
    brigade = "$77 окрема тестова бригада" if mark_in_a else "77 окрема тестова бригада"
    return [
        [brigade, "А7777", "77 отбр", "", "Командиру військової частини А7777", "м. Тест"],
        ["88 окремий тестовий полк", "А8888", "88 отп", "", "Командиру військової частини А8888", "м. Тест"],
    ]


def test_marked_row_stays_in_mapping_for_extracts(tmp_path):
    result = _table(tmp_path, _rows())
    entry = result["mapping"]["77 окрема тестова бригада"]

    assert entry["cipher"] == "А7777"
    assert entry["recipient_to"] == "Командиру військової частини А7777"
    assert result["keep_open"] == ["77 окрема тестова бригада"]


def test_mark_in_any_cell_counts(tmp_path):
    rows = _rows(mark_in_a=False)
    rows[0][1] = "А7777$"
    result = _table(tmp_path, rows)

    assert result["mapping"]["77 окрема тестова бригада"]["cipher"] == "А7777"
    assert result["keep_open"] == ["77 окрема тестова бригада"]


def test_marked_unit_is_not_ciphered_other_units_are(tmp_path):
    mapping = _table(tmp_path, _rows())["mapping"]

    result, _, _ = cipher_unit_names(TEXT, mapping)

    assert result == "командира взводу 77 окремої тестової бригади та військової частини А8888"


def test_without_mark_everything_is_ciphered_as_before(tmp_path):
    mapping = _table(tmp_path, _rows(mark_in_a=False))["mapping"]

    result, _, _ = cipher_unit_names(TEXT, mapping)

    assert result == "командира взводу військової частини А7777 та військової частини А8888"


def test_marked_unit_is_neither_new_nor_highlighted(tmp_path):
    mapping = _table(tmp_path, _rows())["mapping"]
    ciphered, _, _ = cipher_unit_names(TEXT, mapping)

    assert generator.collect_new_unit_names(TEXT, mapping) == []
    assert generator.find_unmatched_open_unit_spans(ciphered)  # без таблиці — підсвітилось би
    assert generator.unmatched_open_unit_spans(ciphered, mapping) == []


def test_marked_unit_stays_open_in_decision_order(tmp_path):
    mapping = _table(tmp_path, _rows())["mapping"]

    decision = generate_decision_order(text=f"НАКАЗУЮ:\n1. Капітана ТЕСТЕНКА, {TEXT}.", mapping=mapping)

    assert "77 окремої тестової бригади" in decision["decision_text"]
    assert "А7777" not in decision["decision_text"]
    assert "А8888" in decision["decision_text"]


def test_routing_still_finds_marked_unit(tmp_path):
    marked = _table(tmp_path, _rows())["mapping"]
    plain = _table(tmp_path, _rows(mark_in_a=False))["mapping"]
    order = f"НАКАЗУЮ:\n1. Капітана ТЕСТЕНКА Тест Тестовича, {TEXT}, ПРИЗНАЧИТИ …"

    def _senders(mapping):
        routes = generator.map_military_units(text=order, mapping=mapping)
        return sorted(row[0] for row in routes["units_table"].rows)

    assert _senders(marked) == _senders(plain) == ["77 отбр А7777", "88 отп А8888"]
