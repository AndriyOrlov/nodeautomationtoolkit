"""Тести модуля генерації примірників (`copy_generator`).

Перевіряються чисті функції, які не потребують Word: розпізнавання прізвища
в рядку підпису та визначення початку тіла наказу.
"""

import pytest

from nodeautomationtoolkit.builtin_nodes.copy_generator import (
    find_signature_name_tail,
    find_body_start_paragraph_index,
)


class _FakeRange:
    def __init__(self, text: str):
        self.Text = text


class _FakeParagraph:
    def __init__(self, text: str):
        self.Range = _FakeRange(text + "\r")


class _FakeParagraphs:
    def __init__(self, texts):
        self._items = [_FakeParagraph(t) for t in texts]
        self.Count = len(self._items)

    def __call__(self, index):
        return self._items[index - 1]


class _FakeDoc:
    def __init__(self, texts):
        self.Paragraphs = _FakeParagraphs(texts)


@pytest.mark.parametrize(
    "line, expected",
    [
        ("полковник                    Іван ПЕТРЕНКО", "Іван ПЕТРЕНКО"),
        ("Міністр оборони України            Андрій КОВАЛЕНКО", "Андрій КОВАЛЕНКО"),
        ("Командувач Сухопутних військ       С.М.ШЕВЧЕНКО", "С.М.ШЕВЧЕНКО"),
        ("генерал армії України              С.Т.БОНДАРЕНКО", "С.Т.БОНДАРЕНКО"),
    ],
)
def test_finds_surname_at_end_of_signature_line(line, expected):
    assert find_signature_name_tail(line) == expected


def test_no_surname_in_plain_sentence():
    assert find_signature_name_tail("Відповідно до пункту 82 Положення") == ""


def test_empty_line_is_safe():
    assert find_signature_name_tail("") == ""
    assert find_signature_name_tail(None) == ""


def test_body_starts_at_paragraph_sign():
    doc = _FakeDoc(["МІНІСТЕРСТВО ОБОРОНИ", "НАКАЗ", "", "§ 1", "1. Пункт."])
    assert find_body_start_paragraph_index(doc) == 4


def test_body_starts_at_numbered_item_without_paragraph_sign():
    doc = _FakeDoc(["Шапка наказу", "", "1. Перший пункт."])
    assert find_body_start_paragraph_index(doc) == 3


def test_body_starts_at_directive_keyword():
    doc = _FakeDoc(["Шапка", "нижчепойменованих осіб ЗВІЛЬНИТИ з посад:"])
    assert find_body_start_paragraph_index(doc) == 2


def test_body_start_falls_back_to_first_paragraph():
    doc = _FakeDoc(["Просто текст", "без ознак тіла"])
    assert find_body_start_paragraph_index(doc) == 1


def test_blank_paragraphs_are_skipped():
    doc = _FakeDoc(["", "   ", "§ 2"])
    assert find_body_start_paragraph_index(doc) == 3
