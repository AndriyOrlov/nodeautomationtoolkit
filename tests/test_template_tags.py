"""Спільні імена тегів: «засвідчувач» і «затверджувач» — те саме поле.

Заготовки писали різні люди в різний час, тому в них трапляються і
«{{засвідчувач_піб}}», і «{{затверджувач_піб}}», і «{{згідно_з_оригіналом_піб}}».
Підстановка не має залежати від того, яку назву обрав автор зразка.
"""

import sys
from pathlib import Path

import pytest

SRC_PATH = Path(__file__).resolve().parents[1] / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.template_tags import (
    SIGNER_TAGS,
    certifier_tags,
    expand_common_tags,
    signer_tags,
    tag_aliases,
)


@pytest.mark.parametrize("tag", [
    "{{засвідчувач_піб}}", "{{затверджувач_піб}}", "{{згідно_з_оригіналом_піб}}",
    "{{ЗАСВІДЧУВАЧ_ПІБ}}",
])
def test_certifier_name_aliases_are_one_group(tag):
    aliases = tag_aliases(tag)
    assert "{{затверджувач_піб}}" in aliases
    assert "{{згідно_з_оригіналом_піб}}" in aliases


@pytest.mark.parametrize("tag", ["{{підписант_посада}}", "{{посада_підписанта}}"])
def test_signer_position_aliases_are_one_group(tag):
    assert set(tag_aliases(tag)) == {"{{підписант_посада}}", "{{посада_підписанта}}"}


def test_unknown_tag_is_left_alone():
    assert tag_aliases("{{невідомий_тег}}") == ("{{невідомий_тег}}",)


def test_expand_adds_synonyms_without_touching_explicit_values():
    values = {"{{затверджувач_піб}}": "Вигаданко П.І.", "{{номер_наказу}}": "№445"}
    expanded = expand_common_tags(values)
    assert expanded["{{засвідчувач_піб}}"] == "Вигаданко П.І."
    assert expanded["{{згідно_з_оригіналом_піб}}"] == "Вигаданко П.І."
    # Явно переданий формат номера не змінюємо.
    assert expanded["{{номер_наказу}}"] == "№445"


def test_explicitly_set_alias_wins_over_generated_one():
    values = {
        "{{засвідчувач_піб}}": "Перший",
        "{{затверджувач_піб}}": "Другий",
    }
    expanded = expand_common_tags(values)
    assert expanded["{{засвідчувач_піб}}"] == "Перший"
    assert expanded["{{затверджувач_піб}}"] == "Другий"


def test_certifier_tags_cover_every_spelling():
    tags = certifier_tags("Начальник штабу", "полковник", "Вигаданко П.І.")
    for expected in (
        "{{засвідчувач_посада}}", "{{затверджувач_посада}}", "{{згідно_з_оригіналом_посада}}",
        "{{засвідчувач_звання}}", "{{затверджувач_звання}}", "{{згідно_з_оригіналом_звання}}",
        "{{засвідчувач_піб}}", "{{затверджувач_піб}}", "{{згідно_з_оригіналом_піб}}",
        "{{засвідчувач}}", "{{затверджувач}}",
    ):
        assert expected in tags
    assert tags["{{затверджувач_звання}}"] == "полковник"


def test_signer_tags_include_the_plain_and_detailed_forms():
    assert "{{підписант}}" in SIGNER_TAGS
    assert "{{підписант_піб}}" in SIGNER_TAGS
    assert "{{прізвище_підписанта}}" in SIGNER_TAGS
    # Засвідчувач — не підписант: сплутати ці блоки не можна.
    assert not any("засвідчувач" in tag for tag in SIGNER_TAGS)


@pytest.mark.parametrize(("first", "second"), [
    ("{{номер_наказу}}", "{{номер}}"),
    ("{{дата_наказу}}", "{{дата}}"),
    ("{{примірник}}", "{{примірник_номер}}"),
    ("{{згідно_з_оригіналом}}", "{{засвідчення}}"),
])
def test_common_document_tags_are_aliases(first, second):
    assert second in tag_aliases(first)
    assert first in tag_aliases(second)


def test_plain_signer_tag_is_the_same_full_block_everywhere():
    tags = signer_tags("Командувач військ", "полковник", "Вигаданко П.І.")
    assert tags["{{підписант}}"] == "Командувач військ\rполковник Вигаданко П.І."
    assert tags["{{підписант_імя}}"] == "Вигаданко П.І."
    assert tags["{{посада_підписанта}}"] == "Командувач військ"


def test_approver_name_spelling_is_enabled():
    tags = certifier_tags("Начальник", "майор", "Прикладенко П.П.")
    assert tags["{{затверджувач_імя}}"] == "Прикладенко П.П."
