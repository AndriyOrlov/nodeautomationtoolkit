"""Шифр у закритому змісті — СТРОГО з таблиці, нові частини видно (AGENT.md 9.5.7).

Закріплює дефекти, знайдені перевіркою генерації повідомлень (усі дані вигадані):

1. Порожній стовпець B: `read_recipient_mapping` кладе в `cipher` відкриту
   назву, і в закритий текст ішло «військової частини 77 окрема тестова бригада».
2. Корпус зі стовпця D, якого немає в таблиці: у текст ішов сам стовпець D —
   «військової частини А0077 військової частини 99 АК».
3. Описки виправлялися в САМОМУ тексті: абзац без жодної частини змінювався
   (`зв’язку` → `зв'язку`, `в. ч.` → `в/ч`).
4. Нові частини («169 батальйону резерву», «12 навчального центру») лишались
   відкритими без жовтої позначки.
"""

import openpyxl
import pytest

import generate_extracts as generator
from nodeautomationtoolkit.builtin_nodes.message_order import cipher_unit_names
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    _format_full_closed_unit_text,
    map_military_units,
    read_recipient_mapping,
)

BRIGADE = "77 окрема тестова бригада"
HEADER = ["Відкрите найменування", "Шифр", "Скорочення", "Корпус", "Кому", "Куди"]


def _entry(open_name, cipher, abbreviation="", corps="", recipient_to=""):
    return {
        "open_name": open_name,
        "cipher": cipher,
        "abbreviation": abbreviation,
        "corps": corps,
        "recipient_to": recipient_to,
    }


def _xlsx(tmp_path, rows, name="словник.xlsx"):
    path = tmp_path / name
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(HEADER)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return path


class _Var:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


# ── 1–2. Шифр лише з таблиці ────────────────────────────────────────────────


def test_empty_column_b_leaves_name_open_and_reports(tmp_path):
    table = tmp_path / "словник.csv"
    table.write_text(";".join(HEADER) + f"\n{BRIGADE};;77 отбр;;;\n", encoding="utf-8")
    mapping = read_recipient_mapping(str(table))["mapping"]
    text = "командира взводу 77 окремої тестової бригади"

    problems = []
    result, count, rows = cipher_unit_names(text, mapping, problems=problems)

    assert result == text
    assert count == 0
    assert rows == []
    assert problems == [("no_cipher", BRIGADE)]


def test_unknown_corps_in_column_d_adds_no_invented_link():
    mapping = {BRIGADE: _entry(BRIGADE, "А0077", "77 отбр", "99 АК")}

    problems = []
    result, _, _ = cipher_unit_names("командира взводу 77 окремої тестової бригади", mapping, problems=problems)

    assert result == "командира взводу військової частини А0077"
    assert problems == [("corps_missing", BRIGADE, "99 АК")]


def test_corps_row_without_cipher_adds_no_link():
    corps = "51 армійський корпус"
    mapping = {
        BRIGADE: _entry(BRIGADE, "А0077", "77 отбр", "51 АК"),
        corps: _entry(corps, corps, "51 АК"),  # B порожній — у cipher відкрита назва
    }

    result, _, _ = cipher_unit_names("командира взводу 77 окремої тестової бригади", mapping)

    assert result == "командира взводу військової частини А0077"
    assert _format_full_closed_unit_text(mapping[BRIGADE], mapping) == "військової частини А0077"


def test_problems_are_reported_once_per_row():
    mapping = {BRIGADE: _entry(BRIGADE, "А0077", "77 отбр", "99 АК")}
    problems = []
    cipher_unit_names(
        "з 77 окремої тестової бригади до 77 окремої тестової бригади", mapping, problems=problems
    )
    assert problems == [("corps_missing", BRIGADE, "99 АК")]


# ── 3. Описки — лише для пошуку ─────────────────────────────────────────────


def test_paragraph_without_units_is_not_rewritten():
    source = "начальника вузла зв’язку в. ч. тестового гарнізону"
    result, count, _ = cipher_unit_names(source, {BRIGADE: _entry(BRIGADE, "А0077")})
    assert result == source
    assert count == 0


def test_name_with_typo_is_found_but_rest_of_text_stays_original():
    mapping = {BRIGADE: _entry(BRIGADE, "А0077")}
    result, _, _ = cipher_unit_names("начальник зв’язку 77 окремої тестової бригди", mapping)
    assert result == "начальник зв’язку військової частини А0077"


# ── 4. Жовта позначка для нових частин ──────────────────────────────────────


@pytest.mark.parametrize(
    "text, expected",
    [
        ("офіцера 169 батальйону резерву", "169 батальйону"),
        ("офіцера 12 навчального центру", "12 навчального центру"),
        ("офіцера 300 військового госпіталю", "300 військового госпіталю"),
        ("офіцера 40 бази зберігання", "40 бази"),
        ("офіцера 88 окремого вузла зв'язку", "88 окремого вузла"),
    ],
)
def test_numbered_units_of_any_kind_are_highlighted(text, expected):
    spans = generator.find_unmatched_open_unit_spans(text)
    assert [text[start:end] for start, end in spans] == [expected]


@pytest.mark.parametrize(
    "text",
    [
        "офіцера 2 відділу Калуського районного територіального центру комплектування",
        "командира 1 механізованого батальйону цієї самої військової частини",
        "командира 2 батальйону військової частини А1111",
        "з 01.06.2026 року",
        "відповідно до пункту 2 частини четвертої статті 26 Закону України",
    ],
)
def test_subunits_dates_and_references_are_not_highlighted(text):
    assert generator.find_unmatched_open_unit_spans(text) == []


def test_new_units_are_collected_once_with_full_name():
    mapping = {
        BRIGADE: _entry(BRIGADE, "А0077"),
        "55 окремий батальйон тестування": _entry(
            "55 окремий батальйон тестування", "55 окремий батальйон тестування"
        ),
    }
    text = (
        "1. Лейтенанта ТЕСТЕНКА Теста Тестовича, командира взводу 77 окремої тестової бригади, "
        "офіцера 169 батальйону резерву.\n"
        "2. Капітана ВИГАДАНОГО Вигада Вигадовича, офіцера 55 окремого батальйону тестування.\n"
        "3. Майора ПРИКЛАДОВА Приклада Прикладовича, офіцера 169 батальйону резерву."
    )

    # Частина з порожнім B у таблиці вже є — нова лише одна, і без повторів.
    assert generator.collect_new_unit_names(text, mapping) == ["169 батальйону резерву"]


def test_recipient_without_cipher_does_not_leak_open_name():
    mapping = {BRIGADE: _entry(BRIGADE, BRIGADE, "77 отбр", "", "Командиру 77 отбр")}
    routes = map_military_units(text="§ 1\n1. Направити до 77 окремої тестової бригади.", mapping=mapping)

    recipients = generator.build_message_recipient_list(mapping, routes)

    assert "окрема тестова бригада" not in " ".join(recipients)


# ── Заготовки нових частин у таблиці ────────────────────────────────────────


def test_stubs_are_appended_to_table_without_formulas(tmp_path):
    path = _xlsx(tmp_path, [[BRIGADE, "А0077", "77 отбр", "", "", ""]])

    result = generator.append_unit_stubs_to_table(
        str(path), ["169 батальйону резерву", "169  батальйону резерву ", BRIGADE]
    )

    assert result["added"] == ["169 батальйону резерву"]
    assert result["separate"] is False
    assert openpyxl.load_workbook(result["backup"]).active.max_row == 2
    sheet = openpyxl.load_workbook(path).active
    assert sheet.cell(row=3, column=1).value == "169 батальйону резерву"
    assert sheet.cell(row=3, column=2).value is None
    assert sheet.cell(row=3, column=1).fill.start_color.rgb.endswith("FFFF00")

    # Заготовка без шифру маршрутизації не зачіпає, а шифр частини цілий.
    mapping = read_recipient_mapping(str(path))["mapping"]
    assert "169 батальйону резерву" not in mapping
    assert mapping[BRIGADE]["cipher"] == "А0077"

    # Повторний запуск не дублює заготовку.
    assert generator.append_unit_stubs_to_table(str(path), ["169 батальйону резерву"])["added"] == []


def test_table_with_formulas_is_not_touched(tmp_path):
    path = _xlsx(tmp_path, [[BRIGADE, '="А"&"0077"', "77 отбр", "", "", ""]])

    result = generator.append_unit_stubs_to_table(str(path), ["169 батальйону резерву"])

    assert result["separate"] is True
    assert result["path"].endswith("— нові частини.xlsx")
    assert openpyxl.load_workbook(path).active.max_row == 2
    stub_sheet = openpyxl.load_workbook(result["path"]).active
    assert stub_sheet.cell(row=2, column=1).value == "169 батальйону резерву"


def test_stubs_are_appended_to_csv_table(tmp_path):
    path = tmp_path / "словник.csv"
    path.write_text(";".join(HEADER) + f"\n{BRIGADE};А0077;77 отбр;;;", encoding="utf-8")

    result = generator.append_unit_stubs_to_table(str(path), ["169 батальйону резерву"])

    assert result["added"] == ["169 батальйону резерву"]
    assert path.read_text(encoding="utf-8").splitlines()[-1] == "169 батальйону резерву;;;;;"
    assert "169 батальйону резерву" not in read_recipient_mapping(str(path))["mapping"]


def test_report_table_gaps_adds_stubs_only_when_enabled(tmp_path):
    path = _xlsx(tmp_path, [[BRIGADE, "А0077", "77 отбр", "", "", ""]])
    app = generator.App.__new__(generator.App)
    logs = []
    app.log = logs.append
    app.excel_path = _Var(str(path))
    mapping = read_recipient_mapping(str(path))["mapping"]
    order = "НАКАЗ\n§ 1\n1. Лейтенанта ТЕСТЕНКА Теста Тестовича, офіцера 169 батальйону резерву.\n"

    assert app._report_table_gaps(order, mapping) == {"problems": 0, "new_units": 1}
    assert any("169 батальйону резерву" in line for line in logs)
    assert openpyxl.load_workbook(path).active.max_row == 2

    app.ADD_NEW_UNITS_TO_TABLE = True
    app._report_table_gaps(order, mapping)
    assert openpyxl.load_workbook(path).active.cell(row=3, column=1).value == "169 батальйону резерву"


def test_only_qt_shell_adds_units_to_table():
    pytest.importorskip("PySide6")
    from nodeautomationtoolkit.generator_qt.main_window import create_qt_app_class

    assert generator.App.ADD_NEW_UNITS_TO_TABLE is False
    assert create_qt_app_class(generator).ADD_NEW_UNITS_TO_TABLE is True
