"""Витяги до управління: окремий вихід, який НЕ бере участі в розсилці.

Зміни на посади в управлінні нікому не надсилаються, тому такий пункт не
має адресата й не потрапляє в розрахунок розсилки. Але витяг за ним усе
одно потрібен — окремим файлом, з тією самою структурою пункту.

Усі назви та реквізити тут вигадані (правило «публічний репозиторій — без
реальних даних»).
"""

import importlib.util
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units


def _load_generator():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_management_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


NODE = "555 інформаційно-телекомунікаційний вузол"
CENTER = "88 центр психологічної підтримки"


def _entry(name: str, cipher: str, abbreviation: str) -> dict:
    return {
        "open_name": name,
        "cipher": cipher,
        "abbreviation": abbreviation,
        "corps": "",
        "recipient_to": f"Командиру військової частини {cipher}",
        "destination_where": "м. Тестове",
    }


MAPPING = {
    NODE: _entry(NODE, "А0555", "555 ітв"),
    CENTER: _entry(CENTER, "А0088", "88 цпп"),
}

# Наказ про укладення контрактів: два пункти — по управлінню, два — по частинах.
ORDER = (
    "Відповідно до статті 19 Закону України, з нижчепойменованими особами "
    "офіцерського складу УКЛАСТИ КОНТРАКТИ:\n"
    "1. Підполковником ТЕСТЕНКОМ Олегом Васильовичем, старшим офіцером відділу "
    "планування розвідувального управління штабу управління.\n"
    "1985 р. н. 0000000000, строком на 24 місяці.\n"
    "2. Підполковником ПРИКЛАДЕНКОМ Олегом Васильовичем, начальником служби "
    "захисту інформації в автоматизованих системах управління.\n"
    "1993 р. н. 000000000, строком на 24 місяці.\n"
    "3. Підполковником ЗРАЗКОВИМ Олегом Васильовичем, заступником командира – "
    "начальником оперативно-технічного відділення 555 інформаційно-"
    "телекомунікаційного вузла.\n"
    "1971 р. н. 0000000, строком на 24 місяці.\n"
    "4. Підполковником ВИГАДАНКОМ Олегом Васильовичем, офіцером відділення "
    "супроводження груп психологічної та моральної підтримки 88 центру "
    "психологічної підтримки.\n"
    "1983 р. н. 0000000000, строком на 24 місяці.\n"
)


@pytest.fixture(scope="module")
def routes() -> dict:
    return map_military_units(text=ORDER, mapping=MAPPING)


def test_unit_items_still_reach_their_units(routes):
    """Пункти з назвою частини маршрутизуються як раніше."""
    assert {
        key: [item["label"] for item in data["items"]]
        for key, data in routes["unit_paragraphs"].items()
    } == {"555 ітв А0555": ["Пункт 3."], "88 цпп А0088": ["Пункт 4."]}
    assert routes["unmatched_items"] == []


def test_management_items_get_their_own_extract(routes):
    """Кожен пункт по управлінню — окремий витяг із читабельним ключем."""
    management = routes["management_paragraphs"]
    assert [
        [item["label"] for item in data["items"]] for data in management.values()
    ] == [["Пункт 1."], ["Пункт 2."]]
    assert list(management) == ["Управління — Пункт 1.", "Управління — Пункт 2."]


def test_management_extract_has_no_addressee(routes):
    """Адреси немає взагалі: такий витяг нікуди не надсилається."""
    for data in routes["management_paragraphs"].values():
        assert data["recipient_to"] == ""
        assert data["destination_where"] == ""
        assert data["open_name"] == "Управління"


def test_management_items_stay_out_of_the_distribution(routes):
    """Розрахунок розсилки не бачить управління ані рядком, ані адресатом."""
    assert routes["units_list"] == ["555 ітв А0555", "88 цпп А0088"]
    table_rows = [row[0] for row in routes["units_table"].rows]
    assert all("правлінн" not in str(cell) for cell in table_rows)

    generator = _load_generator()
    recipients = generator.build_message_recipient_list(MAPPING, routes)
    assert all("правлінн" not in recipient for recipient in recipients)

    # Пункт по управлінню не є «пропущеним»: він у виключених.
    assert [item["label"] for item in routes["skipped_items"]] == ["Пункт 1.", "Пункт 2."]


def test_management_items_keep_the_source_line_span(routes):
    """Витяг збирається з тих самих рядків наказу, що й звичайний."""
    for data in routes["management_paragraphs"].values():
        item = data["items"][0]
        assert item["source_end_line"] >= item["source_start_line"] >= 0
        assert item["text"].strip()


def test_management_output_file_is_named_separately():
    """Окремий файл видно за назвою, формат реквізитів — той самий."""
    generator = _load_generator()
    assert generator.build_extracts_filename("445", "02.09.2026") == (
        "Витяги наказу № 445 від 02.09.2026.docx"
    )
    assert generator.build_extracts_filename(
        "445", "02.09.2026", "Витяги до управління за наказом"
    ) == "Витяги до управління за наказом № 445 від 02.09.2026.docx"


def test_order_without_units_still_produces_management_extracts():
    """Наказ лише про управління не є порожнім результатом."""
    only_management = map_military_units(
        text=(
            "УКЛАСТИ КОНТРАКТИ:\n"
            "1. Підполковником ТЕСТЕНКОМ Олегом Васильовичем, офіцером відділу "
            "планування розвідувального управління штабу управління.\n"
        ),
        mapping=MAPPING,
    )
    assert only_management["unit_paragraphs"] == {}
    assert list(only_management["management_paragraphs"]) == ["Управління — Пункт 1."]


def test_extract_scope_separates_regular_and_management_buttons(routes):
    """Кожна кнопка бере лише свій вихід, а повний цикл — обидва."""
    generator = _load_generator()

    general, management = generator.select_extracts_for_scope(routes, "general")
    assert general == routes["unit_paragraphs"]
    assert management == {}

    general, management = generator.select_extracts_for_scope(routes, "management")
    assert general == {}
    assert management == routes["management_paragraphs"]

    general, management = generator.select_extracts_for_scope(routes, "all")
    assert general == routes["unit_paragraphs"]
    assert management == routes["management_paragraphs"]


def test_main_extract_button_uses_all_scope_by_default():
    """Основна кнопка створює адресний і управлінський файли разом."""
    generator = _load_generator()
    calls = []

    class Stub:
        def _run_extracts_scope_action(self, *args):
            calls.append(args)

    generator.App.run_extracts_action(Stub())

    assert calls == [
        ("all", "Усі витяги для", "Генерація всіх витягів", "генерацію всіх витягів")
    ]


def test_saved_reports_restore_management_as_ui_only_green_row(tmp_path):
    """Перемикання наказу читає XLSX і відновлює управління без перерахунку."""
    generator = _load_generator()
    order = tmp_path / "Наказ № 77 від 15.09.2026.docx"
    order.write_bytes(b"")
    output = tmp_path / "Extracts_Output"
    output.mkdir()
    base = generator.sanitize_filename(order.stem)

    generator._save_table_to_excel(
        str(output / f"Розрахунок_розсилки_{base}.xlsx"),
        ["Військова частина / Відправник", "Номери пунктів витягу", "Кількість пунктів"],
        [("555 ітв А0555", "Пункт 3.", 1)],
    )
    generator._save_table_to_excel(
        str(output / f"Контроль_пропущених_пунктів_{base}.xlsx"),
        ["Пункт", "Текст пункту", "Причина"],
        [],
    )
    generator._save_table_to_excel(
        str(output / f"Контроль_маршрутизації_{base}.xlsx"),
        [
            "Пункт", "Збіги з таблиці", "Застосовані правила",
            "Адресати з пункту", "Адресати з контексту", "Підсумкові адресати",
        ],
        [
            ("Пункт 1.", "—", "зміна до управління: витяг виключено із загального переліку", "—", "—", "—"),
            ("Пункт 3.", "555 інформаційно-телекомунікаційний вузол", "адресат знайдено", "555 ітв А0555", "—", "555 ітв А0555"),
        ],
    )

    saved = generator.load_saved_analysis_reports(str(order), str(output))

    assert saved["calculation"] == [
        ("555 ітв А0555", "Пункт 3.", 1),
        (generator.MANAGEMENT_RESULT_LABEL, "Пункт 1.", 1),
    ]
    assert saved["unmatched"] == []
    assert saved["routing"][0][0] == "Пункт 1."
    assert all(saved["found"].values())
