from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    groups_to_ciphers,
    read_recipient_mapping,
)
from nodeautomationtoolkit.core.table_types import DataTable


def test_reads_csv_mapping_and_builds_markers(tmp_path):
    source = tmp_path / "mapping.csv"
    source.write_text(
        "Відкрите найменування;Шифр;Куди направляється\n"
        "АЛЬФА;Ш-01;Відділ 1\nБРАВО;Ш-01;Відділ 1\n",
        encoding="utf-8",
    )

    result = read_recipient_mapping(str(source))

    assert result["markers"] == ["АЛЬФА", "БРАВО"]
    assert result["mapping"]["АЛЬФА"]["cipher"] == "Ш-01"
    assert isinstance(result["table"], DataTable)


def test_finds_headers_below_a_title_row(tmp_path):
    source = tmp_path / "mapping_with_title.csv"
    source.write_text(
        "ТЕСТОВА ТАБЛИЦЯ;;;\n"
        "Відкрите найменування;Шифр;Куди направляється;Примітка\n"
        "99 тестова бригада;ТЕСТ-А9001;Тестове управління;Вигадано\n",
        encoding="utf-8",
    )

    result = read_recipient_mapping(str(source))

    assert result["count"] == 1
    assert result["mapping"]["99 тестова бригада"]["cipher"] == "ТЕСТ-А9001"


def test_converts_and_merges_groups_with_same_cipher():
    result = groups_to_ciphers(
        groups={"АЛЬФА": "1. Перший", "БРАВО": "2. Другий", "НЕМАЄ": "3. Третій"},
        counts={"АЛЬФА": 1, "БРАВО": 1, "НЕМАЄ": 1},
        mapping={
            "АЛЬФА": {"cipher": "Ш-01", "destination": "Відділ 1"},
            "БРАВО": {"cipher": "Ш-01", "destination": "Відділ 1"},
        },
    )

    assert list(result["documents"]) == ["Ш-01"]
    assert result["documents"]["Ш-01"]["count"] == 2
    assert "1. Перший" in result["documents"]["Ш-01"]["content"]
    assert "2. Другий" in result["documents"]["Ш-01"]["content"]
    assert result["missing"] == ["НЕМАЄ"]
    assert len(result["report"].rows) == 3


def test_analyze_senders_and_split_by_senders():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
        analyze_senders,
        split_by_senders,
    )

    text = (
        "НАКАЗ командира військової частини А0000\n"
        "15 січня 2026 року № 100, м. Київ\n\n"
        "§ 1\n"
        "1. Пункт про 160 окрему механізовану бригаду призначити.\n"
        "2. Пункт про 167 окрему механізовану бригаду звільнити.\n"
    )

    mapping = {
        "160 окрема механізована бригада": {"open_name": "160 окрема механізована бригада", "cipher": "в/ч А1600"},
        "167 окрема механізована бригада": {"open_name": "167 окрема механізована бригада", "cipher": "в/ч А1670"},
    }

    res = analyze_senders(text=text, mapping=mapping)
    assert isinstance(res["sender_paragraphs"], dict)
    assert len(res["senders_list"]) > 0
    assert isinstance(res["table"], DataTable)

    res_split = split_by_senders(text=text, mapping=mapping)
    assert isinstance(res_split["blocks"], dict)
    assert res_split["senders_count"] > 0


def test_army_corps_prioritization_over_subordinate_units():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import analyze_senders

    text = (
        "§ 1\n"
        "1. Молодшого лейтенанта призначити у 167 окрему механізовану бригаду 10-го армійського корпусу.\n"
        "2. Сержанта звільнити з 154 окремої механізованої бригади 3 АК.\n"
    )

    res = analyze_senders(text=text)
    assert "10 армійський корпус" in res["senders_list"]
    assert "3 армійський корпус" in res["senders_list"]
    assert len(res["sender_paragraphs"]["10 армійський корпус"]) == 1


def test_tck_full_wording_extracts_only_oblast():
    from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _extract_tck_sender, analyze_senders

    text = (
        "§ 1\n"
        "1. Направити документи до Ковельського районного територіального центру "
        "комплектування та соціальної підтримки Волинської області.\n"
    )

    extracted = _extract_tck_sender(text)
    assert extracted == "Волинський ОТЦК та СП"

    analysis = analyze_senders(text=text)
    assert "Волинський ОТЦК та СП" in analysis["senders_list"]
