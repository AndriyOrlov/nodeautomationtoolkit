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
