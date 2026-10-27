"""Черга друку пакета (без Word): які файли, у якому порядку, які сторінки.

Назви файлів будуються тими самими функціями, що й у генераторі; дані вигадані.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT / "src") not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT / "src"))

from nodeautomationtoolkit.generator.naming import (  # noqa: E402
    build_copy_two_filename,
    build_extracts_filename,
)
from nodeautomationtoolkit.generator.printing import (  # noqa: E402
    MANAGEMENT_PREFIX,
    OrderFiles,
    plan_steps,
    scan_batch_folder,
)


def _touch(folder: Path, name: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / name
    path.write_bytes(b"")
    return path


def _batch(tmp_path: Path) -> Path:
    for number, date in (("557", "16.08.2026"), ("556", "15.08.2026")):
        folder = tmp_path / f"Наказ {number}"
        _touch(folder, build_copy_two_filename(number, date, f"Наказ № {number} від {date}.docx"))
        _touch(folder, build_extracts_filename(number, date))
        _touch(folder, build_extracts_filename(number, date, MANAGEMENT_PREFIX))
        _touch(folder, "~$тимчасовий файл Word.docx")
    _touch(tmp_path, "Нотатки без номера.docx")
    return tmp_path


def test_files_are_grouped_by_order_and_sorted_by_date(tmp_path):
    orders = scan_batch_folder(str(_batch(tmp_path)))
    assert [(order.number, order.date) for order in orders] == [("556", "15.08.2026"), ("557", "16.08.2026")]
    first = orders[0]
    assert Path(first.extracts).name.startswith("Витяги наказу № 556")
    assert Path(first.management).name.startswith(MANAGEMENT_PREFIX)
    assert Path(first.copy_two).name.startswith("2,3_")


def test_copy_back_page_goes_on_the_reverse_of_the_last_text_page():
    order = OrderFiles("556", "15.08.2026", extracts="e.docx", copy_two="c.docx")
    steps = plan_steps(order, copy_pages=4, include_management=False, feed_hint="лицем догори")
    assert [(step.what.split(" — ")[0], step.pages, step.copy_three) for step in steps] == [
        ("Витяги", "", False),
        ("Примірник № 2", "1-3", False),
        ("Примірник № 2", "4", False),
        ("Примірник № 3", "1-3", True),
        ("Примірник № 3", "4", True),
    ]
    assert steps[0].two_per_sheet
    # Пауза перед задньою сторінкою — верхній аркуш у ручну подачу, з підказкою.
    assert "ВЕРХНІЙ аркуш" in steps[1].after and "лицем догори" in steps[1].after
    assert not steps[2].after


def test_management_extracts_only_when_asked():
    order = OrderFiles("556", "15.08.2026", extracts="e.docx", management="m.docx")
    assert [step.path for step in plan_steps(order, 0, include_management=False)] == ["e.docx"]
    with_management = plan_steps(order, 0, include_management=True)
    assert [step.path for step in with_management] == ["e.docx", "m.docx"]
    assert not with_management[1].two_per_sheet


def test_two_page_copy_and_single_page_copy():
    order = OrderFiles("1", "01.10.2026", copy_two="c.docx")
    two = plan_steps(order, copy_pages=2, include_management=False)
    assert [step.pages for step in two] == ["1", "2", "1", "2"]
    single = plan_steps(order, copy_pages=1, include_management=False)
    assert [(step.pages, step.copy_three) for step in single] == [("", False), ("", True)]


def test_management_extracts_print_two_copies():
    order = OrderFiles("556", "15.08.2026", management="m.docx")
    [step] = plan_steps(order, 0, include_management=True)
    assert step.copies == 2 and not step.two_per_sheet


def test_copy_sheets_are_one_sided_with_the_back_page_on_the_reverse():
    from nodeautomationtoolkit.generator.printing import copy_sheet_count

    assert [copy_sheet_count(pages) for pages in (0, 1, 2, 4, 7)] == [0, 1, 1, 3, 6]
