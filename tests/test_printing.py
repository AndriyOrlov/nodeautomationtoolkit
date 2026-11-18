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
    scan_batch_folders,
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
    # 05.11.2026: 1…n−2 → задня n → пауза → n−1 на звороті задньої.
    assert [(step.what.split(" — ")[0], step.pages, step.copy_three) for step in steps] == [
        ("Витяги", "", False),
        ("Примірник № 2", "1-2", False),
        ("Примірник № 2", "4", False),
        ("Примірник № 2", "3", False),
        ("Примірник № 3", "1-2", True),
        ("Примірник № 3", "4", True),
        ("Примірник № 3", "3", True),
    ]
    assert steps[0].two_per_sheet
    # Пауза після задньої сторінки — верхній аркуш у ручну подачу, з підказкою.
    assert not steps[1].after
    # 06.11.2026: між лицьовими 1…n−2 і задньою n паузи немає — той самий лоток.
    assert not steps[1].pause_after and steps[2].pause_after and steps[3].pause_after
    assert "ВЕРХНІЙ аркуш" in steps[2].after and "лицем догори" in steps[2].after
    assert not steps[3].after


def test_management_extracts_only_when_asked():
    order = OrderFiles("556", "15.08.2026", extracts="e.docx", management="m.docx")
    assert [step.path for step in plan_steps(order, 0, include_management=False)] == ["e.docx"]
    with_management = plan_steps(order, 0, include_management=True)
    assert [step.path for step in with_management] == ["e.docx", "m.docx"]
    assert not with_management[1].two_per_sheet


def test_two_page_copy_and_single_page_copy():
    order = OrderFiles("1", "01.10.2026", copy_two="c.docx")
    two = plan_steps(order, copy_pages=2, include_management=False)
    assert [step.pages for step in two] == ["2", "1", "2", "1"]
    assert two[0].after and not two[1].after
    three = plan_steps(order, copy_pages=3, include_management=False)
    assert [step.pages for step in three] == ["1", "3", "2", "1", "3", "2"]
    single = plan_steps(order, copy_pages=1, include_management=False)
    assert [(step.pages, step.copy_three) for step in single] == [("", False), ("", True)]


def test_management_extracts_print_two_copies():
    order = OrderFiles("556", "15.08.2026", management="m.docx")
    [step] = plan_steps(order, 0, include_management=True)
    assert step.copies == 2 and not step.two_per_sheet


def test_copy_sheets_are_one_sided_with_the_back_page_on_the_reverse():
    from nodeautomationtoolkit.generator.printing import copy_sheet_count

    assert [copy_sheet_count(pages) for pages in (0, 1, 2, 4, 7)] == [0, 1, 1, 3, 6]


def test_copies_and_extracts_from_two_folders_are_joined_by_order(tmp_path):
    copies = tmp_path / "Примірники"
    extracts = tmp_path / "Витяги"
    _touch(copies, build_copy_two_filename("556", "15.08.2026", "Наказ № 556 від 15.08.2026.docx"))
    _touch(copies, build_copy_two_filename("557", "16.08.2026", "Наказ № 557 від 16.08.2026.docx"))
    _touch(extracts / "Готово", build_extracts_filename("556", "15.08.2026"))
    _touch(extracts, build_extracts_filename("556", "15.08.2026", MANAGEMENT_PREFIX))
    orders = scan_batch_folders([str(copies), str(extracts), str(copies), ""])
    assert [(order.number, bool(order.copy_two), bool(order.extracts), bool(order.management)) for order in orders] == [
        ("556", True, True, True),
        ("557", True, False, False),
    ]


def test_print_queue_uses_both_folders(tmp_path):
    from types import SimpleNamespace

    from nodeautomationtoolkit.generator.app_printing import PrintingMixin

    class _Var:
        def __init__(self, value):
            self.value = value

        def get(self):
            return self.value

    copies = tmp_path / "Примірники"
    extracts = tmp_path / "Витяги"
    _touch(copies, build_copy_two_filename("556", "15.08.2026", "Наказ № 556 від 15.08.2026.docx"))
    _touch(extracts, build_extracts_filename("556", "15.08.2026"))
    shell = SimpleNamespace(
        print_folder=_Var(str(copies)),
        print_extracts_folder=_Var(str(extracts)),
        p2_out_folder=_Var(""),
        print_management=_Var(False),
    )
    shell._print_folder = lambda: PrintingMixin._print_folder(shell)
    shell._print_folders = lambda: PrintingMixin._print_folders(shell)
    shell._print_auto_duplex = lambda: PrintingMixin._print_auto_duplex(shell)
    shell._print_flag = lambda name: PrintingMixin._print_flag(shell, name)
    shell._print_queue_entries = lambda: PrintingMixin._print_queue_entries(shell)
    rows = PrintingMixin.print_queue_rows(shell)
    assert [row[1] for row in rows] == [
        "Витяги — 2 сторінки на аркуш",
        "Примірник № 2 + задня сторінка",
        "Примірник № 3 + задня сторінка (лише друк)",
    ]


def test_auto_duplex_prints_only_the_last_sheet_on_both_sides():
    order = OrderFiles("556", "15.08.2026", copy_two="c.docx")
    steps = plan_steps(order, copy_pages=4, include_management=False, auto_duplex=True)
    assert [(step.pages, step.duplex, step.pause_after, step.copy_three) for step in steps] == [
        ("1-2", False, False, False),  # однобічно, без паузи — продовження того самого примірника
        ("3-4", True, True, False),    # останній аркуш: 3 лицем, 4 — на звороті, принтер сам
        ("1-2", False, False, True),
        ("3-4", True, True, True),
    ]
    assert not any(step.after for step in steps)  # жодної ручної подачі
    two = plan_steps(order, copy_pages=2, include_management=False, auto_duplex=True)
    assert [(step.pages, step.duplex) for step in two] == [("1-2", True), ("1-2", True)]
    three = plan_steps(order, copy_pages=3, include_management=False, auto_duplex=True)
    assert [step.pages for step in three] == ["1", "2-3", "1", "2-3"]
    single = plan_steps(order, copy_pages=1, include_management=False, auto_duplex=True)
    assert [(step.pages, step.duplex) for step in single] == [("", False), ("", False)]


def test_printer_duplex_is_restored(monkeypatch):
    import types

    from nodeautomationtoolkit.generator import printing

    class DevMode:
        Duplex = 1
        Fields = 0

    user_devmode = DevMode()
    saved = []
    fake = types.SimpleNamespace(
        PRINTER_ACCESS_USE=8,
        OpenPrinter=lambda name, defaults: "h",
        ClosePrinter=lambda handle: None,
        GetPrinter=lambda handle, level: {"pDevMode": user_devmode} if level in (2, 9) else None,
        SetPrinter=lambda handle, level, info, command: saved.append((level, info["pDevMode"].Duplex)),
    )
    monkeypatch.setitem(sys.modules, "win32print", fake)
    duplex = printing.printer_duplex("Тестовий принтер")
    duplex.set(True)
    duplex.set(False)
    duplex.set(True)
    duplex.restore()
    # Лише власні налаштування користувача (рівень 9); наприкінці — як було (1, однобічно).
    assert saved == [(9, 2), (9, 1), (9, 2), (9, 1)]


def test_back_side_step_is_marked_and_names_the_tray():
    order = OrderFiles("556", "15.08.2026", copy_two="c.docx")
    steps = plan_steps(order, copy_pages=4, include_management=False, back_tray_name="Лоток подачі вручну")
    assert [(step.pages, step.back_side) for step in steps[:3]] == [("1-2", False), ("4", False), ("3", True)]
    assert "у лоток «Лоток подачі вручну»" in steps[1].after
    plain = plan_steps(order, copy_pages=4, include_management=False)
    assert "у ручну подачу" in plain[1].after


def test_paper_tray_is_set_on_every_section():
    from types import SimpleNamespace

    from nodeautomationtoolkit.generator.printing import set_paper_tray

    sections = [SimpleNamespace(PageSetup=SimpleNamespace(FirstPageTray=0, OtherPagesTray=0)) for _ in range(2)]
    doc = SimpleNamespace(Sections=sections)
    set_paper_tray(doc, 0)  # «як у принтера» — нічого не міняємо
    assert all(s.PageSetup.FirstPageTray == 0 for s in sections)
    set_paper_tray(doc, 4)
    assert all((s.PageSetup.FirstPageTray, s.PageSetup.OtherPagesTray) == (4, 4) for s in sections)


def test_queue_is_one_row_per_order_and_unchecked_orders_are_skipped(tmp_path):
    from types import SimpleNamespace

    from nodeautomationtoolkit.generator.app_printing import PrintingMixin

    class _Var:
        def __init__(self, value):
            self.value = value

        def get(self):
            return self.value

    folder = tmp_path / "Пакет"
    for number, date in (("556", "15.08.2026"), ("557", "16.08.2026")):
        _touch(folder, build_copy_two_filename(number, date, f"Наказ № {number} від {date}.docx"))
        _touch(folder, build_extracts_filename(number, date))

    class Shell(PrintingMixin):
        pass

    shell = Shell()
    shell.print_folder = _Var(str(folder))
    shell.print_extracts_folder = _Var("")
    shell.p2_out_folder = _Var("")
    shell.print_management = _Var(False)
    orders = shell.print_queue_orders()
    assert [(marked, title) for marked, title, _what, _files, _key in orders] == [
        (True, "Наказ № 556 від 15.08.2026"),
        (True, "Наказ № 557 від 16.08.2026"),
    ]
    shell.set_print_order_marked(orders[0][4], False)
    assert [marked for marked, *_rest in shell.print_queue_orders()] == [False, True]
    # Знята галочка лишається після оновлення черги, поставлена — повертає наказ.
    shell.set_print_order_marked(orders[0][4], True)
    assert [marked for marked, *_rest in shell.print_queue_orders()] == [True, True]


def test_pages_are_addressed_by_section_not_by_printed_number():
    """18.11.2026: у Word «Pages=1» — надрукований номер; у кількох розділах він повторюється."""
    from types import SimpleNamespace

    from nodeautomationtoolkit.generator.printing import absolute_pages_spec

    # 4 фізичні сторінки: розділ 1 — сторінки 1–2, розділ 2 — нумерація знову з 1.
    layout = {1: (1, 1), 2: (2, 1), 3: (1, 2), 4: (2, 2)}

    def go_to(_what, _which, number):
        adjusted, section = layout[number]
        return SimpleNamespace(Information=lambda kind: adjusted if kind == 1 else section)

    doc = SimpleNamespace(GoTo=go_to)
    assert absolute_pages_spec(doc, "3") == "p1s2"
    assert absolute_pages_spec(doc, "1-2") == "p1s1-p2s1"
    assert absolute_pages_spec(doc, "2-4") == "p2s1-p2s2"
    assert absolute_pages_spec(doc, "") == ""


def test_extracts_and_copies_can_be_printed_separately():
    """18.11.2026: витяги окремо від примірників № 2/3."""
    order = OrderFiles("556", "15.08.2026", extracts="e.docx", management="m.docx", copy_two="c.docx")
    only_copies = plan_steps(order, 2, include_management=False, include_extracts=False)
    assert {step.path for step in only_copies} == {"c.docx"}
    only_extracts = plan_steps(order, 2, include_management=True, include_copies=False)
    assert [step.path for step in only_extracts] == ["e.docx", "m.docx"]
    assert plan_steps(order, 2, include_management=False, include_extracts=False, include_copies=False) == []


def test_print_queue_follows_what_to_print(tmp_path):
    from nodeautomationtoolkit.generator.app_printing import PrintingMixin

    class _Var:
        def __init__(self, value):
            self.value = value

        def get(self):
            return self.value

    folder = tmp_path / "Пакет"
    _touch(folder, build_copy_two_filename("556", "15.08.2026", "Наказ № 556 від 15.08.2026.docx"))
    _touch(folder, build_extracts_filename("556", "15.08.2026"))

    class Shell(PrintingMixin):
        pass

    shell = Shell()
    shell.print_folder = _Var(str(folder))
    shell.print_extracts_folder = _Var("")
    shell.p2_out_folder = _Var("")
    shell.print_management = _Var(False)
    shell.print_extracts = _Var(False)
    shell.print_copies = _Var(True)
    assert all("Примірник" in row[1] for row in shell.print_queue_rows())
    shell.print_extracts, shell.print_copies = _Var(True), _Var(False)
    assert [row[1] for row in shell.print_queue_rows()] == ["Витяги — 2 сторінки на аркуш"]
