"""Чотири етапи генератора наказів: індексація → генерація → перевірка → збірка.

Це тонкий шар над модулями, які роблять роботу, щоб і вікно, і тести, і
майбутній командний рядок ходили одним шляхом:

1. **Індексація** — `order_index.store.build_index`: папка зі старими наказами
   перетворюється на локальний індекс (посади, підрозділи, пункти про осіб).
   Індекс містить ПІБ і РНОКПП, тому лишається лише на машині користувача
   (PROJECT_RULES 1.1, у git його немає).
2. **Генерація** — джерела (план переміщення, подання чи інший документ,
   ручний ввід) зводяться в записи про осіб (`record` + `merge`), а записи —
   в текст наказу за зразками додатка 53 (`compose`).
3. **Перевірка** — `check`: що бракує, що не сходиться, де відмінок під
   сумнівом. Помилки не дають зібрати документ, попередження — дають.
4. **Збірка** — `build`: .docx за правилами верстки витягів (PROJECT_RULES 5).
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from ..order_index.store import IndexSummary, build_index, find_person_items
from . import merge, sources
from .build import OrderDocumentParts, build_order_document, order_filename
from .check import CheckResult, check_draft
from .compose import OrderDraft, OrderParams, compose_order
from .ocr import OcrUnavailable
from .plan import read_plan
from .record import PersonRecord
from .resolve import resolve_document

Log = Callable[[str], None]


@dataclass
class OrderRun:
    """Наслідок одного прогону: записи, проєкт, перевірка, зібраний файл."""

    records: list[PersonRecord] = field(default_factory=list)
    draft: OrderDraft | None = None
    check: CheckResult | None = None
    document: Path | None = None
    log: list[str] = field(default_factory=list)


# ─────────────────────────────── 1. Індексація ───────────────────────────────
def index_orders(
    orders_folder: str | Path,
    index_folder: str | Path,
    *,
    recursive: bool = True,
    force: bool = False,
    log: Log = lambda _message: None,
    progress: Callable[[int, int, str], None] = lambda _done, _total, _name: None,
    cancelled: Callable[[], bool] = lambda: False,
) -> IndexSummary:
    """Перебудовує локальний індекс наказів (крок 1)."""
    return build_index(
        orders_folder,
        index_folder,
        recursive=recursive,
        force=force,
        log=log,
        progress=progress,
        cancelled=cancelled,
    )


# ──────────────────────────────── 2. Генерація ───────────────────────────────
def records_from_plan(
    plan_path: str | Path,
    index_folder: str | Path | None = None,
    log: Log = lambda _m: None,
    action: str = merge.APPOINTMENT,
) -> list[PersonRecord]:
    """План переміщення + найсвіжіший пункт наказу про кожну особу з індексу."""
    plan = read_plan(plan_path)
    log(f"План переміщення: рядків {len(plan.entries)}")
    records = []
    for entry in plan.entries:
        plan_record = sources.from_plan(entry)
        order_record = None
        if index_folder:
            previous = find_person_items(
                index_folder,
                ipn=str(plan_record.ipn),
                full_name=plan_record.full_name,
                limit=1,
            )
            if previous:
                order_record = sources.from_order_item(previous[0])
            else:
                log(f"У наказах не знайдено: {plan_record.full_name}")
        records.append(merge.build_record(plan=plan_record, order=order_record, action=action))
    return records


def records_from_documents(
    paths: list[str | Path],
    index_folder: str | Path | None = None,
    log: Log = lambda _m: None,
    action: str = merge.APPOINTMENT,
) -> list[PersonRecord]:
    """Подання, рапорт, аркуш бесіди, скан — усе, що вміє читати `documents.py`."""
    records = []
    for path in paths:
        name = Path(path).name
        try:
            resolved = resolve_document(path, index_folder, action=action)
        except OcrUnavailable as error:
            # Скан без Tesseract — окремий випадок: решту файлів це не спиняє.
            log(f"{name}: не вдалося розпізнати скан. {error}")
            continue
        except FileNotFoundError:
            log(f"{name}: файла немає на місці — його перенесли або видалили.")
            continue
        except Exception as error:  # пошкоджений або захищений файл
            log(f"{name}: прочитати не вдалося ({type(error).__name__}). Відкрийте файл і перевірте.")
            continue
        log(f"{name}: {resolved.kind}, осіб {len(resolved.people)}")
        if not resolved.people:
            log(f"{name}: жодної особи не впізнано — допоможе ручний ввід.")
        records.extend(person.record for person in resolved.people)
    return records


def records_from_manual(
    rows: list[dict[str, str]], action: str = merge.APPOINTMENT
) -> list[PersonRecord]:
    """Ручний ввід: список словників з іменами полів `PersonRecord`."""
    return [
        merge.build_record(manual=sources.from_manual(values), action=action or merge.APPOINTMENT)
        for values in rows
    ]


def generate(records: list[PersonRecord], params: OrderParams | None = None) -> OrderDraft:
    """Записи про осіб → проєкт наказу (крок 2)."""
    return compose_order(records, params)


# ──────────────────────────────── 3. Перевірка ───────────────────────────────
def verify(draft: OrderDraft) -> CheckResult:
    """Перевіряє проєкт перед збіркою (крок 3)."""
    return check_draft(draft)


# ───────────────────────────────── 4. Збірка ─────────────────────────────────
def assemble(
    draft: OrderDraft,
    out_folder: str | Path,
    template_path: str | Path = "",
    parts: OrderDocumentParts | None = None,
    filename: str = "",
) -> Path:
    """Збирає документ Word у теку результату (крок 4)."""
    name = filename or order_filename(draft)
    return build_order_document(draft, Path(out_folder) / name, template_path, parts)


def run(
    records: list[PersonRecord],
    params: OrderParams,
    out_folder: str | Path,
    template_path: str | Path = "",
    parts: OrderDocumentParts | None = None,
    *,
    assemble_with_errors: bool = False,
    log: Log = lambda _message: None,
) -> OrderRun:
    """Генерація → перевірка → збірка. Помилки зупиняють збірку, попередження — ні."""
    result = OrderRun(records=list(records))
    result.draft = generate(records, params)
    result.check = verify(result.draft)
    for line in result.check.lines():
        log(line)
        result.log.append(line)
    if result.check.ready or assemble_with_errors:
        result.document = assemble(result.draft, out_folder, template_path, parts)
        log(f"Зібрано: {result.document}")
    else:
        log("Документ не зібрано: спершу виправте помилки, позначені ✖.")
    return result
