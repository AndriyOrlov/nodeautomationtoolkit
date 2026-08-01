from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import uuid
from pathlib import Path

from nodeautomationtoolkit.core.batch_types import DocumentVariant, WordDocumentBatch
from nodeautomationtoolkit.core.definition import node


def _required_docx(path: str, label: str) -> Path:
    candidate = Path(path).expanduser()
    if not path.strip() or not candidate.is_file():
        raise FileNotFoundError(f"{label} не знайдено: {path or '(шлях порожній)'}")
    if candidate.suffix.casefold() != ".docx":
        raise ValueError(f"{label} має бути DOCX: {candidate}")
    return candidate.resolve()


def _parse_variants(
    groups: dict | None,
    names: list | None,
    names_text: str,
    fields_json: str,
) -> tuple[DocumentVariant, ...]:
    variants: list[DocumentVariant] = []
    if groups:
        for name, content in groups.items():
            if isinstance(content, dict):
                values = dict(content)
                values.setdefault("name", str(name))
                values.setdefault("marker", str(name))
                if str(values.get("content", "")).strip():
                    variants.append(DocumentVariant(str(values["name"]), tuple(values.items())))
            elif str(content).strip():
                variants.append(
                    DocumentVariant(str(name), (("marker", str(name)), ("content", content)))
                )
    elif fields_json.strip():
        payload = json.loads(fields_json)
        if isinstance(payload, dict):
            for name, fields in payload.items():
                values = fields if isinstance(fields, dict) else {"value": fields}
                variants.append(DocumentVariant(str(name), tuple(values.items())))
        elif isinstance(payload, list):
            for index, item in enumerate(payload, start=1):
                if isinstance(item, dict):
                    values = dict(item)
                    name = str(values.pop("name", f"Документ {index}"))
                    variants.append(DocumentVariant(name, tuple(values.items())))
                else:
                    variants.append(DocumentVariant(str(item)))
        else:
            raise ValueError("Поля варіантів мають бути JSON-об'єктом або списком")
    elif names:
        variants = [DocumentVariant(str(item)) for item in names if str(item).strip()]
    else:
        variants = [
            DocumentVariant(line.strip()) for line in names_text.splitlines() if line.strip()
        ]
    if not variants:
        variants = [DocumentVariant("Витяг")]
    return tuple(variants)


@node(
    name="Створити пакет документів",
    category="Word · Пакет",
    description=(
        "Бере один DOCX і створює план декількох вихідних документів. "
        "Назви можна подати списком, рядками або JSON із полями."
    ),
    type_id="builtin.word_batch.create",
)
def create_document_batch(
    source_path: str = "",
    template_path: str = "",
    groups: dict | None = None,
    names: list | None = None,
    names_text: str = "Витяг",
    fields_json: str = "",
) -> WordDocumentBatch:
    source = _required_docx(
        template_path or source_path,
        "DOCX-заготовка" if template_path else "Вхідний документ",
    )
    return WordDocumentBatch(
        source_path=str(source),
        variants=_parse_variants(groups, names, names_text, fields_json),
    )


@node(
    name="Замінити текст у пакеті",
    category="Word · Пакет",
    description=(
        "Додає заміну для кожного документа. У новому тексті працюють поля {{name}} та поля з JSON."
    ),
    type_id="builtin.word_batch.replace_text",
)
def batch_replace_text(
    batch: WordDocumentBatch,
    find: str = "",
    replacement_text: str = "",
    ignore_case: bool = True,
    replace_all: bool = True,
) -> WordDocumentBatch:
    if not find:
        raise ValueError("Не вказано текст для пошуку")
    return batch.with_operation(
        "replace_text",
        find=find,
        replacement_text=replacement_text,
        ignore_case=ignore_case,
        replace_all=replace_all,
    )


@node(
    name="Заповнити плейсхолдер",
    category="Word · Пакет",
    description=(
        "Шукає мітку на кшталт {{ТЕКСТ}} і вставляє заданий багаторядковий текст "
        "у кожен документ пакета."
    ),
    type_id="builtin.word_batch.fill_placeholder",
)
def batch_fill_placeholder(
    batch: WordDocumentBatch,
    placeholder: str = "{{ТЕКСТ}}",
    replacement_text: str = "",
) -> WordDocumentBatch:
    if not placeholder:
        raise ValueError("Плейсхолдер не може бути порожнім")
    return batch.with_operation(
        "replace_text",
        find=placeholder,
        replacement_text=replacement_text,
        ignore_case=False,
        replace_all=True,
    )


@node(
    name="Замінити основну частину",
    category="Word · Пакет",
    description=(
        "У кожній копії наказу знаходить початок параграфів/причин/пунктів та "
        "межу перед підписом, після чого вставляє згрупований {{content}}."
    ),
    type_id="builtin.word_batch.replace_body",
)
def batch_replace_body(
    batch: WordDocumentBatch,
    start_pattern: str = (
        r"(?im)^\s*(?:§\s*\d+|параграф\s+\d+|розділ\s+[IVXLC\d]+|"
        r"відповідно|підстава|причина|у\s+зв'язку|на\s+підставі|"
        r"на\s+виконання|\d+(?:\.\d+)*[.)])"
    ),
    end_pattern: str = (
        r"(?im)^\s*(?:командир|начальник|заступник|керівник|голова|директор|"
        r"т\.в\.о\.|тимчасово\s+виконуючий)\b"
    ),
    replacement_text: str = "{{content}}",
) -> WordDocumentBatch:
    _compile_block_pattern(start_pattern, "початку")
    _compile_block_pattern(end_pattern, "кінця")
    return batch.with_operation(
        "replace_regex_block",
        start_pattern=start_pattern,
        end_pattern=end_pattern,
        replacement_text=replacement_text,
    )


@node(
    name="Остання сторінка з прикладу",
    category="Word · Пакет",
    description=(
        "Видаляє останню сторінку кожного документа й вставляє замість неї "
        "вміст указаного DOCX-прикладу. Потребує Microsoft Word."
    ),
    type_id="builtin.word_batch.replace_last_page",
)
def batch_replace_last_page(
    batch: WordDocumentBatch,
    example_page_path: str = "",
) -> WordDocumentBatch:
    example = _required_docx(example_page_path, "Приклад останньої сторінки")
    return batch.with_operation("replace_last_page", example_page_path=str(example))


def _normalize_page_selector(page: str | int) -> str | int:
    if isinstance(page, int):
        if page < 1:
            raise ValueError("Номер сторінки має бути більшим за нуль")
        return page
    value = str(page).strip().casefold()
    first_aliases = {"перша", "першу", "першої", "first", "початок", "на початку"}
    last_aliases = {"остання", "останню", "останньої", "last", "кінець", "в кінці"}
    if value in first_aliases:
        return "first"
    if value in last_aliases:
        return "last"
    try:
        page_number = int(value)
    except ValueError as error:
        raise ValueError("Сторінка має бути числом, словом 'перша' або словом 'остання'") from error
    if page_number < 1:
        raise ValueError("Номер сторінки має бути більшим за нуль")
    return page_number


@node(
    name="Видалити сторінку",
    category="Word · Пакет",
    description=(
        "Видаляє одну сторінку в кожному документі пакета. У полі page вкажіть "
        "номер, 'перша' або 'остання'. Точні межі сторінок визначає Microsoft Word."
    ),
    type_id="builtin.word_batch.delete_page",
)
def batch_delete_page(
    batch: WordDocumentBatch,
    page: str = "остання",
) -> WordDocumentBatch:
    return batch.with_operation("delete_page", page=_normalize_page_selector(page))


@node(
    name="Вставити сторінку",
    category="Word · Пакет",
    description=(
        "Вставляє нову сторінку в кожен документ пакета. 'перша' додає її на "
        "початок, 'остання' — у кінець, а номер — перед сторінкою з цим номером. "
        "Поле page_docx можна лишити порожнім для чистої сторінки або вказати "
        "DOCX-заготовку з готовим оформленням. Потребує Microsoft Word."
    ),
    type_id="builtin.word_batch.insert_page",
)
def batch_insert_page(
    batch: WordDocumentBatch,
    page: str = "остання",
    page_docx: str = "",
) -> WordDocumentBatch:
    template = ""
    if page_docx.strip():
        template = str(_required_docx(page_docx, "DOCX-заготовка сторінки"))
    return batch.with_operation(
        "insert_page",
        page=_normalize_page_selector(page),
        page_docx=template,
    )


@node(
    name="Замінити підпис або блок",
    category="Word · Пакет",
    description=(
        "Знаходить початкову мітку підпису, видаляє блок до кінцевої мітки "
        "або кінця документа та вставляє текст чи DOCX-приклад."
    ),
    type_id="builtin.word_batch.replace_signature",
)
def batch_replace_signature(
    batch: WordDocumentBatch,
    start_marker: str = "",
    end_marker: str = "",
    replacement_text: str = "",
    replacement_docx: str = "",
) -> WordDocumentBatch:
    if not start_marker:
        raise ValueError("Не вказано початкову мітку блока підпису")
    if not replacement_text and not replacement_docx:
        raise ValueError("Вкажіть новий текст підпису або DOCX-приклад")
    template = ""
    if replacement_docx:
        template = str(_required_docx(replacement_docx, "Приклад підпису"))
    return batch.with_operation(
        "replace_block",
        start_marker=start_marker,
        end_marker=end_marker,
        replacement_text=replacement_text,
        replacement_docx=template,
    )


@node(
    name="Форматувати сторінки",
    category="Word · Пакет",
    description=(
        "Змінює шрифт, розмір, вирівнювання та інтервали лише на вказаних "
        "сторінках. 0 в останній сторінці означає кінець документа."
    ),
    type_id="builtin.word_batch.format_pages",
)
def batch_format_pages(
    batch: WordDocumentBatch,
    first_page: int = 1,
    last_page: int = 0,
    font_name: str = "",
    font_size: float = 0.0,
    line_spacing: float = 1.0,
    space_before: float = 0.0,
    space_after: float = 0.0,
    alignment: str = "не змінювати",
    keep_together: bool = True,
) -> WordDocumentBatch:
    if first_page < 1 or last_page < 0:
        raise ValueError("Номери сторінок мають бути додатними")
    if last_page and last_page < first_page:
        raise ValueError("Остання сторінка не може бути перед першою")
    if line_spacing <= 0:
        raise ValueError("Міжрядковий інтервал має бути більшим за нуль")
    return batch.with_operation(
        "format_pages",
        first_page=first_page,
        last_page=last_page,
        font_name=font_name,
        font_size=font_size,
        line_spacing=line_spacing,
        space_before=space_before,
        space_after=space_after,
        alignment=alignment,
        keep_together=keep_together,
    )


@node(
    name="Не розривати пункти",
    category="Word · Пакет",
    description=(
        "Позначає кожен нумерований пункт як неподільний блок, щоб Word не переносив "
        "половину пункту на наступну сторінку."
    ),
    type_id="builtin.word_batch.keep_items_together",
)
def batch_keep_items_together(
    batch: WordDocumentBatch,
    item_pattern: str = r"^\s*\d+(?:\.\d+)*[.)]\s*",
    signature_pattern: str = (
        r"^\s*(?:командир|начальник|заступник|керівник|голова|директор|"
        r"т\.в\.о\.|тимчасово\s+виконуючий)\b"
    ),
) -> WordDocumentBatch:
    _compile_item_pattern(item_pattern)
    _compile_item_pattern(signature_pattern)
    return batch.with_operation(
        "keep_items_together",
        item_pattern=item_pattern,
        signature_pattern=signature_pattern,
    )


def _expand_fields(text: str, variant: DocumentVariant) -> str:
    result = text
    for key, value in variant.values().items():
        result = result.replace("{{" + str(key) + "}}", str(value))
    return result


def _safe_file_stem(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" .")
    return cleaned or "Документ"


def _compile_item_pattern(pattern: str) -> re.Pattern:
    try:
        return re.compile(pattern)
    except re.error as error:
        raise ValueError(f"Некоректне правило пунктів: {error}") from error


def _compile_block_pattern(pattern: str, label: str) -> re.Pattern:
    try:
        return re.compile(pattern)
    except re.error as error:
        raise ValueError(f"Некоректне правило {label} блока: {error}") from error


def _output_paths(
    batch: WordDocumentBatch,
    output_folder: Path,
    filename_template: str,
    overwrite: bool,
) -> list[Path]:
    result = []
    used: set[str] = set()
    for variant in batch.variants:
        filename = _expand_fields(filename_template, variant)
        path = Path(filename)
        stem = _safe_file_stem(path.stem)
        suffix = ".docx"
        candidate = output_folder / f"{stem}{suffix}"
        index = 2
        while candidate.name.casefold() in used:
            candidate = output_folder / f"{stem} ({index}){suffix}"
            index += 1
        used.add(candidate.name.casefold())
        if candidate.exists() and not overwrite:
            raise FileExistsError(f"Файл уже існує: {candidate}")
        result.append(candidate)
    return result


def _find_word_range(document, text: str, start: int = 0, end: int | None = None):
    search_range = document.Range(Start=start, End=end or document.Content.End - 1)
    finder = search_range.Find
    finder.ClearFormatting()
    finder.Text = text
    finder.Forward = True
    finder.Wrap = 0
    finder.Format = False
    return search_range if finder.Execute() else None


def _replace_word_text(
    document,
    find_text: str,
    replacement_text: str,
    *,
    ignore_case: bool,
    replace_all: bool,
) -> int:
    count = 0
    search_start = 0
    while search_start < document.Content.End - 1:
        search_range = document.Range(Start=search_start, End=document.Content.End - 1)
        finder = search_range.Find
        finder.ClearFormatting()
        finder.Text = find_text
        finder.MatchCase = not ignore_case
        finder.Forward = True
        finder.Wrap = 0
        finder.Format = False
        if not finder.Execute():
            break
        search_range.Text = replacement_text.replace("\n", "\r")
        count += 1
        search_start = search_range.End
        if not replace_all:
            break
    return count


def _page_range(document, first_page: int, last_page: int):
    page_count = document.ComputeStatistics(2)
    if first_page > page_count:
        raise ValueError(f"У документі лише {page_count} сторінок")
    resolved_last = page_count if last_page == 0 else min(last_page, page_count)
    start = document.GoTo(What=1, Which=1, Count=first_page).Start
    if resolved_last < page_count:
        end = document.GoTo(What=1, Which=1, Count=resolved_last + 1).Start
    else:
        end = document.Content.End - 1
    return document.Range(Start=start, End=end)


def _resolve_page_number(document, selector: str | int, *, allow_append: bool = False) -> int:
    page_count = int(document.ComputeStatistics(2))
    normalized = _normalize_page_selector(selector)
    if normalized == "first":
        return 1
    if normalized == "last":
        return page_count + 1 if allow_append else page_count
    maximum = page_count + 1 if allow_append else page_count
    if normalized > maximum:
        if allow_append:
            raise ValueError(
                f"У документі {page_count} сторінок; вставити можна не далі сторінки "
                f"{page_count + 1}"
            )
        raise ValueError(f"У документі лише {page_count} сторінок")
    return normalized


def _insert_page(document, selector: str | int, page_docx: str = "") -> None:
    document.Repaginate()
    page_number = _resolve_page_number(document, selector, allow_append=True)
    page_count = int(document.ComputeStatistics(2))
    if page_number == page_count + 1:
        insertion_position = document.Content.End - 1
        document.Range(
            Start=insertion_position,
            End=insertion_position,
        ).InsertBreak(7)
        if page_docx:
            insertion_position = document.Content.End - 1
            document.Range(
                Start=insertion_position,
                End=insertion_position,
            ).InsertFile(page_docx)
    else:
        insertion_position = document.GoTo(
            What=1,
            Which=1,
            Count=page_number,
        ).Start
        document.Range(
            Start=insertion_position,
            End=insertion_position,
        ).InsertBreak(7)
        if page_docx:
            document.Range(
                Start=insertion_position,
                End=insertion_position,
            ).InsertFile(page_docx)
    document.Repaginate()


def _apply_word_operation(application, document, operation, variant: DocumentVariant) -> None:
    options = operation.options()
    if operation.kind == "replace_text":
        replacement = _expand_fields(options["replacement_text"], variant)
        _replace_word_text(
            document,
            options["find"],
            replacement,
            ignore_case=options["ignore_case"],
            replace_all=options["replace_all"],
        )
        return
    if operation.kind == "replace_last_page":
        document.Repaginate()
        page_count = document.ComputeStatistics(2)
        target = _page_range(document, page_count, page_count)
        target.Delete()
        target.InsertFile(options["example_page_path"])
        return
    if operation.kind == "delete_page":
        document.Repaginate()
        page_number = _resolve_page_number(document, options["page"])
        _page_range(document, page_number, page_number).Delete()
        document.Repaginate()
        return
    if operation.kind == "insert_page":
        _insert_page(document, options["page"], options["page_docx"])
        return
    if operation.kind == "replace_block":
        start_range = _find_word_range(document, options["start_marker"])
        if start_range is None:
            raise ValueError(f"Не знайдено початок блока: {options['start_marker']}")
        start = start_range.Start
        if options["end_marker"]:
            end_range = _find_word_range(
                document,
                options["end_marker"],
                start=start_range.End,
            )
            if end_range is None:
                raise ValueError(f"Не знайдено кінець блока: {options['end_marker']}")
            end = end_range.End
        else:
            end = document.Content.End - 1
        target = document.Range(Start=start, End=end)
        target.Delete()
        if options["replacement_docx"]:
            target.InsertFile(options["replacement_docx"])
        else:
            target.Text = _expand_fields(options["replacement_text"], variant).replace("\n", "\r")
        return
    if operation.kind == "replace_regex_block":
        content = str(document.Content.Text)
        start_match = _compile_block_pattern(options["start_pattern"], "початку").search(content)
        if start_match is None:
            raise ValueError("Не знайдено початок основної частини документа")
        end_match = _compile_block_pattern(options["end_pattern"], "кінця").search(
            content,
            start_match.end(),
        )
        if end_match is None:
            raise ValueError("Не знайдено межу перед підписом")
        target = document.Range(Start=start_match.start(), End=end_match.start())
        replacement = _expand_fields(options["replacement_text"], variant)
        target.Text = replacement.replace("\n", "\r") + "\r"
        return
    if operation.kind == "format_pages":
        target = _page_range(document, options["first_page"], options["last_page"])
        if options["font_name"]:
            target.Font.Name = options["font_name"]
        if options["font_size"] > 0:
            target.Font.Size = options["font_size"]
        paragraph_format = target.ParagraphFormat
        spacing = options["line_spacing"]
        if abs(spacing - 1.0) < 0.001:
            paragraph_format.LineSpacingRule = 0
        elif abs(spacing - 1.5) < 0.001:
            paragraph_format.LineSpacingRule = 1
        elif abs(spacing - 2.0) < 0.001:
            paragraph_format.LineSpacingRule = 2
        else:
            paragraph_format.LineSpacingRule = 5
            paragraph_format.LineSpacing = application.LinesToPoints(spacing)
        paragraph_format.SpaceBefore = options["space_before"]
        paragraph_format.SpaceAfter = options["space_after"]
        paragraph_format.KeepTogether = bool(options["keep_together"])
        alignment = options["alignment"].strip().casefold()
        alignments = {
            "ліворуч": 0,
            "по центру": 1,
            "центр": 1,
            "праворуч": 2,
            "по ширині": 3,
        }
        if alignment in alignments:
            paragraph_format.Alignment = alignments[alignment]
        return
    if operation.kind == "keep_items_together":
        pattern = _compile_item_pattern(options["item_pattern"])
        signature = _compile_item_pattern(options["signature_pattern"])
        paragraphs = list(document.Paragraphs)
        starts = [
            index
            for index, paragraph in enumerate(paragraphs)
            if pattern.search(str(paragraph.Range.Text).rstrip("\r\x07"))
        ]
        for position, start in enumerate(starts):
            stop = starts[position + 1] if position + 1 < len(starts) else len(paragraphs)
            for index in range(start + 1, stop):
                text = str(paragraphs[index].Range.Text).rstrip("\r\x07")
                if signature.search(text):
                    stop = index
                    break
            for index in range(start, stop):
                paragraph_format = paragraphs[index].Range.ParagraphFormat
                paragraph_format.KeepTogether = True
                paragraph_format.KeepWithNext = index < stop - 1
        return
    raise ValueError(f"Невідома пакетна операція: {operation.kind}")


def _word_application():
    try:
        import win32com.client
    except ImportError as error:
        raise RuntimeError(
            "Для сторінок і точного форматування потрібні Windows, Microsoft Word та pywin32"
        ) from error
    try:
        application = win32com.client.DispatchEx("Word.Application")
    except Exception as error:  # noqa: BLE001 - COM activation boundary
        raise RuntimeError(
            "Не вдалося запустити Microsoft Word. Перевірте, чи він встановлений."
        ) from error
    application.Visible = False
    application.DisplayAlerts = 0
    return application


def _save_batch_with_word(
    batch: WordDocumentBatch,
    targets: list[Path],
) -> list[int]:
    source = Path(batch.source_path)
    temporary_paths = [
        target.parent / f".nat-{uuid.uuid4().hex}-{target.name}" for target in targets
    ]
    application = _word_application()
    paragraph_counts: list[int] = []
    try:
        for temporary, variant in zip(temporary_paths, batch.variants, strict=True):
            shutil.copy2(source, temporary)
            document = application.Documents.Open(
                str(temporary),
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            try:
                for operation in batch.operations:
                    _apply_word_operation(application, document, operation, variant)
                paragraph_counts.append(int(document.Paragraphs.Count))
                document.Save()
            finally:
                document.Close(SaveChanges=False)
        for temporary, target in zip(temporary_paths, targets, strict=True):
            os.replace(temporary, target)
    finally:
        application.Quit()
        for temporary in temporary_paths:
            temporary.unlink(missing_ok=True)
    return paragraph_counts


def _replace_paragraph_text(
    paragraph,
    find_text: str,
    replacement_text: str,
    *,
    ignore_case: bool,
    replace_all: bool,
) -> int:
    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text or not find_text:
        return 0
    flags = re.IGNORECASE if ignore_case else 0
    matches = list(re.finditer(re.escape(find_text), full_text, flags))
    if not replace_all:
        matches = matches[:1]
    count = 0
    for match in reversed(matches):
        spans = []
        position = 0
        for index, run in enumerate(paragraph.runs):
            end = position + len(run.text)
            spans.append((index, position, end))
            position = end
        touched = [item for item in spans if item[1] < match.end() and item[2] > match.start()]
        if not touched:
            continue
        first_index, first_start, _first_end = touched[0]
        last_index, last_start, _last_end = touched[-1]
        first_run = paragraph.runs[first_index]
        last_run = paragraph.runs[last_index]
        prefix = first_run.text[: match.start() - first_start]
        suffix = last_run.text[match.end() - last_start :]
        if first_index == last_index:
            first_run.text = prefix + replacement_text + suffix
        else:
            first_run.text = prefix + replacement_text
            for index in range(first_index + 1, last_index):
                paragraph.runs[index].text = ""
            last_run.text = suffix
        count += 1
    return count


def _iter_document_paragraphs(document):
    seen_cells: set[int] = set()

    def table_paragraphs(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_id = id(cell._tc)
                    if cell_id in seen_cells:
                        continue
                    seen_cells.add(cell_id)
                    yield from cell.paragraphs
                    yield from table_paragraphs(cell.tables)

    yield from document.paragraphs
    yield from table_paragraphs(document.tables)
    for section in document.sections:
        for part in (section.header, section.footer):
            yield from part.paragraphs
            yield from table_paragraphs(part.tables)


def _save_batch_portable(batch: WordDocumentBatch, targets: list[Path]) -> list[int]:
    from docx import Document

    source_bytes = Path(batch.source_path).read_bytes()
    temporary_paths = [
        target.parent / f".nat-{uuid.uuid4().hex}-{target.name}" for target in targets
    ]
    paragraph_counts = []
    try:
        for temporary, variant in zip(temporary_paths, batch.variants, strict=True):
            temporary.write_bytes(source_bytes)
            document = Document(temporary)
            for operation in batch.operations:
                if operation.kind == "keep_items_together":
                    pattern = _compile_item_pattern(operation.options()["item_pattern"])
                    signature = _compile_item_pattern(operation.options()["signature_pattern"])
                    paragraphs = list(_iter_document_paragraphs(document))
                    starts = [
                        index
                        for index, paragraph in enumerate(paragraphs)
                        if pattern.search(paragraph.text)
                    ]
                    for position, start in enumerate(starts):
                        stop = (
                            starts[position + 1] if position + 1 < len(starts) else len(paragraphs)
                        )
                        for index in range(start + 1, stop):
                            if signature.search(paragraphs[index].text):
                                stop = index
                                break
                        for index in range(start, stop):
                            paragraphs[index].paragraph_format.keep_together = True
                            paragraphs[index].paragraph_format.keep_with_next = index < stop - 1
                    continue
                if operation.kind != "replace_text":
                    raise RuntimeError("Ця операція потребує встановленого Microsoft Word")
                options = operation.options()
                replacement = _expand_fields(options["replacement_text"], variant)
                remaining_one = not options["replace_all"]
                for paragraph in _iter_document_paragraphs(document):
                    replaced = _replace_paragraph_text(
                        paragraph,
                        options["find"],
                        replacement,
                        ignore_case=options["ignore_case"],
                        replace_all=not remaining_one,
                    )
                    if remaining_one and replaced:
                        break
            paragraph_counts.append(len(document.paragraphs))
            document.save(temporary)
        for temporary, target in zip(temporary_paths, targets, strict=True):
            os.replace(temporary, target)
    finally:
        for temporary in temporary_paths:
            temporary.unlink(missing_ok=True)
    return paragraph_counts


@node(
    name="Повне прев'ю у Word",
    category="Word · Пакет",
    description=(
        "Збирає перший документ пакета у тимчасовий DOCX і відкриває його в "
        "Microsoft Word з реальними полями, таблицями, відступами й сторінками."
    ),
    type_id="builtin.word_batch.preview_in_word",
    outputs={"path": "str", "name": "str", "message": "str"},
    execution_inputs=("exec",),
    preview_policy="never",
)
def preview_document_batch(batch: WordDocumentBatch) -> dict:
    if not batch.variants:
        raise ValueError("У пакеті немає документів для перегляду")
    preview_folder = Path(tempfile.gettempdir()) / "NodeAutomationToolkit" / "Preview"
    preview_folder.mkdir(parents=True, exist_ok=True)
    variant = batch.variants[0]
    preview_batch = WordDocumentBatch(
        source_path=batch.source_path,
        variants=(variant,),
        operations=batch.operations,
    )
    target = preview_folder / f"{_safe_file_stem(variant.name)}-{uuid.uuid4().hex[:8]}.docx"
    _save_batch_with_word(preview_batch, [target])
    if not hasattr(os, "startfile"):
        raise RuntimeError("Повне прев'ю потребує Windows і Microsoft Word")
    os.startfile(str(target))  # type: ignore[attr-defined]  # noqa: S606
    return {
        "path": str(target),
        "name": target.name,
        "message": "Прев'ю відкрито у Microsoft Word",
    }


@node(
    name="Зберегти пакет DOCX",
    category="Word · Пакет",
    description=(
        "Фінальна нода: за один запуск застосовує всі правила й створює всі DOCX у вибраній папці."
    ),
    type_id="builtin.word_batch.save",
    outputs={"paths": "List", "count": "int", "message": "str"},
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def save_document_batch(
    batch: WordDocumentBatch,
    output_folder: str = "",
    filename_template: str = "{{name}}.docx",
    overwrite: bool = False,
) -> dict:
    if not output_folder.strip():
        raise ValueError("Не вибрано папку для збереження")
    destination = Path(output_folder).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    targets = _output_paths(batch, destination, filename_template, overwrite)
    portable_operations = {"replace_text", "keep_items_together"}
    requires_word = any(operation.kind not in portable_operations for operation in batch.operations)
    requires_word = requires_word or any(
        operation.kind == "replace_text"
        and any(
            "\n"
            in _expand_fields(
                operation.options()["replacement_text"],
                variant,
            )
            for variant in batch.variants
        )
        for operation in batch.operations
    )
    if requires_word:
        _save_batch_with_word(batch, targets)
    else:
        _save_batch_portable(batch, targets)
    return {
        "paths": [str(path) for path in targets],
        "count": len(targets),
        "message": f"Створено документів: {len(targets)}",
    }
