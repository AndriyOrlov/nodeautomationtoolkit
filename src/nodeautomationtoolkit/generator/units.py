"""Відкриті назви частин, яких немає в таблиці, і заготовки рядків для неї.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path
import openpyxl
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    _build_unit_fuzzy_pattern,
    keep_open_names_of,
)
from nodeautomationtoolkit.builtin_nodes.message_order import (
    cipher_unit_names,
)



# Вид частини з номером: «169 батальйону резерву», «12 навчального центру»,
# «300 військового госпіталю». Без цього жовтим позначались лише бригади,
# полки й «окремі …», а решта нових частин лишалась у закритому повідомленні
# відкритою БЕЗ жодної позначки (розд. 9.5.7).
_NUMBERED_UNIT_KIND = (
    r"(?:бригад(?:а|и|і|у|ою)|полк(?:у|ом|і|ові)?|батальйон(?:у|і|ом|ові)?"
    r"|дивізіон(?:у|і|ом|ові)?|загін|загон(?:у|і|ом)|центр(?:у|і|ом)?"
    r"|госпітал(?:ь|ю|і|ем)|баз(?:а|и|і|у|ою)|вуз(?:ол|ла|лу|лі|лом)"
    r"|арсенал(?:у|і|ом)?|корпус(?:у|і|ом)?|дивізі(?:я|ї|ю|єю)|комендатур(?:а|и|і|у|ою)"
    # Види зі словника користувача: станція та вузол ФПЗ, ремонтна майстерня,
    # ремонтна майстерня. «Частина» до переліку НЕ входить: вона ловила б
    # «до пункту 2 частини четвертої статті» — у картографічної частини своє правило.
    r"|станці(?:я|ї|ю|єю)|майстерн(?:я|і|ю|ею))"
)


# Слово між номером і видом частини. Номер підрозділу («2 відділу», «3 взводу»)
# і номер статті/пункту частиною не є — такі слова проміжок не пропускає.
_NUMBERED_UNIT_GAP_WORD = (
    r"(?!(?:року|рік|років|відділ\w*|взвод\w*|рот[аиіу]|роті|батаре\w*|груп\w*|служб\w*"
    r"|пункт\w*|частин\w*|стат\w*|наказ\w*)\b)"
    r"[а-яіїєґʼ'’-]+"
)


_UNMATCHED_OPEN_UNIT_RE = re.compile(
    r"\b(?:"
    r"(?:\d{1,3}\s*(?:-?[а-яіїєґ]+)?\s*)?(?:окрем\w+\s+)+(?:механізован\w+|танков\w+|десантн\w+|артилерійськ\w+|піхотн\w+|єгерськ\w+|стрілецьк\w+|штурмов\w+|розвідувальн\w+|гірсько-штурмов\w+|десантно-штурмов\w+|аеромобільн\w+|повітряно-десантн\w+|зв['’]язку)?\s*(?:бригад\w*|полк\w*|батальйон\w*|дивізіон\w*|загін\w*|центр\w*)"
    r"|"
    r"\d{1,3}\s*(?:-?[а-яіїєґ]+)?\s*(?:механізован\w+|танков\w+|десантн\w+|артилерійськ\w+|піхотн\w+|стрілецьк\w+|штурмов\w+|десантно-штурмов\w+|аеромобільн\w+)?\s*(?:бригад\w*|полк\w*|армійськ\w+\s+корпус\w*|АК)"
    r"|"
    r"(?:армійськ\w+\s+корпус\w*|\b\d{1,3}\s*АК\b)"
    r"|"
    # «10 командного пункту протиповітряної оборони», «командно-розвідувальний пункт».
    # Саме слово «пункт» у наказі означає пункт наказу, тому поруч потрібне «командн…».
    r"(?<![\d./,:-])\d{1,4}\s+(?:[\w’'ʼ-]+\s+){0,3}?команд\w*[-\s]?\w*\s+пункт\w*"
    r"|"
    # «картографічна частина» — лише з цим прикметником, інакше ловилося б
    # «частини четвертої статті».
    r"(?:картографічн\w+\s+частин\w*)"
    r"|"
    # «окремої танкової Тестівської бригади» — топонім між словами назви (додаток 53)
    r"(?:окрем\w+\s+)(?:[\w’'ʼ-]+\s+){1,3}?(?:бригад\w*|полк\w*|батальйон\w*|дивізіон\w*|центр\w*)"
    r"|"
    r"(?<![\d./,:-])\d{1,4}(?:\s*-?\s*(?:й|го|му|ий|ого|ому|им))?\s+"
    rf"(?:{_NUMBERED_UNIT_GAP_WORD}\s+){{0,4}}?{_NUMBERED_UNIT_KIND}"
    r")\b",
    re.IGNORECASE | re.UNICODE,
)


# Після назви стоїть шифр або посилання «цієї самої …» — частину вже названо.
_UNIT_ALREADY_CLOSED_AFTER_RE = re.compile(
    r"^\s*(?:(?:військов\w+\s+частин\w*|в\s*/?\s*ч)\s+[АA]?\d+"
    r"|(?:ціє|цьо|тіє|то|цій|цим|цією|тим|тією)\w*\s+сам\w*)",
    re.IGNORECASE,
)


def find_unmatched_open_unit_spans(text: str) -> list[tuple[int, int]]:
    """Повертає діапазони відкритих назв частин, які лишилися після шифрування.

    Не підсвічує лінійні внутрішні батальйони/дивізіони, які вже належать закритій в/ч,
    і територіальні центри комплектування (ТЦК лишається відкритим, розд. 9.5.6).
    """
    if not text:
        return []
    spans = []
    for match in _UNMATCHED_OPEN_UNIT_RE.finditer(text):
        start, end = match.start(), match.end()
        following_text = text[end:end + 60]
        if _UNIT_ALREADY_CLOSED_AFTER_RE.match(following_text):
            continue
        if re.match(r"^\s*комплектуванн", following_text, re.IGNORECASE):
            continue
        spans.append((start, end + _quoted_name_length(text, end)))
    return spans


# Родові означення після виду частини належать до назви: «батальйону резерву»,
# «вузла зв'язку», «центру підготовки». Жовта позначка на них закінчується, а
# для заготовки в таблиці їх треба забрати разом із назвою.
_UNIT_NAME_TAIL_WORD_RE = re.compile(
    r"[ \t]+(?!(?:військов\w*|частин\w*|цієї|цього|того|тієї|та|і|й|до|з|із|зі|на|у|в|для|від)\b)"
    r"[а-яіїєґʼ'’-]+(?:у|ю|и|і|ї|ння|ня|ів|ей)(?=[\s,.;:)»”]|$)",
    re.IGNORECASE,  # «КУДИ» пишеться ВЕЛИКИМИ — хвіст там такий самий
)


# Номер і вид частини — щоб відрізнити справді нову частину від ЗГАДКИ вже
# наявної, яка в наказі написана не повністю («55 ОКРЕМОГО ПОЛКУ» при рядку
# «55 окремий полк радіотехнічного забезпечення»). Шифрувати таку згадку не
# можна (пошук лише за номером заборонений, 4.2.7), але й дописувати її в
# таблицю дублем — теж.
_UNIT_KIND_KEYS = (
    ("бригада", r"бригад"), ("полк", r"полк"), ("батальйон", r"батальйон"),
    ("дивізіон", r"дивізіон"), ("загін", r"заг[іо]н"), ("центр", r"центр"),
    ("госпіталь", r"госпітал"), ("база", r"\bбаз[аиіуо]"), ("вузол", r"вуз(?:ол|л)"),
    ("арсенал", r"арсенал"), ("корпус", r"корпус"), ("дивізія", r"дивізі[яїює]"),
    ("комендатура", r"комендатур"),
)


def _unit_number_and_kind(name: str) -> tuple[str, str] | None:
    text = str(name or "")
    number = re.match(r"\s*(\d{1,4})(?!\d)", text)
    if not number:
        return None
    found = []
    for key, pattern in _UNIT_KIND_KEYS:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            found.append((match.start(), key))
    return (number.group(1), min(found)[1]) if found else None


def _similar_table_row(name: str, mapping: dict) -> str:
    """Стовпець A рядка з тим самим номером і видом частини, або порожньо."""
    wanted = _unit_number_and_kind(name)
    if not wanted:
        return ""
    for key, value in (mapping or {}).items():
        column_a = str(value.get("open_name") or key) if isinstance(value, dict) else str(key)
        if _unit_number_and_kind(column_a) == wanted:
            return column_a
    return ""


# Назва в лапках після виду частини («командування «Тест»») — частина назви,
# тож жовта позначка має захоплювати і її.
_QUOTED_NAME_AFTER_RE = re.compile(r"[ \t]*[«“„\"][^«»“”„\"]{0,40}[»”\"]")


def _quoted_name_length(text: str, end: int) -> int:
    match = _QUOTED_NAME_AFTER_RE.match(text, end)
    return match.end() - end if match else 0


def _unit_name_tail_length(text: str, end: int, max_words: int = 3) -> int:
    position = end
    for _ in range(max_words):
        match = _UNIT_NAME_TAIL_WORD_RE.match(text, position)
        if not match:
            break
        position = match.end()
    return position - end


def unmatched_open_unit_spans(text: str, mapping=None) -> list[tuple[int, int]]:
    """`find_unmatched_open_unit_spans` без назв рядків таблиці з позначкою «$».

    Такі назви лишаються відкритими навмисно, тож жовтої позначки не отримують.
    """
    spans = find_unmatched_open_unit_spans(text)
    kept_names = keep_open_names_of(mapping)
    if not spans or not kept_names:
        return spans
    from nodeautomationtoolkit.builtin_nodes.message_order import soften_unit_text

    kept_patterns = [_build_unit_fuzzy_pattern(soften_unit_text(name)) for name in kept_names]
    result = []
    for start, end in spans:
        span_text = soften_unit_text(text[start:end]).casefold()
        window = soften_unit_text(text[max(0, start - 200):end + 200])
        if any(
            span_text in match.group(0).casefold()
            for pattern in kept_patterns
            for match in pattern.finditer(window)
        ):
            continue
        result.append((start, end))
    return result


def collect_new_unit_names(text: str, mapping: dict, similar: list | None = None) -> list[str]:
    """Відкриті назви частин із наказу, яких немає в таблиці (стовпець A).

    Шукається так само, як жовта позначка в повідомленні: текст шифрується, і
    все, що лишилося відкритим, — нові частини. Рядок таблиці з порожнім
    шифром сюди не потрапляє: частина в таблиці вже є, бракує лише шифру.
    """
    ciphered, _, _ = cipher_unit_names(text or "", mapping)
    # Той самий пом'якшувальний фільтр, що й у шифруванні: апострофи, дефіси.
    from nodeautomationtoolkit.builtin_nodes.message_order import soften_unit_text

    known_patterns = [
        _build_unit_fuzzy_pattern(soften_unit_text(str(name)))
        for name in (*(mapping or {}), *keep_open_names_of(mapping))
        if str(name).strip()
    ]
    names: list[str] = []
    seen: set[str] = set()
    for start, end in unmatched_open_unit_spans(ciphered, mapping):
        tail = _unit_name_tail_length(ciphered, end)
        name = re.sub(r"\s+", " ", ciphered[start:end + tail]).strip()
        key = name.casefold()
        if not name or key in seen:
            continue
        seen.add(key)
        softened_name = soften_unit_text(name)
        if any(pattern.search(softened_name) for pattern in known_patterns):
            continue
        # Частина з таким номером і видом у таблиці вже є — це неповна згадка,
        # а не нова частина: у `similar` (для журналу), але не в заготовки.
        close_row = _similar_table_row(name, mapping)
        if close_row:
            if similar is not None:
                similar.append((name, close_row))
            continue
        names.append(name)
    return names


def describe_cipher_problems(problems: list) -> list[str]:
    """Рядки журналу про те, чого бракує в таблиці для закритого змісту."""
    lines = []
    for problem in problems or []:
        if problem[0] == "no_cipher":
            lines.append(
                f"УВАГА: у таблиці порожній шифр (стовпець B) для «{problem[1]}» — "
                "у повідомленні назва лишається відкритою."
            )
        elif problem[0] == "corps_missing":
            lines.append(
                f"УВАГА: для «{problem[1]}» у стовпці D стоїть «{problem[2]}», але рядка "
                "цього корпусу з шифром у таблиці немає — ланку корпусу не додано."
            )
        elif problem[0] == "merged_cell":
            lines.append(
                f"УВАГА: у стовпці A «{problem[1]}» записано дві частини, а шифр у рядку той самий, "
                f"що й у «{problem[2]}» — першу частину в повідомленні не зашифровано. "
                "Запишіть її окремим рядком зі своїм шифром."
            )
    return lines


def _table_open_name_column(sheet) -> int:
    """Номер стовпця відкритої назви (A) — так само, як його визначає читання таблиці."""
    for row in sheet.iter_rows(min_row=1, max_row=10):
        for cell in row:
            key = str(cell.value or "").strip().casefold()
            if "відкрит" in key or ("назва" in key and "закрит" not in key and "скороч" not in key):
                return cell.column
    return 1


def _append_rows_with_excel(source, names: list[str], column: int) -> bool:
    """Дописує рядки-заготовки самим Excel і повертає, чи вдалося.

    openpyxl зберігає формули, але НЕ їхні обчислені значення, а словник
    читається саме за значеннями (розд. 2.2) — після такого запису шифри
    з формул читались би порожніми. Excel під час збереження перераховує
    книгу, тож результати формул лишаються на місці.
    """
    try:
        import win32com.client
    except Exception:
        return False

    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(Path(source).resolve()))
        try:
            sheet = book.ActiveSheet
            used = sheet.UsedRange
            last_row = used.Row + used.Rows.Count - 1
            for offset, name in enumerate(names, start=1):
                row = last_row + offset
                sheet.Cells(row, column).Value = name
                for shift in (0, 1):
                    sheet.Cells(row, column + shift).Interior.Color = 0x00FFFF  # жовтий (BGR)
            book.Save()
        finally:
            book.Close(SaveChanges=False)
        return True
    except Exception:
        return False
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def append_unit_stubs_to_table(table_path: str, names: list[str]) -> dict:
    """Дописує в словник рядки-заготовки для нових частин (розд. 9.5.7).

    Стовпець A — назва так, як її знайдено в наказі (відмінок і повну назву
    виправляє користувач), решта порожня. Такі рядки `read_recipient_mapping`
    пропускає, доки не заповнено шифр, тож маршрутизації вони не зачіпають.

    `.csv` і `.xlsx` без формул дописує openpyxl. Таблицю З ФОРМУЛАМИ дописує
    сам Excel (`_append_rows_with_excel`): openpyxl не зберігає обчислених
    значень, а словник читається саме за ними (розд. 2.2). Якщо Excel
    недоступний, заготовки йдуть в окремий файл поруч. Перед будь-яким записом
    у таблицю робиться резервна копія.

    Повертає `{"added": [...], "path": ..., "backup": ..., "separate": bool}`.
    """
    source = Path(table_path)
    result = {"added": [], "path": str(source), "backup": "", "separate": False}
    wanted: list[str] = []
    for name in names or []:
        clean = re.sub(r"\s+", " ", str(name or "")).strip()
        if clean and clean.casefold() not in {item.casefold() for item in wanted}:
            wanted.append(clean)
    if not wanted or not source.is_file():
        return result

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")

    def make_backup() -> None:
        backup_path = source.with_name(f"{source.stem}.backup-{stamp}{source.suffix}")
        shutil.copy2(source, backup_path)
        result["backup"] = str(backup_path)

    def busy(path: Path) -> PermissionError:
        return PermissionError(
            f"Файл «{path.name}» відкритий в іншій програмі (найімовірніше в Excel). "
            "Закрийте його, щоб дописати нові частини."
        )

    suffix = source.suffix.casefold()
    if suffix == ".csv":
        import csv

        raw = source.read_text(encoding="utf-8-sig", errors="replace")
        delimiter = ";" if raw.count(";") >= raw.count(",") else ","
        existing = {
            re.sub(r"\s+", " ", row[0]).strip().casefold()
            for row in csv.reader(raw.splitlines(), delimiter=delimiter)
            if row
        }
        added = [name for name in wanted if name.casefold() not in existing]
        if added:
            make_backup()
            try:
                with source.open("a", encoding="utf-8", newline="") as handle:
                    if raw and not raw.endswith(("\n", "\r")):
                        handle.write("\r\n")
                    writer = csv.writer(handle, delimiter=delimiter)
                    for name in added:
                        writer.writerow([name, "", "", "", "", ""])
            except PermissionError:
                raise busy(source)
        result["added"] = added
        return result

    if suffix == ".xlsx":
        from openpyxl.styles import PatternFill

        workbook = openpyxl.load_workbook(source)
        has_formulas = any(
            cell.data_type == "f"
            for sheet in workbook.worksheets
            for row in sheet.iter_rows()
            for cell in row
        )
        sheet = workbook.active
        column = _table_open_name_column(sheet)
        # Наявні назви беремо ЗА ЗНАЧЕННЯМИ: назва може бути результатом формули.
        values_sheet = openpyxl.load_workbook(source, data_only=True).active
        existing = {
            re.sub(r"\s+", " ", str(cell.value)).strip().casefold()
            for (cell,) in values_sheet.iter_rows(min_col=column, max_col=column)
            if cell.value is not None
        }
        added = [name for name in wanted if name.casefold() not in existing]
        if not added:
            result["added"] = []
            return result

        make_backup()
        if has_formulas:
            # Таблицю з формулами дописує сам Excel — інакше зникнуть обчислені
            # значення. Якщо Excel недоступний, нижче спрацює окремий файл.
            if _append_rows_with_excel(source, added, column):
                result["added"] = added
                return result
        else:
            last_row = max(
                (cell.row for row in sheet.iter_rows() for cell in row if cell.value not in (None, "")),
                default=0,
            )
            fill = PatternFill(fill_type="solid", start_color="FFFF00", end_color="FFFF00")
            for offset, name in enumerate(added, start=1):
                sheet.cell(row=last_row + offset, column=column, value=name).fill = fill
                sheet.cell(row=last_row + offset, column=column + 1).fill = fill
            try:
                workbook.save(source)
            except PermissionError:
                raise busy(source)
            result["added"] = added
            return result

    # Excel недоступний або формат не той — окремий файл поруч.
    separate = source.with_name(f"{source.stem} — нові частини.xlsx")
    if separate.is_file():
        stub_book = openpyxl.load_workbook(separate)
        stub_sheet = stub_book.active
    else:
        stub_book = openpyxl.Workbook()
        stub_sheet = stub_book.active
        stub_sheet.append(
            ["Відкрита назва (A)", "Шифр (B)", "Скорочення (C)", "Корпус (D)", "Кому (E)", "Куди (F)"]
        )
    existing = {
        re.sub(r"\s+", " ", str(row[0] or "")).strip().casefold()
        for row in stub_sheet.iter_rows(min_row=2, values_only=True)
        if row
    }
    added = [name for name in wanted if name.casefold() not in existing]
    for name in added:
        stub_sheet.append([name, "", "", "", "", ""])
    if added:
        try:
            stub_book.save(separate)
        except PermissionError:
            raise busy(separate)
    result.update(added=added, path=str(separate), separate=True)
    return result
