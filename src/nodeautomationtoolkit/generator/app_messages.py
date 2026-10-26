"""Повідомлення за наказом.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import os
import re
import shutil
from tkinter import filedialog, messagebox
from tkinter.constants import DISABLED, NORMAL
import win32com.client
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    read_recipient_mapping,
    map_military_units,
)
from nodeautomationtoolkit.builtin_nodes.message_order import (
    generate_decision_order,
    cipher_unit_names,
    clear_table_caches,
    find_content_start_line,
    reflow_soft_breaks,
)
from nodeautomationtoolkit.builtin_nodes.blank_images import remove_blank_images
from nodeautomationtoolkit.builtin_nodes.template_tags import (
    expand_common_tags, signer_tags,
)

from .errors import UserError, explain_error
from .messages import MESSAGE_BY_CALCULATION_TEXT, _MESSAGE_ADDRESSEE_KIND_TAGS, build_addressee_kind_text, build_message_recipient_groups, message_item_contexts, message_skipped_item_lines
from .naming import _slash_to_lines, extract_metadata_from_filename, format_message_date
from .order_text import text_before_order_signer
from .units import append_unit_stubs_to_table, collect_new_unit_names, describe_cipher_problems, unmatched_open_unit_spans
from .word_doc import _apply_carried_formatting, _read_source_formatting, copy_template_for_editing, force_quit_word, is_biographical_paragraph, is_page_break_only_text, is_path_writable, paragraph_line_count, read_document_text, remove_page_breaks_in_range
from nodeautomationtoolkit.builtin_nodes.typography import (
    apply_ukrainian_typography,
    clean_duplicated_units,
    ensure_blank_line_before_items,
)


class MessagesMixin:
    """Повідомлення за наказом."""

    def select_message_cover_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.message_cover_template_path.set(path)
            self.save_config()

    def select_message_content_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.message_content_template_path.set(path)
            self.save_config()

    def select_message_calc_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.doc")])
        if path:
            self.message_calc_template_path.set(path)
            self.save_config()

    def select_message_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.message_out_folder.set(path)
            self.save_config()

    def show_message_tags(self):
        messagebox.showinfo(
            "Теги повідомлень",
            "Спільні для обох шаблонів:\n"
            "{{номер_наказу}} / {{номер}}, {{дата_наказу}} / {{дата}}, "
            "{{кому_список}}, {{куди}}, {{виконавець}}.\n"
            "Також діють спільні теги підписанта й засвідчувача/затверджувача, "
            "як у витягах і примірниках.\n\n"
            "Розмістіть {{кому_список}} у кожному рядку таблиці окремо: один тег "
            "отримує одного унікального адресата. Якщо рядків замало, програма "
            "додає перед наступним блоком копії останнього рядка таблиці — штамп "
            "зсувається вниз і не перекривається. Порядок адресатів: корпуси, "
            "частини, лише обласні ТЦК.\n\n"
            "{{тцк чі вч}} — тип адресата, визначається автоматично:\n"
            "• лише військові частини → «командирам військових частин»;\n"
            "• лише ТЦК → «начальникам ОТЦК та СП»;\n"
            "• і те, і те → «командирам військових частин та начальникам ОТЦК та СП».\n\n"
            "Розрахунок розсилки (необов'язковий третій зразок): таблиця з "
            "{{кому_список}} у рядку, як у супроводі; рядки нумеруються 1., 2., … "
            "Якщо його вибрано, у супроводі в {{кому_список}} пишеться "
            "«Згідно з розрахунком розсилки», а зайві заготовлені рядки прибираються.\n\n"
            "У шаблоні зі змістом:\n"
            "{{зміст_шифр}} — пункти наказу із заміною відкритих назв частин "
            "на шифри з Excel. Назви, які лишилися без збігу, виділяються жовтим.\n\n"
            "{{виконавець}} підставляється рівно так, як введено у формі, і "
            "розміщується за 3 см від низу сторінки, якщо тег поза таблицею. "
            "Службові рядки («ВІДКРИТА ІНФОРМАЦІЯ» тощо) програма НЕ додає — "
            "їх треба прописати у самому зразку; текст шаблону після "
            "{{виконавець}} зберігається без змін.\n\n"
            "Номер наказу береться з назви файла, дата — теж, у вигляді "
            "20.05.2025 року (без розкриття словами)."
        )

    @staticmethod
    def _replace_message_tags(document, replacements: dict[str, str]) -> list[tuple[int, int]]:
        """Замінює теги й повертає діапазони завершальних блоків виконавця."""
        executor_blocks = []
        for tag, value in expand_common_tags(replacements).items():
            find_obj = document.Content.Find
            find_obj.Text = tag
            iterations = 0
            while find_obj.Execute() and iterations < 100:
                iterations += 1
                replacement_start = find_obj.Parent.Start
                # Виконавець підставляється рівно так, як його ввели.
                # Рядки «ВІДКРИТА ІНФОРМАЦІЯ» та «(Обмежено в розповсюдженні…)»
                # програма НЕ додає: вони прописані у самому зразку окремо.
                replacement = value
                find_obj.Parent.Text = replacement
                if tag == "{{виконавець}}":
                    # Текст шаблону після виконавця більше не видаляється:
                    # там лежать службові рядки зразка, які мають зберігатися.
                    executor_blocks.append((replacement_start, replacement_start + len(replacement)))
                find_obj = document.Content.Find
                find_obj.Text = tag
        return executor_blocks

    @staticmethod
    def _replace_executor_in_headers(document, value: str) -> int:
        """Підставляє `{{виконавець}}` у КОЛОНТИТУЛАХ і повертає кількість замін.

        Колонтитул — окрема «історія» Word: `document.Content` бачить лише
        основну частину документа, тому тег, поставлений у колонтитул зразка,
        лишався незаповненим і друкувався як є.

        Позицію тут не підбираємо: колонтитул і так унизу сторінки, а блок
        виконавця в тілі документа працює як раніше — його закладку й
        вирівнювання ця функція не чіпає.
        """
        tag = "{{виконавець}}"
        replaced = 0
        try:
            sections = list(document.Sections)
        except TypeError:
            sections = [document.Sections(index) for index in range(1, document.Sections.Count + 1)]
        for section in sections:
            for collection in (section.Headers, section.Footers):
                try:
                    stories = list(collection)
                except TypeError:
                    # Первинний, першої сторінки та парних сторінок.
                    stories = [collection(index) for index in (1, 2, 3)]
                for header_footer in stories:
                    for _ in range(20):  # запобіжник від вічного циклу
                        try:
                            find_obj = header_footer.Range.Find
                            find_obj.Text = tag
                            if not find_obj.Execute():
                                break
                            find_obj.Parent.Text = value
                        except Exception:
                            break
                        replaced += 1
        return replaced

    @staticmethod
    def _position_message_executor_at_page_bottom(document, bookmark_name: str):
        """Тримає завершальний блок повідомлення за 3 см віднизу сторінки."""
        if not document.Bookmarks.Exists(bookmark_name):
            return
        block_range = document.Bookmarks(bookmark_name).Range
        executor_start = block_range.Start
        executor_index = block_last_index = None
        for index in range(1, document.Paragraphs.Count + 1):
            paragraph_range = document.Paragraphs(index).Range
            if paragraph_range.Start <= executor_start < paragraph_range.End:
                executor_index = index
            block_end = max(block_range.Start, block_range.End - 1)
            if paragraph_range.Start <= block_end < paragraph_range.End:
                block_last_index = index
            if executor_index is not None and block_last_index is not None:
                break
        if executor_index is None or block_last_index is None:
            return

        try:
            # У таблиці завершальний блок не рухаємо: там відступи контролює шаблон.
            if document.Paragraphs(executor_index).Range.Information(12):  # wdWithInTable
                return
            for index in range(executor_index, block_last_index + 1):
                paragraph_format = document.Paragraphs(index).Range.ParagraphFormat
                paragraph_format.KeepTogether = True
                paragraph_format.KeepWithNext = index < block_last_index
            while executor_index > 1:
                previous = document.Paragraphs(executor_index - 1).Range
                if previous.Text.replace("\r", "").replace("\x07", "").strip():
                    break
                previous.Delete()
                executor_index -= 1
                block_last_index -= 1

            document.Repaginate()
            page_number = document.Paragraphs(executor_index).Range.Information(3)  # wdActiveEndPageNumber
            target_vertical = document.PageSetup.PageHeight - document.Application.CentimetersToPoints(3)
            if document.Paragraphs(block_last_index).Range.Information(3) != page_number:
                return
            # Оптимізоване наближення замість 120 покрокових Repaginate
            block_vertical = document.Paragraphs(block_last_index).Range.Information(6)  # wdVerticalPositionRelativeToPage
            vertical_diff = target_vertical - block_vertical
            if vertical_diff > 25:
                batch_enters = max(1, int(vertical_diff / 18))
                pos = document.Paragraphs(executor_index).Range.Start
                document.Range(pos, pos).InsertBefore("\r" * batch_enters)
                executor_index += batch_enters
                block_last_index += batch_enters
                document.Repaginate()
                # Перевірка: якщо пакетна вставка виштовхнула блок на нову сторінку — відкат
                batch_page = document.Paragraphs(block_last_index).Range.Information(3)
                if batch_page != page_number:
                    for _ in range(batch_enters):
                        document.Paragraphs(executor_index - 1).Range.Delete()
                        executor_index -= 1
                        block_last_index -= 1
                    document.Repaginate()

            for _ in range(15):
                block_vertical = document.Paragraphs(block_last_index).Range.Information(6)  # wdVerticalPositionRelativeToPage
                if block_vertical >= target_vertical:
                    break
                position = document.Paragraphs(executor_index).Range.Start
                document.Range(position, position).InsertBefore("\r")
                executor_index += 1
                block_last_index += 1
                document.Repaginate()
                block_page = document.Paragraphs(block_last_index).Range.Information(3)
                block_vertical = document.Paragraphs(block_last_index).Range.Information(6)
                if block_page != page_number or block_vertical > target_vertical:
                    document.Paragraphs(executor_index - 1).Range.Delete()
                    document.Repaginate()
                    break
        except Exception:
            # Верстка виконавця не має блокувати створення повідомлення.
            return

    @staticmethod
    def _find_recipient_table_cell(document, position: int):
        """Повертає таблицю, рядок і комірку, що містять задану позицію Word.

        Використовує table.Range.Cells для безпечного обходу таблиць
        з вертикально об'єднаними комірками без збоїв COM.
        """
        try:
            for table_index in range(1, document.Tables.Count + 1):
                table = document.Tables(table_index)
                for cell_index in range(1, table.Range.Cells.Count + 1):
                    cell = table.Range.Cells(cell_index)
                    cell_range = cell.Range
                    if cell_range.Start <= position < cell_range.End:
                        return table, cell.RowIndex, cell.ColumnIndex
        except Exception:
            pass
        return None

    @staticmethod
    def _fill_recipient_slots(
        document, recipients: list[str], drop_unused_rows: bool = False
    ) -> tuple[int, int]:
        """Заповнює кожен {{кому_список}} одним наступним адресатом.

        Теги зазвичай лежать у рядках таблиці поруч зі штампом. Тому список
        не вставляється в одну комірку: жоден адресат не повторюється. Якщо
        готових рядків бракує, нижче останнього рядка-адресата копіюється його
        форматований рядок. Нові рядки вставляються до наступного блока, тож
        штамп зсувається вниз, а не перекривається текстом.
        """
        slot_count = 0
        last_row_context = None
        unused_rows = []
        find_obj = document.Content.Find
        find_obj.Text = "{{кому_список}}"
        iterations = 0
        while find_obj.Execute() and iterations < 500:
            iterations += 1
            found_range = find_obj.Parent.Duplicate
            row_context = MessagesMixin._find_recipient_table_cell(document, found_range.Start)
            last_row_context = row_context or last_row_context
            value = recipients[slot_count] if slot_count < len(recipients) else ""
            find_obj.Parent.Text = value
            if drop_unused_rows and not value and row_context:
                unused_rows.append(row_context)
            slot_count += 1
            find_obj = document.Content.Find
            find_obj.Text = "{{кому_список}}"

        # Заготовлені рядки, яким не дісталось адресата, прибираються, якщо в
        # них більше нічого немає (супровід «згідно з розрахунком розсилки»).
        for table, row_index, _ in sorted(
            unused_rows, key=lambda context: (context[0].Range.Start, context[1]), reverse=True
        ):
            try:
                row = table.Rows(row_index)
                if not row.Range.Text.replace("\r", "").replace("\x07", "").strip():
                    row.Delete()
            except Exception:
                continue

        if slot_count >= len(recipients) or last_row_context is None:
            return slot_count, max(0, len(recipients) - slot_count)

        table, row_index, column_index = last_row_context
        remaining = recipients[slot_count:]
        try:
            source_row = table.Rows(row_index)
            after_address_rows = table.Rows.Count > row_index
            anchor_row = table.Rows(row_index + 1) if after_address_rows else None
            source_value = recipients[slot_count - 1] if slot_count else ""
            for recipient in remaining:
                new_row = table.Rows.Add(anchor_row) if anchor_row else table.Rows.Add()
                source_cells = source_row.Cells.Count
                for cell_index in range(1, min(source_cells, new_row.Cells.Count) + 1):
                    source_range = source_row.Cells(cell_index).Range.Duplicate
                    source_range.End -= 1  # Службовий знак кінця комірки Word.
                    target_range = new_row.Cells(cell_index).Range.Duplicate
                    target_range.End -= 1
                    target_range.FormattedText = source_range.FormattedText

                target_cell = new_row.Cells(column_index)
                target_find = target_cell.Range.Find
                target_find.Text = source_value
                if target_find.Execute():
                    target_find.Parent.Text = recipient
                else:
                    # Резервний варіант для порожнього або нестандартно
                    # відформатованого рядка таблиці.
                    cell_range = target_cell.Range.Duplicate
                    cell_range.End -= 1
                    cell_range.Text = recipient
                source_row = new_row
                source_value = recipient
                slot_count += 1
        except Exception:
            # Якщо Word не дає продублювати конкретний рядок, повертаємо
            # невміщені адресати як попередження, не вставляючи їх поверх штампа.
            return slot_count, len(recipients) - slot_count
        return slot_count, 0

    @staticmethod
    def _renumber_recipient_rows(document, recipients: list[str]) -> int:
        """Номери «1.», «2.» … у рядках таблиці з адресатами.

        `_fill_recipient_slots` копіює рядок разом з номером, тож у всіх рядках
        стояло б «1.». Номер міняється лише в першій комірці рядка, де стоїть
        число («1», «1.»); решта тексту й форматування лишаються. Автонумерацію
        Word (номера в тексті комірки немає) не чіпаємо — Word рахує її сам.
        """
        def norm(value: str) -> str:
            return " ".join(str(value or "").replace("\x07", " ").split())

        wanted = {norm(recipient) for recipient in recipients if norm(recipient)}
        changed = 0
        for table in document.Tables:
            counter = 0
            try:
                for row in table.Rows:
                    cells = row.Cells
                    texts = [norm(cells(index).Range.Text) for index in range(1, cells.Count + 1)]
                    if not any(text in wanted for text in texts[1:]):
                        continue
                    counter += 1
                    number = re.fullmatch(r"(\d{1,3})\.?", texts[0] if texts else "")
                    if not number or number.group(1) == str(counter):
                        continue
                    find_obj = cells(1).Range.Find
                    find_obj.Text = number.group(1)
                    if find_obj.Execute():
                        find_obj.Parent.Text = str(counter)
                        changed += 1
            except Exception:
                # Таблиця з вертикально об'єднаними комірками не дає пройти
                # рядки — номери в ній лишаються як є.
                continue
        return changed

    def _copy_order_content(self, doc, tag_range, content_source: dict) -> int:
        """Переносить зміст наказу у повідомлення РАЗОМ ІЗ ФОРМАТУВАННЯМ.

        Копіюється `FormattedText` кожного абзацу — точно так само, як у
        витягах, тому стилі, вирівнювання та відступи лишаються такими ж, як
        в оригіналі наказу. Шифрування виконується поабзацно вже ПІСЛЯ
        копіювання, завдяки чому форматування абзацу зберігається (вставка
        простим текстом його втрачала й «кривила» верстку).

        Повертає кількість підсвічених невпізнаних назв частин.
        """
        source_doc = content_source["doc"]
        first_para = int(content_source["first_para"])
        last_para = int(content_source["last_para"])
        mapping = content_source.get("mapping") or {}

        # Валідація ДО будь-яких змін документа: якщо діапазон некоректний,
        # викликач ще може безпечно перейти на запасний спосіб вставки.
        total = source_doc.Paragraphs.Count
        if not 1 <= first_para <= last_para <= total:
            raise ValueError(
                f"некоректний діапазон абзаців наказу {first_para}-{last_para} із {total}"
            )

        insert_point = tag_range.Paragraphs(1).Range.Start
        tag_range.Paragraphs(1).Range.Delete()
        content_start = insert_point
        # Формат кожного абзацу читаємо одразу, а пишемо ПІСЛЯ копіювання —
        # суцільними відрізками з однаковим форматом: у наказі більшість
        # пунктів оформлені однаково, і замість ~10 звернень на абзац
        # виходить кілька на весь зміст. Пізніші вставки йдуть ПІСЛЯ вже
        # записаних меж, тож ті не зсуваються.
        carried: list[tuple[int, int, dict, dict]] = []
        try:
            skip_paragraphs = content_source.get("skip_paragraphs") or set()
            # `Paragraphs(i)` щоразу проходить документ із початку — цикл був
            # квадратичним. Беремо перший абзац один раз і йдемо `Next()`.
            next_paragraph = source_doc.Paragraphs(first_para)
            span_range = source_doc.Range(
                next_paragraph.Range.Start, source_doc.Paragraphs(last_para).Range.End
            )
            if span_range.Tables.Count == 0:
                # Без таблиць зміст копіюється суцільними блоками — одне
                # `FormattedText` на блок замість одного на абзац. З таблицями
                # лишається поабзацний шлях нижче: блок відтворив би таблицю,
                # а поабзацний — ні, і результат змінився б.
                insert_point = self._copy_order_blocks(
                    doc, insert_point, next_paragraph, first_para, last_para, skip_paragraphs
                )
                next_paragraph = None
            for p_index in range(first_para, last_para + 1):
                if next_paragraph is None:
                    break
                source_paragraph = next_paragraph
                if p_index < last_para:
                    next_paragraph = source_paragraph.Next()
                if p_index in skip_paragraphs:
                    continue
                source_range = source_paragraph.Range.Duplicate
                source_text = source_range.Text or ""
                if is_page_break_only_text(source_text):
                    continue  # абзац-розрив сторінки з наказу не переносимо
                destination = doc.Range(insert_point, insert_point)
                destination.FormattedText = source_range.FormattedText
                paragraph_start, insert_point = destination.Start, destination.End
                font_values, geometry = _read_source_formatting(source_range)
                if "\x0c" in source_text:
                    # Сам розрив прибираємо, текст абзацу лишається.
                    insert_point = remove_page_breaks_in_range(doc, paragraph_start, insert_point)
                if carried and carried[-1][1] == paragraph_start and carried[-1][2:] == (font_values, geometry):
                    carried[-1] = (carried[-1][0], insert_point, font_values, geometry)
                else:
                    carried.append((paragraph_start, insert_point, font_values, geometry))
            for run_start, run_end, font_values, geometry in carried:
                _apply_carried_formatting(doc, run_start, run_end, font_values, geometry)
        except Exception as error:
            # Частина абзаців уже вставлена — повертатись до простого тексту
            # не можна, інакше зміст задвоївся б.
            raise RuntimeError(f"збій копіювання абзацу наказу: {error}") from error

        # Шифрування змінює довжину абзаців, тому межі змісту тримаємо
        # закладкою — після правок вона вкаже актуальний діапазон.
        # Порожні білі зображення наказу в повідомлення не переносимо.
        pasted_range = doc.Range(content_start, insert_point)
        removed_images = remove_blank_images(doc, pasted_range)
        if removed_images:
            insert_point = pasted_range.End
            self.log(f"Прибрано порожніх білих зображень: {removed_images}.")

        bookmark_name = "nat_message_content"
        try:
            doc.Bookmarks.Add(bookmark_name, doc.Range(content_start, insert_point))
        except Exception:
            bookmark_name = ""

        # У примірниках текст наказу лишається 1-в-1 (без шифрів), тому
        # шифрування вмикається лише там, де воно справді потрібне.
        highlighted = 0
        if content_source.get("cipher", True):
            highlighted = self._cipher_inserted_content(doc, content_start, insert_point, mapping)

        if bookmark_name and doc.Bookmarks.Exists(bookmark_name):
            content_range = doc.Bookmarks(bookmark_name).Range
            try:
                self._apply_message_layout_rules(doc, content_range.Start, content_range.End)
            except Exception as layout_error:
                self.log(f"УВАГА: не вдалося застосувати правила верстки змісту: {layout_error}")
            try:
                doc.Bookmarks(bookmark_name).Delete()
            except Exception:
                pass
        return highlighted

    @staticmethod
    def _copy_order_blocks(doc, insert_point, first_paragraph, first_para, last_para, skip_paragraphs) -> int:
        """Копіює абзаци наказу суцільними блоками; повертає новий кінець.

        Блок обривається на пропущеному абзаці (внутрішнє переміщення) й на
        абзаці, де є лише розрив сторінки, — так само, як у поабзацному шляху.
        Формат кожного абзацу переноситься явно (`_carry_source_formatting`):
        якщо в блоці він однаковий — одним записом на блок, інакше поабзацно,
        за зсувом від початку блоку (FormattedText копіює текст 1-в-1).
        """
        blocks: list[list[tuple[int, int, object, bool]]] = []
        current: list[tuple[int, int, object, bool]] = []
        paragraph = first_paragraph
        for p_index in range(first_para, last_para + 1):
            source_paragraph = paragraph
            if p_index < last_para:
                paragraph = source_paragraph.Next()
            if p_index in skip_paragraphs:
                if current:
                    blocks.append(current)
                current = []
                continue
            source_range = source_paragraph.Range
            source_text = source_range.Text or ""
            if is_page_break_only_text(source_text):
                if current:
                    blocks.append(current)
                current = []
                continue  # абзац-розрив сторінки з наказу не переносимо
            current.append((source_range.Start, source_range.End, source_range, "\x0c" in source_text))
        if current:
            blocks.append(current)

        source_doc = first_paragraph.Range.Document
        for block in blocks:
            source_start, source_end = block[0][0], block[-1][1]
            source_block = source_doc.Range(source_start, source_end)
            destination = doc.Range(insert_point, insert_point)
            destination.FormattedText = source_block.FormattedText
            block_start, insert_point = destination.Start, destination.End

            block_fonts, block_geometry = _read_source_formatting(source_block)
            uniform = (
                {"Name", "Size"} <= set(block_fonts)
                and len(block_geometry) == 4
                and all(value != 9999999 for value in block_geometry.values())  # wdUndefined
            )
            if uniform:
                _apply_carried_formatting(doc, block_start, insert_point, block_fonts, block_geometry)
            elif insert_point - block_start == source_end - source_start:
                runs: list[tuple[int, int, dict, dict]] = []
                for start, end, source_range, _has_break in block:
                    font_values, geometry = _read_source_formatting(source_range)
                    run_start = block_start + start - source_start
                    run_end = block_start + end - source_start
                    if runs and runs[-1][1] == run_start and runs[-1][2:] == (font_values, geometry):
                        runs[-1] = (runs[-1][0], run_end, font_values, geometry)
                    else:
                        runs.append((run_start, run_end, font_values, geometry))
                for run_start, run_end, font_values, geometry in runs:
                    _apply_carried_formatting(doc, run_start, run_end, font_values, geometry)
            else:
                # Довжина не збіглася — межі абзаців за зсувом не визначити.
                # Парою «абзац наказу ↔ абзац результату» за порядком.
                pasted = [item.Range for item in doc.Range(block_start, insert_point).Paragraphs]
                for (start, end, source_range, _has_break), target in zip(block, pasted):
                    font_values, geometry = _read_source_formatting(source_range)
                    _apply_carried_formatting(doc, target.Start, target.End, font_values, geometry)

            if any(has_break for *_rest, has_break in block):
                # Сам розрив прибираємо, текст абзацу лишається.
                insert_point = remove_page_breaks_in_range(doc, block_start, insert_point)
        return insert_point

    def _apply_message_layout_rules(self, doc, content_start: int, content_end: int) -> None:
        """Нерозривність блоків і заповнення сторінки — як у витягах.

        1. Пункт разом зі своїм біографічним блоком (р.н., освіта, ІПН/РНОКПП,
           ВОС, «Підлягає направленню…») лишається неподільним: він не може
           розриватись між сторінками.
        2. Шапки (§, «Відповідно до …:», «У ЗАПАС ЗА ПІДПУНКТОМ …:») зчеплені
           з наступним пунктом.
        3. Міжрядковий інтервал добирається в діапазоні 16 → 14 пт так, щоб
           зміст займав якнайменше сторінок і не лишав напівпорожніх.
        """
        content_range = doc.Range(content_start, content_end)
        # Скидаємо KeepWithNext, який міг прийти разом із FormattedText.
        # Далі він виставляється лише за нашою структурою блоків.
        content_range.ParagraphFormat.KeepWithNext = False
        spans = []
        list_numbers = []
        texts = []
        for paragraph in content_range.Paragraphs:
            paragraph_range = paragraph.Range
            spans.append((paragraph_range.Start, paragraph_range.End))
            texts.append(paragraph_range.Text)
            try:
                list_numbers.append(str(paragraph_range.ListFormat.ListString or ""))
            except Exception:
                list_numbers.append("")

        def paragraph_kind(text: str, list_number: str = "") -> str:
            clean = (text or "").strip()
            if not clean:
                return "blank"
            if clean.startswith("§"):
                return "heading"
            # У конвертованих DOCX номер часто є автонумерацією Word і не
            # входить у Range.Text. Без ListString усі такі пункти ставали
            # «продовженням» і злипались в один завеликий блок.
            if (
                re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s", clean)
                or re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]?$", list_number.strip())
            ):
                return "item"
            if clean.endswith(":"):
                return "heading"
            return "continuation"

        kinds = [paragraph_kind(text, list_number) for text, list_number in zip(texts, list_numbers)]

        # KeepTogether (непорожні абзаци) і KeepWithNext рахуємо для кожного
        # абзацу, а пишемо суцільними відрізками однакових значень — одне
        # звернення до Word на відрізок замість кількох на кожен абзац.
        keep_flags = [False] * len(spans)

        # Останній абзац кожної групи «пункт + біографія» не тягне наступний.
        for index in range(len(spans)):
            if kinds[index] == "blank":
                continue

            next_meaningful = next(
                (j for j in range(index + 1, len(kinds)) if kinds[j] != "blank"), None
            )
            if next_meaningful is None:
                keep_with_next = False
            elif kinds[index] == "heading":
                keep_with_next = True
            else:
                # Пункт тримає свій біографічний блок; новий пункт або шапка
                # починають окрему групу.
                keep_with_next = kinds[next_meaningful] == "continuation"
            keep_flags[index] = keep_with_next

            # Порожній абзац після шапки не повинен розривати її зв'язок із
            # першим пунктом. Порожній абзац після завершеного пункту, навпаки,
            # лишається дозволеною точкою розриву між двома пунктами.
            gap_end = next_meaningful if next_meaningful is not None else len(kinds)
            for blank_index in range(index + 1, gap_end):
                keep_flags[blank_index] = keep_with_next

        # KeepWithNext скинуто на False вище, тож пишемо лише відрізки з True.
        def write_true_runs(flags: list[bool], prop: str) -> None:
            run_start = None
            for index, flag in enumerate(flags + [False]):
                if flag and run_start is None:
                    run_start = index
                elif not flag and run_start is not None:
                    paragraph_format = doc.Range(spans[run_start][0], spans[index - 1][1]).ParagraphFormat
                    setattr(paragraph_format, prop, True)
                    run_start = None

        write_true_runs([kind != "blank" for kind in kinds], "KeepTogether")
        write_true_runs(keep_flags, "KeepWithNext")

        # Заповнення сторінки: найбільший інтервал із діапазону, що дає
        # найменшу кількість сторінок. Інтервал однаковий для всіх абзаців,
        # тож він ставиться на весь зміст двома зверненнями, а не 2×N.
        spacing_range = doc.Range(spans[0][0], spans[-1][1]) if spans else content_range

        def set_spacing(points: float) -> None:
            fmt = spacing_range.ParagraphFormat
            fmt.LineSpacingRule = 4  # wdLineSpaceExactly
            fmt.LineSpacing = points

        best_points = 16.0
        best_pages = None
        points = 16.0
        while points >= 14.0 - 1e-6:
            set_spacing(points)
            doc.Repaginate()
            pages = doc.ComputeStatistics(2)  # wdStatisticPages
            if best_pages is None or pages < best_pages:
                best_pages, best_points = pages, points
            points -= 0.5
        set_spacing(best_points)
        doc.Repaginate()

    @staticmethod
    def _cipher_inserted_content(doc, content_start: int, content_end: int, mapping: dict) -> int:
        """Шифрує вставлені абзаци поабзацно, зберігаючи їх форматування."""
        content_range = doc.Range(content_start, content_end)
        # Один прохід колекцією: `Paragraphs(i)` у циклі квадратичний, а текст
        # читаємо там само, без другого діапазону на кожен абзац.
        spans = []
        for paragraph in content_range.Paragraphs:
            paragraph_range = paragraph.Range
            spans.append((paragraph_range.Start, paragraph_range.Text or ""))

        contexts = message_item_contexts([raw.rstrip("\r\x07") for _, raw in spans])
        highlighted = 0
        # Йдемо з кінця: правки не зсувають позиції попередніх абзаців.
        for (start, raw), context_before in reversed(list(zip(spans, contexts))):
            core = raw.rstrip("\r\x07")
            if not core.strip():
                continue

            ciphered, _, _ = cipher_unit_names(core, mapping, context_before=context_before)
            ciphered = apply_ukrainian_typography(clean_duplicated_units(ciphered))
            # Мʼякий перенос посеред фрази зшивається САМЕ ТУТ — у повідомленні
            # довга відкрита назва стає трьома словами, і рядок, розірваний у
            # наказі, лишався б напівпорожнім. Переноси після закінчення
            # (кома, крапка) — тобто біографічний блок — не чіпаються.
            # Витягів і примірників це не стосується: там текст 1-в-1.
            ciphered = reflow_soft_breaks(ciphered)
            if ciphered != core:
                doc.Range(start, start + len(core)).Text = ciphered

            # Відкриті назви, які лишилися без шифру, підсвічуємо жовтим.
            for span_start, span_end in unmatched_open_unit_spans(ciphered, mapping):
                doc.Range(start + span_start, start + span_end).HighlightColorIndex = 7  # wdYellow
                highlighted += 1
        return highlighted

    def _insert_plain_content(self, doc, tag_range, encrypted_content: str, mapping=None) -> int:
        """Запасний спосіб вставки змісту — простим текстом.

        Використовується ЛИШЕ тоді, коли перенести форматування з наказу не
        вдалося. Форматування тут доводиться вгадувати за вмістом рядка, тому
        воно менш точне, ніж копіювання з оригіналу.
        """
        if not encrypted_content:
            return 0

        content_start = tag_range.Start
        cleaned_content = clean_duplicated_units(encrypted_content)
        separated_content = ensure_blank_line_before_items(cleaned_content)
        formatted_content = apply_ukrainian_typography(separated_content)
        tag_range.Text = formatted_content
        content_range = doc.Range(content_start, content_start + len(formatted_content))

        for pi in range(1, content_range.Paragraphs.Count + 1):
            paragraph = content_range.Paragraphs(pi)
            p_text = paragraph.Range.Text.strip()
            if not p_text:
                continue
            p_format = paragraph.Range.ParagraphFormat
            if is_biographical_paragraph(p_text):
                p_format.Alignment = 1  # wdAlignParagraphCenter
                p_format.LeftIndent = 0
                p_format.RightIndent = 0
                p_format.FirstLineIndent = 0
                p_format.SpaceBefore = 0
                p_format.SpaceAfter = 0
                p_format.LineSpacingRule = 0  # wdLineSpaceSingle
            elif re.match(r"^\d{1,3}[\.\)]", p_text):
                p_format.Alignment = 3  # wdAlignParagraphJustify
                p_format.LeftIndent = 0
                p_format.RightIndent = 0
                p_format.FirstLineIndent = 35.45  # 1.25 см
                p_format.SpaceBefore = 6
                p_format.SpaceAfter = 0
            elif (
                p_text.startswith("Відповідно до")
                or "ЗВІЛЬНИТИ" in p_text.upper()
                or "ПРИЗНАЧИТИ:" in p_text.upper()
            ):
                p_format.Alignment = 3
                p_format.LeftIndent = 0
                p_format.RightIndent = 0
                p_format.FirstLineIndent = 35.45
                p_format.SpaceBefore = 6
                p_format.SpaceAfter = 6
            elif "Призначається на" in p_text or "шпк" in p_text.lower():
                p_format.Alignment = 3
                p_format.LeftIndent = 0
                p_format.RightIndent = 0
                p_format.FirstLineIndent = 35.45
                p_format.SpaceBefore = 0
                p_format.SpaceAfter = 6

        highlighted = 0
        for span_start, span_end in unmatched_open_unit_spans(formatted_content, mapping):
            doc.Range(content_start + span_start, content_start + span_end).HighlightColorIndex = 7
            highlighted += 1
        return highlighted

    def _create_message_file(
        self,
        word,
        template_path: str,
        output_path: str,
        replacements: dict[str, str],
        encrypted_content: str = "",
        recipients: list[str] | None = None,
        content_source: dict | None = None,
        number_recipient_rows: bool = False,
    ) -> tuple[int, int, int]:
        # Якщо результат уже відкритий у Word, зберегти його неможливо.
        # Перевіряємо це ДО всієї роботи й повідомляємо зрозуміло.
        if not is_path_writable(output_path):
            raise UserError(
                f"файл «{os.path.basename(output_path)}» уже відкритий — "
                "найімовірніше у Word, з минулого разу.",
                "Закрийте це вікно Word і натисніть кнопку ще раз. "
                "Поки файл відкритий, програма не може його перезаписати.",
            )

        # Робоча копія шаблону створюється у тимчасовій папці, а не за
        # кінцевим шляхом: так оригінал шаблону лишається недоторканим, а
        # заблокований результат не зриває обробку. Розширення копії
        # відповідає реальному вмісту шаблону (він може бути у форматі
        # Word 97-2003), інакше Word відмовиться її відкрити.
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(output_path)), "_nat_temp")
        os.makedirs(temp_dir, exist_ok=True)
        working_path = copy_template_for_editing(
            template_path,
            os.path.join(temp_dir, os.path.basename(output_path)),
            label="шаблон повідомлення",
        )
        doc = word.Documents.Open(os.path.abspath(working_path), ReadOnly=False)
        try:
            executor_blocks = self._replace_message_tags(doc, replacements)
            if "{{виконавець}}" in replacements:
                header_hits = self._replace_executor_in_headers(
                    doc, replacements["{{виконавець}}"]
                )
                if header_hits:
                    self.log(
                        f"  Виконавця підставлено в колонтитул: {header_hits} раз(и) "
                        f"({os.path.basename(output_path)})."
                    )
            executor_bookmarks = []
            for index, (start, end) in enumerate(executor_blocks):
                bookmark_name = f"nat_message_executor_{index}"
                doc.Bookmarks.Add(bookmark_name, doc.Range(start, end))
                executor_bookmarks.append(bookmark_name)
            recipient_slots = recipient_overflow = 0
            if recipients is not None:
                recipient_slots, recipient_overflow = self._fill_recipient_slots(
                    doc, recipients, drop_unused_rows=recipients == [MESSAGE_BY_CALCULATION_TEXT]
                )
                if number_recipient_rows:
                    self._renumber_recipient_rows(doc, recipients)
            highlighted_count = 0
            if encrypted_content or content_source:
                find_obj = doc.Content.Find
                find_obj.Text = "{{зміст_шифр}}"
                if not find_obj.Execute():
                    self.log("УВАГА: у шаблоні зі змістом не знайдено тег {{зміст_шифр}}.")
                else:
                    # Основний шлях — перенос змісту РАЗОМ ІЗ ФОРМАТУВАННЯМ
                    # наказу, точно як у витягах. Запасний шлях (простий текст)
                    # лишається на випадок, коли діапазон абзаців визначити не вдалося.
                    copied_with_formatting = False
                    if content_source is not None:
                        try:
                            highlighted_count = self._copy_order_content(
                                doc, find_obj.Parent, content_source
                            )
                            copied_with_formatting = True
                        except ValueError as prepare_error:
                            self.log(
                                "УВАГА: зміст не вдалося перенести з форматуванням "
                                f"({prepare_error}); вставляємо простим текстом."
                            )

                    if not copied_with_formatting:
                        highlighted_count = self._insert_plain_content(
                            doc,
                            find_obj.Parent,
                            encrypted_content,
                            (content_source or {}).get("mapping"),
                        )
            for bookmark_name in executor_bookmarks:
                self._position_message_executor_at_page_bottom(doc, bookmark_name)
                if doc.Bookmarks.Exists(bookmark_name):
                    doc.Bookmarks(bookmark_name).Delete()
            doc.SaveAs2(os.path.abspath(output_path), 16)
            return highlighted_count, recipient_slots, recipient_overflow
        finally:
            doc.Close(False)
            shutil.rmtree(temp_dir, ignore_errors=True)

    # Дописувати нові частини в таблицю заготовками (розд. 9.5.7). Увімкнено
    # лише в Qt-оболонці (`QtShellMixin`); Tk-версія тільки попереджає.
    ADD_NEW_UNITS_TO_TABLE = False

    def _report_table_gaps(self, order_text: str, mapping: dict) -> dict:
        """Показує, чого бракує в таблиці для закритого змісту повідомлення.

        Збій запису заготовок генерацію не зупиняє — лише попереджає.
        """
        body = "\n".join(order_text.splitlines()[find_content_start_line(order_text):])
        problems: list = []
        cipher_unit_names(body, mapping, problems=problems)
        for line in describe_cipher_problems(problems):
            self.log(line)

        similar: list = []
        new_units = collect_new_unit_names(body, mapping, similar=similar)
        for name, row in similar:
            self.log(
                f"УВАГА: «{name}» не зашифровано: у таблиці є схожий рядок «{row}», але назва "
                "в наказі збігається з ним не повністю — перевірте написання в наказі або в стовпці A."
            )
        if new_units:
            self.log(
                f"УВАГА: частин немає в таблиці ({len(new_units)}): {'; '.join(new_units)}. "
                "У повідомленні вони лишаються відкритими й виділені жовтим."
            )
            if self.ADD_NEW_UNITS_TO_TABLE:
                try:
                    stubs = append_unit_stubs_to_table(self.excel_path.get(), new_units)
                except Exception as error:
                    self.log(f"УВАГА: не вдалося дописати нові частини в таблицю: {error}")
                else:
                    if stubs["added"] and stubs["separate"]:
                        self.log(
                            f"УВАГА: дописати в таблицю не вдалося (потрібен Excel), тому нові частини "
                            f"({len(stubs['added'])}) записано окремо: {stubs['path']}. "
                            "Перенесіть рядки в таблицю й заповніть шифри."
                        )
                    elif stubs["added"]:
                        self.log(
                            f"Дописано в таблицю заготовок: {len(stubs['added'])} (виділені жовтим, "
                            f"шифр порожній). Резервна копія: {stubs['backup']}. "
                            "Заповніть шифр і за потреби «Кому»/«Куди», потім створіть повідомлення ще раз."
                        )
                    elif stubs["separate"]:
                        # Таблиця з формулами: заготовки лежать в окремому файлі,
                        # а не в таблиці — інакше користувач шукав би їх не там.
                        self.log(
                            f"Заготовки для цих частин уже є в окремому файлі: {stubs['path']}. "
                            "Перенесіть рядки в таблицю й заповніть шифри."
                        )
                    else:
                        self.log("Заготовки для цих частин уже є в таблиці — заповніть у них шифри.")
        return {"problems": len(problems) + len(similar), "new_units": len(new_units)}

    def run_generate_messages(self):
        self.save_config()
        required = (
            self.excel_path.get(),
            self.doc_path.get(),
            self.message_cover_template_path.get(),
            self.message_content_template_path.get(),
        )
        if not all(required):
            messagebox.showwarning(
                "Помилка",
                "Виберіть словник Excel, наказ і два DOCX-шаблони повідомлень.",
            )
            return

        order_path = os.path.abspath(self.doc_path.get())
        if not os.path.isfile(order_path):
            messagebox.showwarning("Помилка", "Файл наказу не знайдено.")
            return

        out_folder = self.message_out_folder.get() or os.path.join(os.path.dirname(order_path), "Messages_Output")
        try:
            os.makedirs(out_folder, exist_ok=True)
        except OSError as error:
            messagebox.showwarning("Помилка папки", f"Не вдалося створити папку результату:\n{error}")
            return
        self.message_out_folder.set(out_folder)

        self.btn_generate_messages.config(state=DISABLED)
        word = None
        source_doc = None
        try:
            # Таблиця читається з файла ОДИН раз; далі всі пошуки — у памʼяті.
            clear_table_caches()
            mapping = read_recipient_mapping(path=self.excel_path.get()).get("mapping", {})

            # Наказ відкриваємо в тому самому екземплярі Word, з якого потім
            # копіюємо зміст із форматуванням, — інакше номери абзаців не
            # відповідали б прочитаному тексту.
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone: не показувати блокуючі діалоги
            source_doc = word.Documents.Open(order_path, ReadOnly=True)

            # З абзаців, як у витягах: `Content.Text` не має номерів автонумерації
            # Word — такий пункт зливався з шапкою, і внутрішнє переміщення в
            # управлінні не розпізнавалось; ще й таблиці зсували рядки ↔ абзаци.
            message_bio_lines: set[int] = set()
            source_text = read_document_text(source_doc, bio_lines=message_bio_lines)
            order_text, order_signer = text_before_order_signer(source_text)
            order_num, order_date = extract_metadata_from_filename(os.path.basename(order_path))
            # У повідомленнях дата не розкривається словами (на відміну від витягів).
            order_date_formatted = format_message_date(order_date)

            table_gaps = self._report_table_gaps(order_text, mapping)
            routes = map_military_units(text=order_text, mapping=mapping, bio_lines=message_bio_lines)
            recipient_groups = build_message_recipient_groups(mapping, routes)
            recipients = (
                recipient_groups["corps"] + recipient_groups["units"] + recipient_groups["tck"]
            )
            addressee_kind = build_addressee_kind_text(recipient_groups)
            destinations = []
            for data in routes.get("unit_paragraphs", {}).values():
                destination = str(data.get("destination_where") or "").strip()
                if destination and destination not in destinations:
                    destinations.append(destination)

            replacements = {}
            if order_num:
                replacements["{{номер_наказу}}"] = f"№{order_num}"
            if order_date_formatted:
                replacements["{{дата_наказу}}"] = order_date_formatted
            if destinations:
                replacements["{{куди}}"] = "\r".join(destinations)
            # Тег типу адресата підставляється ДО {{виконавець}}: обробка
            # виконавця видаляє весь службовий хвіст шаблону після себе.
            if addressee_kind:
                for kind_tag in _MESSAGE_ADDRESSEE_KIND_TAGS:
                    replacements[kind_tag] = addressee_kind
                self.log(f"Тип адресата ({{{{тцк чі вч}}}}): {addressee_kind}.")
            # Підписант наказу — ті самі теги, що у витягах. Реквізити беруться
            # з наказу, а якщо користувач заповнив поля вручну — з полів.
            gui_signer_position = self.order_signer_position.get().strip()
            gui_signer_rank = self.order_signer_rank.get().strip()
            gui_signer_name = self.order_signer_name.get().strip()
            if gui_signer_position or gui_signer_rank or gui_signer_name:
                order_signer = {
                    "position": gui_signer_position.replace(" / ", chr(10)) or order_signer.get("position", ""),
                    "rank": gui_signer_rank or order_signer.get("rank", ""),
                    "name": gui_signer_name or order_signer.get("name", ""),
                }
            signer_position = _slash_to_lines(str(order_signer.get("position", "")).strip())
            signer_rank = str(order_signer.get("rank", "")).strip()
            signer_name = str(order_signer.get("name", "")).strip()
            replacements.update(signer_tags(signer_position, signer_rank, signer_name))

            # Засвідчувач — СПІЛЬНІ поля з витягами (AGENT.md 11.2).
            replacements["{{згідно_з_оригіналом}}"] = "Згідно з оригіналом"
            replacements["{{засвідчення}}"] = "Згідно з оригіналом"
            if self.certifier_position.get().strip():
                certifier_position = _slash_to_lines(self.certifier_position.get().strip())
                replacements["{{засвідчувач_посада}}"] = certifier_position
                replacements["{{згідно_з_оригіналом_посада}}"] = certifier_position
            if self.certifier_rank.get().strip():
                certifier_rank = self.certifier_rank.get().strip()
                replacements["{{засвідчувач_звання}}"] = certifier_rank
                replacements["{{згідно_з_оригіналом_звання}}"] = certifier_rank
            if self.certifier_name.get().strip():
                certifier_name = self.certifier_name.get().strip()
                replacements["{{засвідчувач_піб}}"] = certifier_name
                replacements["{{згідно_з_оригіналом_піб}}"] = certifier_name
                replacements["{{засвідчувач}}"] = certifier_name

            if self.message_executor.get().strip():
                replacements["{{виконавець}}"] = self.message_executor.get().strip()

            # Внутрішнє переміщення в управлінні — не для повідомлень (галочка
            # на вкладці). Витягів до управління це не стосується.
            skip_lines, skipped_labels = set(), []
            if self.message_skip_internal_management.get():
                skip_lines, skipped_labels = message_skipped_item_lines(routes)
            if skipped_labels:
                self.log(
                    "Не додано до повідомлення (внутрішнє переміщення в управлінні): "
                    + ", ".join(skipped_labels)
                )
            content_text = "\n".join(
                line for index, line in enumerate(order_text.splitlines())
                if index not in skip_lines
            )

            decision = generate_decision_order(text=content_text, mapping=mapping, new_header="")
            encrypted_content = decision.get("decision_text", "")

            # Відповідність рядків тексту абзацам Word — так само, як у витягах.
            content_source = None
            try:
                source_line_to_para = []
                for paragraph_index in range(1, source_doc.Paragraphs.Count + 1):
                    raw = source_doc.Paragraphs(paragraph_index).Range.Text
                    source_line_to_para.extend([paragraph_index] * paragraph_line_count(raw))

                order_lines = order_text.splitlines()
                line_map = source_line_to_para[: len(order_lines)]
                body_start_line = find_content_start_line(order_text)
                if line_map and 0 <= body_start_line < len(line_map):
                    # Абзац пропускається, лише якщо ВСІ його рядки пропущені.
                    paragraph_lines: dict[int, list[int]] = {}
                    for line_index, paragraph_index in enumerate(line_map):
                        paragraph_lines.setdefault(paragraph_index, []).append(line_index)
                    content_source = {
                        "doc": source_doc,
                        "first_para": line_map[body_start_line],
                        "last_para": line_map[-1],
                        "mapping": mapping,
                        "skip_paragraphs": {
                            paragraph_index
                            for paragraph_index, indices in paragraph_lines.items()
                            if skip_lines and all(index in skip_lines for index in indices)
                        },
                    }
            except Exception as map_error:
                self.log(
                    f"УВАГА: не вдалося зіставити абзаци наказу ({map_error}); "
                    "зміст буде вставлено простим текстом."
                )

            safe_number = re.sub(r'[\\/:*?"<>|]', "_", order_num) if order_num else ""
            suffix = f"_№{safe_number}" if safe_number else ""
            cover_output = os.path.join(out_folder, f"Повідомлення_супровід{suffix}.docx")
            content_output = os.path.join(out_folder, f"Повідомлення_шифрований_зміст{suffix}.docx")

            # Обрано зразок «Розрахунок розсилки» — адресати йдуть у нього, а в
            # супроводі замість переліку пишеться фраза (вимога 26.10.2026).
            calc_template = self.message_calc_template_path.get().strip()
            if calc_template and not os.path.isfile(calc_template):
                self.log(f"УВАГА: зразок розрахунку розсилки не знайдено: {calc_template}")
                calc_template = ""
            cover_recipients = [MESSAGE_BY_CALCULATION_TEXT] if calc_template else recipients
            _, recipient_slots, recipient_overflow = self._create_message_file(
                word,
                self.message_cover_template_path.get(),
                cover_output,
                replacements,
                recipients=cover_recipients,
            )
            highlights, _, _ = self._create_message_file(
                word,
                self.message_content_template_path.get(),
                content_output,
                replacements,
                encrypted_content,
                content_source=content_source,
            )
            created = [cover_output, content_output]
            if calc_template:
                calc_output = os.path.join(out_folder, f"Повідомлення_розрахунок_розсилки{suffix}.docx")
                _, recipient_slots, recipient_overflow = self._create_message_file(
                    word,
                    calc_template,
                    calc_output,
                    replacements,
                    recipients=recipients,
                    number_recipient_rows=True,
                )
                created.append(calc_output)
            if recipient_overflow:
                self.log(
                    f"УВАГА: у шаблоні {'розрахунку розсилки' if calc_template else 'супроводу'} "
                    f"{recipient_slots} комірок {{кому_список}}, "
                    f"але адресатів {len(recipients)}. Не вмістилося: {recipient_overflow}."
                )
            self.save_config()
            self.log(
                f"Повідомлення створено: {'; '.join(created)}. "
                f"Адресатів: {len(recipients)}; жовтих позначок: {highlights}."
            )
            messagebox.showinfo(
                "Успіх",
                f"Створено файлів: {len(created)}.\n\n"
                f"Заповнено комірок {{кому_список}}: {min(recipient_slots, len(recipients))} з {len(recipients)}\n"
                f"Невпізнаних назв, виділених жовтим: {highlights}\n"
                f"Частин, яких немає в таблиці: {table_gaps['new_units']}\n"
                f"Без шифру або корпусу в таблиці: {table_gaps['problems']}"
                + (
                    f"\nНе додано (внутрішнє переміщення в управлінні): {', '.join(skipped_labels)}"
                    if skipped_labels else ""
                )
                + (f"\nУВАГА: не вмістилося адресатів: {recipient_overflow}" if recipient_overflow else ""),
            )
        except Exception as error:
            explanation = explain_error(error)
            self.log(f"ПОМИЛКА повідомлень:\n{explanation}")
            messagebox.showerror("Повідомлення не створені", explanation)
        finally:
            if source_doc is not None:
                try:
                    source_doc.Close(False)
                except Exception:
                    pass
            if word:
                force_quit_word(word)
            # Таблиця жила в памʼяті лише на час цього запуску.
            clear_table_caches()
            self.btn_generate_messages.config(state=NORMAL)
