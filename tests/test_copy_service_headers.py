"""Колонтитули примірника: номер сторінки та гриф обмеження доступу.

Розкладка, узгоджена з користувачем:

* верхній колонтитул — ТРИ абзаци: номер сторінки (12 пт), гриф (14 пт) і
  порожній абзац, щоб текст наказу не зливався з колонтитулом;
* нижній колонтитул — лише гриф (14 пт);
* усе по центру, Times New Roman;
* на ПЕРШІЙ і на ОСТАННІЙ сторінці немає ні номера, ні грифа;
* нумерація наскрізна, рахується з першої сторінки (друга сторінка — «2»).

У заготовці цього немає нічого — усе будує код.
"""

import pytest

from nodeautomationtoolkit.builtin_nodes.copy_generator import (
    SERVICE_MARK_TEXT,
    apply_service_headers,
    isolate_last_page_section,
)

_WD_ALIGN_CENTER = 1
_WD_FIELD_PAGE = 33
_WD_HEADER_FOOTER_PRIMARY = 1
_WD_HEADER_FOOTER_FIRST_PAGE = 2
_WD_SECTION_BREAK_NEXT_PAGE = 2


class _Font:
    # Перелік закритий навмисно — див. коментар до `_PageSetup`.
    __slots__ = ("Name", "Size", "Bold", "Italic", "Underline", "Color")

    def __init__(self):
        self.Name = None
        self.Size = None
        self.Bold = None
        self.Italic = None
        self.Underline = None
        # Червоний — щоб було видно, що колір ЗАДАЄТЬСЯ, а не успадковується
        # від стилю заготовки. Саме таким колонтитул і виходив у користувача.
        self.Color = 255


class _ParagraphFormat:
    __slots__ = (
        "Alignment",
        "LeftIndent",
        "RightIndent",
        "FirstLineIndent",
        "SpaceBefore",
        "SpaceAfter",
        "LineSpacingRule",
        "PageBreakBefore",
    )

    def __init__(self):
        self.Alignment = None
        self.LeftIndent = None
        self.RightIndent = None
        self.FirstLineIndent = None
        self.SpaceBefore = None
        self.SpaceAfter = None
        self.LineSpacingRule = None
        self.PageBreakBefore = False


class _StoryParagraph:
    """Абзац колонтитула: тримає власний шрифт і геометрію."""

    def __init__(self, story, index):
        self._story = story
        self._index = index

    @property
    def Range(self):
        return _StoryParagraphRange(self._story, self._index)

    @property
    def Format(self):
        return self._story.formats[self._index]


class _StoryParagraphRange:
    def __init__(self, story, index):
        self._story = story
        self._index = index

    @property
    def Font(self):
        return self._story.fonts[self._index]

    @property
    def ParagraphFormat(self):
        return self._story.formats[self._index]

    def Collapse(self, direction):
        self._story.collapsed = True

    def Delete(self):
        self._story.drop(self._index)


class _StoryParagraphs:
    def __init__(self, story):
        self._story = story

    @property
    def Count(self):
        return len(self._story.lines)

    def __call__(self, index):
        return _StoryParagraph(self._story, index - 1)


class _StoryRange:
    def __init__(self, story):
        self._story = story

    @property
    def Text(self):
        return "\r".join(self._story.lines)

    @Text.setter
    def Text(self, value):
        self._story.set_text(value)

    @property
    def Paragraphs(self):
        return _StoryParagraphs(self._story)

    def InsertParagraphAfter(self):
        self._story.append_line("")


class _PageNumbers:
    __slots__ = ("RestartNumberingAtSection", "StartingNumber")

    def __init__(self):
        self.RestartNumberingAtSection = None
        self.StartingNumber = None


class _Story:
    """Колонтитул (Header або Footer) одного розділу."""

    def __init__(self, kind):
        self.kind = kind
        self.lines = [""]
        self.fonts = [_Font()]
        self.formats = [_ParagraphFormat()]
        self.fields = []
        self.LinkToPrevious = True
        self.PageNumbers = _PageNumbers()
        self.collapsed = False

    # --- Word API ---
    @property
    def Range(self):
        return _StoryRange(self)

    # --- поведінка імітації ---
    def set_text(self, value):
        self.lines = (value or "").split("\r")
        self.fonts = [_Font() for _ in self.lines]
        self.formats = [_ParagraphFormat() for _ in self.lines]
        self.fields = []

    def append_line(self, value):
        self.lines.append(value)
        self.fonts.append(_Font())
        self.formats.append(_ParagraphFormat())

    def drop(self, index):
        del self.lines[index]
        del self.fonts[index]
        del self.formats[index]

    def add_field(self, kind):
        self.fields.append(kind)

    # --- зручні перевірки ---
    @property
    def visible_lines(self):
        return [line for line in self.lines if line]


class _Collection:
    """Headers/Footers: індексуються кодом розкладки (1 — основна, 2 — перша)."""

    def __init__(self, kind):
        self._items = {
            _WD_HEADER_FOOTER_PRIMARY: _Story(f"{kind}-primary"),
            _WD_HEADER_FOOTER_FIRST_PAGE: _Story(f"{kind}-first"),
        }

    def __call__(self, index):
        return self._items[index]

    def __iter__(self):
        return iter(self._items.values())


class _PageSetup:
    """`PageSetup` із ЗАКРИТИМ переліком властивостей.

    Справжній Word на невідому властивість кидає «Property ... can not be set»,
    і саме так генерація впала на кожному наказі: у Word немає
    `DifferentOddAndEvenPagesHeaderFooter`, властивість зветься
    `OddAndEvenPagesHeaderFooter`. Вільний фейк цього не ловив, тому тут
    `__slots__` — щоб описка в назві провалювала тест, а не прогін користувача.
    """

    __slots__ = ("DifferentFirstPageHeaderFooter", "OddAndEvenPagesHeaderFooter")

    def __init__(self):
        self.DifferentFirstPageHeaderFooter = False
        self.OddAndEvenPagesHeaderFooter = False


class _SectionRange:
    def __init__(self, section):
        self._section = section

    @property
    def Start(self):
        return self._section.start

    @property
    def Paragraphs(self):
        return _StoryParagraphs(self._section.first_paragraph_story)


class _Section:
    def __init__(self, start):
        self.start = start
        self.PageSetup = _PageSetup()
        self.Headers = _Collection("header")
        self.Footers = _Collection("footer")
        # Окремий носій для першого абзацу розділу — потрібен лише заради
        # прапорця «з нової сторінки».
        self.first_paragraph_story = _Story("body")

    @property
    def Range(self):
        return _SectionRange(self)


class _Sections:
    def __init__(self, doc):
        self._doc = doc

    @property
    def Count(self):
        return len(self._doc.sections)

    def __call__(self, index):
        return self._doc.sections[index - 1]

    def __iter__(self):
        return iter(self._doc.sections)


class _Fields:
    def __init__(self, doc):
        self._doc = doc

    def Add(self, range_obj, kind, text, preserve):
        self._doc.last_field_story = range_obj._story
        range_obj._story.add_field(kind)


class _DocRange:
    def __init__(self, doc, start, end):
        self._doc = doc
        self.Start = start
        self.End = end

    @property
    def Text(self):
        return self._doc.body[self.Start:self.End]

    def Delete(self):
        self._doc.body = self._doc.body[: self.Start] + self._doc.body[self.End:]
        self._doc.deleted.append((self.Start, self.End))

    def InsertBreak(self, kind):
        self._doc.breaks.append((self.Start, kind))
        self._doc.sections.append(_Section(self.Start))


class _FakeDoc:
    """Мінімальна імітація документа Word: сторінки, розділи, колонтитули."""

    def __init__(self, pages=3, body="A\fB\fC", last_page_start=4, sections=1):
        self._pages = pages
        self.body = body
        self._last_page_start = last_page_start
        self.sections = [_Section(0) for _ in range(sections)]
        if sections > 1:
            self.sections[-1].start = last_page_start
        self.Fields = _Fields(self)
        self.breaks = []
        self.deleted = []
        self.repaginated = 0
        self.last_field_story = None

    # --- Word API ---
    @property
    def Sections(self):
        return _Sections(self)

    def Repaginate(self):
        self.repaginated += 1

    def ComputeStatistics(self, kind):
        assert kind == 2  # wdStatisticPages
        return self._pages

    def GoTo(self, what, which, count):
        return _DocRange(self, self._last_page_start, self._last_page_start)

    def Range(self, start, end):
        return _DocRange(self, start, end)


def _two_section_doc():
    """Документ, у якому остання сторінка вже є окремим розділом."""
    return _FakeDoc(pages=3, last_page_start=4, sections=2)


# --------------------------------------------------------------------------
# Розкладка верхнього колонтитула
# --------------------------------------------------------------------------


def test_header_has_number_mark_and_empty_paragraph():
    doc = _two_section_doc()
    assert apply_service_headers(doc) is True

    header = doc.sections[0].Headers(_WD_HEADER_FOOTER_PRIMARY)
    assert header.lines == ["", SERVICE_MARK_TEXT, ""], (
        "верхній колонтитул: 1-й абзац — номер, 2-й — гриф, 3-й — порожній"
    )


def test_page_number_field_goes_into_the_first_paragraph():
    doc = _two_section_doc()
    apply_service_headers(doc)

    header = doc.sections[0].Headers(_WD_HEADER_FOOTER_PRIMARY)
    assert header.fields == [_WD_FIELD_PAGE]
    assert doc.last_field_story is header


def test_font_sizes_are_twelve_for_number_and_fourteen_for_the_mark():
    doc = _two_section_doc()
    apply_service_headers(doc)

    header = doc.sections[0].Headers(_WD_HEADER_FOOTER_PRIMARY)
    assert header.fonts[0].Size == 12.0
    assert header.fonts[1].Size == 14.0
    assert all(font.Name == "Times New Roman" for font in header.fonts)


def test_everything_in_the_header_is_centred():
    doc = _two_section_doc()
    apply_service_headers(doc)

    header = doc.sections[0].Headers(_WD_HEADER_FOOTER_PRIMARY)
    assert [fmt.Alignment for fmt in header.formats] == [_WD_ALIGN_CENTER] * 3
    assert all(fmt.FirstLineIndent == 0 for fmt in header.formats)


# --------------------------------------------------------------------------
# Нижній колонтитул
# --------------------------------------------------------------------------


def test_footer_holds_only_the_mark():
    doc = _two_section_doc()
    apply_service_headers(doc)

    footer = doc.sections[0].Footers(_WD_HEADER_FOOTER_PRIMARY)
    assert footer.lines == [SERVICE_MARK_TEXT]
    assert footer.fonts[0].Size == 14.0
    assert footer.formats[0].Alignment == _WD_ALIGN_CENTER


def test_page_number_never_appears_in_the_footer():
    doc = _two_section_doc()
    apply_service_headers(doc)

    footer = doc.sections[0].Footers(_WD_HEADER_FOOTER_PRIMARY)
    assert footer.fields == []


# --------------------------------------------------------------------------
# Чиста перша та чиста остання сторінка
# --------------------------------------------------------------------------


def test_first_page_has_neither_number_nor_mark():
    doc = _two_section_doc()
    apply_service_headers(doc)

    section = doc.sections[0]
    assert section.PageSetup.DifferentFirstPageHeaderFooter is True
    assert section.Headers(_WD_HEADER_FOOTER_FIRST_PAGE).visible_lines == []
    assert section.Footers(_WD_HEADER_FOOTER_FIRST_PAGE).visible_lines == []


def test_last_page_has_neither_number_nor_mark():
    doc = _two_section_doc()
    apply_service_headers(doc)

    last = doc.sections[-1]
    assert last.Headers(_WD_HEADER_FOOTER_PRIMARY).visible_lines == []
    assert last.Footers(_WD_HEADER_FOOTER_PRIMARY).visible_lines == []


def test_last_section_is_unlinked_from_the_previous_one():
    """Інакше очищення останньої сторінки стерло б колонтитули всього тексту."""
    doc = _two_section_doc()
    apply_service_headers(doc)

    last = doc.sections[-1]
    assert last.Headers(_WD_HEADER_FOOTER_PRIMARY).LinkToPrevious is False
    assert last.Footers(_WD_HEADER_FOOTER_PRIMARY).LinkToPrevious is False


def test_single_page_copy_gets_no_headers_at_all():
    """Одна сторінка є водночас першою й останньою."""
    doc = _FakeDoc(pages=1, body="A", last_page_start=0, sections=1)
    assert apply_service_headers(doc) is False

    header = doc.sections[0].Headers(_WD_HEADER_FOOTER_PRIMARY)
    assert header.visible_lines == []
    assert header.fields == []


def test_two_page_copy_gets_no_visible_header_either():
    """Обидві сторінки — перша й остання, тож показувати нічого."""
    doc = _FakeDoc(pages=2, body="A\fB", last_page_start=2, sections=2)
    apply_service_headers(doc)

    assert doc.sections[0].Headers(_WD_HEADER_FOOTER_FIRST_PAGE).visible_lines == []
    assert doc.sections[-1].Headers(_WD_HEADER_FOOTER_PRIMARY).visible_lines == []


# --------------------------------------------------------------------------
# Нумерація
# --------------------------------------------------------------------------


def test_numbering_starts_at_one_in_the_first_section():
    doc = _two_section_doc()
    apply_service_headers(doc)

    numbers = doc.sections[0].Headers(_WD_HEADER_FOOTER_PRIMARY).PageNumbers
    assert numbers.StartingNumber == 1
    assert numbers.RestartNumberingAtSection is True


def test_last_section_continues_the_numbering():
    """Розділ під останню сторінку не має починати відлік заново."""
    doc = _two_section_doc()
    apply_service_headers(doc)

    numbers = doc.sections[-1].Headers(_WD_HEADER_FOOTER_PRIMARY).PageNumbers
    assert numbers.RestartNumberingAtSection is False
    assert numbers.StartingNumber is None


# --------------------------------------------------------------------------
# Виділення останньої сторінки в окремий розділ
# --------------------------------------------------------------------------


def test_last_page_is_moved_into_its_own_section():
    doc = _FakeDoc(pages=3, body="A\fB\fC", last_page_start=4, sections=1)
    assert isolate_last_page_section(doc) is True
    assert doc.Sections.Count == 2
    assert doc.breaks and doc.breaks[0][1] == _WD_SECTION_BREAK_NEXT_PAGE


def test_manual_page_break_is_replaced_not_doubled():
    """Два розриви підряд дали б зайву порожню сторінку."""
    doc = _FakeDoc(pages=3, body="A\fB\fC", last_page_start=4, sections=1)
    isolate_last_page_section(doc)

    assert doc.deleted == [(3, 4)], "ручний розрив сторінки має бути видалений"
    assert doc.breaks[0][0] == 3, "розрив розділу стає на його місце"


def test_natural_page_start_keeps_the_text_intact():
    """Якщо сторінка почалась переливанням тексту, видаляти нічого не можна."""
    doc = _FakeDoc(pages=3, body="ABCDEF", last_page_start=4, sections=1)
    isolate_last_page_section(doc)

    assert doc.deleted == []
    assert doc.body == "ABCDEF"
    assert doc.breaks[0][0] == 4


def test_page_break_before_is_cleared_on_the_new_section():
    """Розрив розділу вже переносить на нову сторінку — другий не потрібен."""
    doc = _FakeDoc(pages=3, body="A\fB\fC", last_page_start=4, sections=1)
    isolate_last_page_section(doc)

    new_section = doc.sections[-1]
    assert new_section.Range.Paragraphs(1).Format.PageBreakBefore is False


def test_existing_last_page_section_is_reused():
    """Заготовка може вже мати окремий розділ — другий розрив зайвий."""
    doc = _two_section_doc()
    assert isolate_last_page_section(doc) is True
    assert doc.breaks == []
    assert doc.Sections.Count == 2


def test_single_page_document_is_not_split():
    doc = _FakeDoc(pages=1, body="A", last_page_start=0, sections=1)
    assert isolate_last_page_section(doc) is False
    assert doc.breaks == []


# --------------------------------------------------------------------------
# Текст грифа
# --------------------------------------------------------------------------


def test_mark_wording_is_koristuvannya():
    """Саме «КОРИСТУВАННЯ» — узгоджено з користувачем і зі зразком 12.3."""
    assert SERVICE_MARK_TEXT == "ДЛЯ СЛУЖБОВОГО КОРИСТУВАННЯ"


@pytest.mark.parametrize("pages", [3, 5, 12])
def test_middle_pages_always_get_the_full_header(pages):
    doc = _FakeDoc(pages=pages, body="A\fB\fC", last_page_start=4, sections=2)
    assert apply_service_headers(doc) is True

    header = doc.sections[0].Headers(_WD_HEADER_FOOTER_PRIMARY)
    assert header.lines == ["", SERVICE_MARK_TEXT, ""]
    assert header.fields == [_WD_FIELD_PAGE]


def test_odd_and_even_headers_are_switched_off_by_the_real_property_name():
    """Регресія: у Word немає `DifferentOddAndEvenPagesHeaderFooter`.

    Описка в назві валила генерацію КОЖНОГО наказу помилкою
    «Property ... can not be set», і жоден примірник не створювався.
    """
    doc = _two_section_doc()
    apply_service_headers(doc)

    assert doc.sections[0].PageSetup.OddAndEvenPagesHeaderFooter is False


def test_header_text_is_black_not_inherited_from_the_template():
    """Регресія: колонтитул виходив ЧЕРВОНИМ.

    Колір — самостійна властивість шрифту: ні `Name`, ні `Size` його не
    перекривають, тож без явного присвоєння діяв колір стилю `Header`
    заготовки (правило 5.2.1, поширене на колонтитули).
    """
    doc = _two_section_doc()
    apply_service_headers(doc)

    header = doc.sections[0].Headers(_WD_HEADER_FOOTER_PRIMARY)
    footer = doc.sections[0].Footers(_WD_HEADER_FOOTER_PRIMARY)

    assert [font.Color for font in header.fonts] == [0, 0, 0]
    assert [font.Color for font in footer.fonts] == [0]
