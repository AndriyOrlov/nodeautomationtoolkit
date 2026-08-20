"""Видалення порожньої останньої сторінки примірника.

Деякі накази мають порожню останню сторінку — у примірнику її бути не повинно.
Сторінку з текстом (зокрема службову «останню сторінку» заготовки) чіпати не можна.
"""

from nodeautomationtoolkit.builtin_nodes.copy_generator import remove_trailing_empty_page


class _Range:
    def __init__(self, doc, text: str, index: int | None = None):
        self._doc = doc
        self.Text = text
        self._index = index
        self.Start = 0
        self.End = len(text)

    def Delete(self):
        self._doc.delete_paragraph(self._index)


class _Paragraph:
    def __init__(self, doc, text: str, index: int):
        self.Range = _Range(doc, text, index)


class _Paragraphs:
    def __init__(self, doc):
        self._doc = doc

    @property
    def Count(self):
        return len(self._doc.texts)

    def __call__(self, index):
        return _Paragraph(self._doc, self._doc.texts[index - 1], index)


class _FakeDoc:
    """Мінімальна імітація документа Word: абзаци + сторінки.

    Сторінка вважається порожньою, якщо всі абзаци після `last_page_start_index`
    не містять тексту.
    """

    def __init__(self, texts, pages, last_page_start_index):
        self.texts = list(texts)
        self._pages = pages
        self._last_page_start_index = last_page_start_index
        self.Paragraphs = _Paragraphs(self)
        self.repaginated = 0

    # --- Word API ---
    def Repaginate(self):
        self.repaginated += 1

    def ComputeStatistics(self, kind):
        assert kind == 2  # wdStatisticPages
        return self._pages

    def GoTo(self, what, which, count):
        return _Range(self, "", None)

    @property
    def Content(self):
        return _Range(self, "".join(self.texts), None)

    def Range(self, start, end):
        tail = "".join(self.texts[self._last_page_start_index:])
        return _Range(self, tail, None)

    # --- поведінка імітації ---
    def delete_paragraph(self, index):
        del self.texts[index - 1]
        if len(self.texts) <= self._last_page_start_index:
            self._pages -= 1


def test_single_page_document_is_untouched():
    doc = _FakeDoc(["Текст\r"], pages=1, last_page_start_index=0)
    assert remove_trailing_empty_page(doc) is False
    assert doc.texts == ["Текст\r"]


def test_last_page_with_text_is_kept():
    """Службова «остання сторінка» заготовки має лишитись."""
    doc = _FakeDoc(
        ["Пункт\r", "Розрахунок розсилки\r"], pages=2, last_page_start_index=1
    )
    assert remove_trailing_empty_page(doc) is False
    assert len(doc.texts) == 2


def test_empty_last_page_is_removed():
    doc = _FakeDoc(["Пункт\r", "\r", "\r"], pages=2, last_page_start_index=1)

    assert remove_trailing_empty_page(doc) is True
    assert doc.texts == ["Пункт\r"]


def test_stops_before_deleting_meaningful_text():
    doc = _FakeDoc(["Підписант\r", "\r"], pages=2, last_page_start_index=1)

    remove_trailing_empty_page(doc)

    assert "Підписант\r" in doc.texts


def test_whitespace_only_page_counts_as_empty():
    doc = _FakeDoc(["Пункт\r", "   \r"], pages=2, last_page_start_index=1)
    assert remove_trailing_empty_page(doc) is True
