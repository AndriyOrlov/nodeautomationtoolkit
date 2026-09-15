"""Регресія: нерозривність не переходить із завершеного пункту на наступний."""

from types import SimpleNamespace

import generate_extracts
from nodeautomationtoolkit.builtin_nodes.copy_generator import apply_keep_together_rules


class _FormatProxy:
    def __init__(self, document, indexes):
        object.__setattr__(self, "document", document)
        object.__setattr__(self, "indexes", indexes)

    def __setattr__(self, name, value):
        for index in self.indexes:
            setattr(self.document.formats[index], name, value)


class _Range:
    def __init__(self, document, start, end):
        self.document = document
        self.Start = start
        self.End = end

    @property
    def _indexes(self):
        return [
            index
            for index, (start, end) in enumerate(self.document.bounds)
            if start >= self.Start and end <= self.End
        ]

    @property
    def Text(self):
        return "".join(self.document.texts[index] for index in self._indexes)

    @property
    def ParagraphFormat(self):
        return _FormatProxy(self.document, self._indexes)

    @property
    def Paragraphs(self):
        return _Paragraphs(self.document, self._indexes)

    @property
    def ListFormat(self):
        index = self._indexes[0]
        return SimpleNamespace(ListString=self.document.list_numbers[index])


class _Paragraph:
    def __init__(self, document, index):
        self.document = document
        self.index = index

    @property
    def Range(self):
        start, end = self.document.bounds[self.index]
        return _Range(self.document, start, end)


class _Paragraphs:
    def __init__(self, document, indexes):
        self.document = document
        self.indexes = indexes

    @property
    def Count(self):
        return len(self.indexes)

    def __call__(self, one_based_index):
        return _Paragraph(self.document, self.indexes[one_based_index - 1])

    def __iter__(self):
        return (_Paragraph(self.document, index) for index in self.indexes)


class _Document:
    def __init__(self, lines):
        self.texts = [line + "\r" for line in lines]
        self.bounds = []
        position = 0
        for text in self.texts:
            self.bounds.append((position, position + len(text)))
            position += len(text)
        # Усе навмисно успадковано як один ланцюг — це і є рідкісний дефект.
        self.formats = [
            SimpleNamespace(KeepTogether=False, KeepWithNext=True, LineSpacingRule=None, LineSpacing=None)
            for _ in lines
        ]
        self.list_numbers = ["" for _ in lines]
        self.end = position

    def Range(self, start, end):
        return _Range(self, start, end)

    def Repaginate(self):
        pass

    def ComputeStatistics(self, _kind):
        return 1


LINES = [
    "§ 2",
    "",
    "Відповідно до пункту ... ПРИЗНАЧИТИ:",
    "",
    "6. Перший пункт розділу",
    "біографічний блок пункту 6",
    "",
    "7. Наступний пункт",
    "біографічний блок пункту 7",
]


def _assert_expected_boundaries(document):
    flags = [value.KeepWithNext for value in document.formats]
    # § → порожній → шапка → порожній → пункт 6 → його біографія.
    assert flags[:5] == [True, True, True, True, True]
    # Після біографії пункту 6 та на порожньому рядку розрив дозволений.
    assert flags[5:7] == [False, False]
    # Пункт 7 тримає лише власну біографію.
    assert flags[7:] == [True, False]


def test_copy_keep_rules_clear_inherited_chain_between_items():
    document = _Document(LINES)
    apply_keep_together_rules(document, 0, document.end, signer_inside=False)
    _assert_expected_boundaries(document)


def test_message_keep_rules_clear_inherited_chain_between_items():
    document = _Document(LINES)
    generate_extracts.App._apply_message_layout_rules(None, document, 0, document.end)
    _assert_expected_boundaries(document)


def test_message_keep_rules_recognize_word_auto_numbering():
    lines = list(LINES)
    lines[4] = "Перший пункт розділу"
    lines[7] = "Наступний пункт"
    document = _Document(lines)
    document.list_numbers[4] = "6."
    document.list_numbers[7] = "7."

    generate_extracts.App._apply_message_layout_rules(None, document, 0, document.end)

    _assert_expected_boundaries(document)
