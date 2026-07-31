from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class WordParagraph:
    index: int
    text: str
    style: str
    is_empty: bool = False

    def __str__(self) -> str:
        return self.text


@dataclass(frozen=True, slots=True)
class WordParagraphs:
    items: tuple[WordParagraph, ...] = ()
    source_path: str = ""

    def __iter__(self):
        return iter(self.items)

    def __len__(self) -> int:
        return len(self.items)

    def __bool__(self) -> bool:
        return bool(self.items)

    def text(self, separator: str = "\n") -> str:
        return separator.join(item.text for item in self.items)


@dataclass(frozen=True, slots=True)
class WordDocument:
    path: str
    file_name: str
    paragraphs: WordParagraphs
    text: str

    @property
    def paragraph_count(self) -> int:
        return len(self.paragraphs)

    def __str__(self) -> str:
        return f"{self.file_name}: {self.paragraph_count} абзаців"


@dataclass(frozen=True, slots=True)
class WordSaveResult:
    path: str
    paragraph_count: int
    message: str = "Документ збережено"

    @property
    def file_name(self) -> str:
        return Path(self.path).name

    def __str__(self) -> str:
        return f"{self.message}: {self.path} ({self.paragraph_count} абзаців)"
