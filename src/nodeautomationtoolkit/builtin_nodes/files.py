from __future__ import annotations

from pathlib import Path

from nodeautomationtoolkit.core.definition import node


@node(
    name="Шлях до файла",
    category="Файли",
    description="Передає локальний шлях до файла без читання його вмісту.",
    type_id="builtin.files.file_path",
)
def file_path(path: str = "") -> str:
    return path


@node(
    name="Назва файла",
    category="Файли",
    description="Повертає назву файла з розширенням.",
    type_id="builtin.files.filename",
)
def filename(path: str) -> str:
    return Path(path).name


@node(
    name="Назва без розширення",
    category="Файли",
    type_id="builtin.files.stem",
)
def stem(path: str) -> str:
    return Path(path).stem


@node(
    name="Список файлів",
    category="Файли",
    description="Знаходить файли за маскою в локальній папці.",
    type_id="builtin.files.list_files",
)
def list_files(folder: str, pattern: str = "*") -> list[str]:
    return [str(item) for item in sorted(Path(folder).glob(pattern)) if item.is_file()]

