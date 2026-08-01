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


@node(
    name="Створити файл",
    category="Файли",
    description=(
        "Створює локальний текстовий файл із заданим вмістом. Існуючий файл "
        "не перезаписується без окремого дозволу."
    ),
    type_id="builtin.files.create",
    outputs={"path": "str", "name": "str", "message": "str"},
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def create_file(
    path: str = "",
    content: str = "",
    encoding: str = "utf-8",
    overwrite: bool = False,
) -> dict:
    target = Path(path).expanduser()
    if not path.strip() or not target.name:
        raise ValueError("Не вказано шлях нового файла")
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and not overwrite:
        raise FileExistsError(f"Файл уже існує: {target}")
    target.write_text(content, encoding=encoding)
    resolved = target.resolve()
    return {
        "path": str(resolved),
        "name": resolved.name,
        "message": f"Створено: {resolved.name}",
    }


@node(
    name="Перейменувати файл",
    category="Файли",
    description=(
        "Перейменовує локальний файл. Нова назва без папки залишає файл у тій "
        "самій папці; можна передати й повний новий шлях."
    ),
    type_id="builtin.files.rename",
    outputs={"path": "str", "old_name": "str", "new_name": "str", "message": "str"},
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def rename_file(path: str = "", new_name_or_path: str = "", overwrite: bool = False) -> dict:
    source = Path(path).expanduser()
    if not source.is_file():
        raise FileNotFoundError(f"Файл не знайдено: {path or '(шлях порожній)'}")
    if not new_name_or_path.strip():
        raise ValueError("Не вказано нову назву")
    requested = Path(new_name_or_path).expanduser()
    target = requested if requested.parent != Path(".") else source.with_name(requested.name)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and not overwrite:
        raise FileExistsError(f"Файл уже існує: {target}")
    if target.exists():
        target.unlink()
    old_name = source.name
    renamed = source.rename(target).resolve()
    return {
        "path": str(renamed),
        "old_name": old_name,
        "new_name": renamed.name,
        "message": f"Перейменовано: {old_name} → {renamed.name}",
    }
