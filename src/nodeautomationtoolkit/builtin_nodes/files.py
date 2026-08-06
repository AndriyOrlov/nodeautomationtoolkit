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


@node(
    name="Розширене перейменування",
    category="Файли",
    description=(
        "Повнофункціональний ренеймер: додавання префікса/суфікса, видалення N символів, "
        "видалення тексту до/після символу-маркера, заміна за текстом чи regex."
    ),
    type_id="builtin.files.advanced_rename",
    outputs={
        "path": "str",
        "old_name": "str",
        "new_name": "str",
        "summary": "str",
    },
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def advanced_rename_file(
    path: str = "",
    prefix: str = "",
    suffix: str = "",
    trim_first_n: int = 0,
    trim_last_n: int = 0,
    trim_before_symbol: str = "",
    trim_after_symbol: str = "",
    find_text: str = "",
    replace_text: str = "",
    use_regex: bool = False,
    apply_rename_on_disk: bool = True,
    overwrite: bool = False,
) -> dict:
    source = Path(path).expanduser()
    if not source.is_file():
        raise FileNotFoundError(f"Файл не знайдено: {path or '(порожній шлях)'}")

    stem = source.stem
    ext = source.suffix
    old_name = source.name

    new_stem = stem

    if trim_first_n > 0:
        new_stem = new_stem[trim_first_n:]

    if trim_last_n > 0:
        new_stem = new_stem[:-trim_last_n] if trim_last_n < len(new_stem) else ""

    if trim_before_symbol:
        pos = new_stem.find(trim_before_symbol)
        if pos != -1:
            new_stem = new_stem[pos + len(trim_before_symbol):]

    if trim_after_symbol:
        pos = new_stem.find(trim_after_symbol)
        if pos != -1:
            new_stem = new_stem[:pos]

    if find_text:
        if use_regex:
            import re
            new_stem = re.sub(find_text, replace_text, new_stem)
        else:
            new_stem = new_stem.replace(find_text, replace_text)

    new_stem = f"{prefix}{new_stem}{suffix}"
    new_filename = f"{new_stem}{ext}"
    target = source.with_name(new_filename)

    if apply_rename_on_disk and target != source:
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists() and not overwrite:
            raise FileExistsError(f"Файл уже існує: {target}")
        if target.exists():
            target.unlink()
        renamed_file = source.rename(target).resolve()
        final_path = str(renamed_file)
        final_name = renamed_file.name
    else:
        final_path = str(target.resolve())
        final_name = target.name

    return {
        "path": final_path,
        "old_name": old_name,
        "new_name": final_name,
        "summary": f"Перейменовано: '{old_name}' → '{final_name}'",
    }


@node(
    name="Створити папку",
    category="Файли",
    description="Створює нову папку за вказаним шляхом або назвою.",
    type_id="builtin.files.create_folder",
    outputs={"path": "str", "name": "str", "message": "str"},
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def create_folder(folder_path: str = "", folder_name: str = "") -> dict:
    base = Path(folder_path).expanduser() if folder_path.strip() else Path.cwd()
    target = base / folder_name if folder_name.strip() else base
    target.mkdir(parents=True, exist_ok=True)
    resolved = target.resolve()
    return {
        "path": str(resolved),
        "name": resolved.name,
        "message": f"Папку створено: {resolved.name}",
    }


@node(
    name="Перейменувати папку",
    category="Файли",
    description="Перейменовує існуючу папку.",
    type_id="builtin.files.rename_folder",
    outputs={"path": "str", "old_name": "str", "new_name": "str", "message": "str"},
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def rename_folder(folder_path: str = "", new_name: str = "") -> dict:
    source = Path(folder_path).expanduser()
    if not source.is_dir():
        raise NotADirectoryError(f"Папку не знайдено: {folder_path or '(порожній шлях)'}")
    if not new_name.strip():
        raise ValueError("Не вказано нову назву папки")
    target = source.with_name(new_name.strip())
    old_name = source.name
    renamed = source.rename(target).resolve()
    return {
        "path": str(renamed),
        "old_name": old_name,
        "new_name": renamed.name,
        "message": f"Папку перейменовано: {old_name} → {renamed.name}",
    }


@node(
    name="Перемістити файл у папку",
    category="Файли",
    description="Переміщує файл у задану папку призначення.",
    type_id="builtin.files.move_file",
    outputs={"new_path": "str", "file_name": "str", "message": "str"},
    execution_inputs=("exec",),
    execution_outputs=("then",),
    preview_policy="never",
)
def move_file(file_path: str = "", destination_folder: str = "", overwrite: bool = True) -> dict:
    import shutil
    source = Path(file_path).expanduser()
    if not source.is_file():
        raise FileNotFoundError(f"Файл не знайдено: {file_path or '(порожній шлях)'}")
    dest_dir = Path(destination_folder).expanduser()
    dest_dir.mkdir(parents=True, exist_ok=True)
    target = dest_dir / source.name

    if target.exists() and not overwrite and target != source:
        raise FileExistsError(f"Файл уже існує у папці призначення: {target}")
    if target.exists() and target != source:
        target.unlink()

    moved = shutil.move(str(source), str(target))
    resolved = Path(moved).resolve()
    return {
        "new_path": str(resolved),
        "file_name": resolved.name,
        "message": f"Файл {source.name} переміщено у {dest_dir.name}",
    }
