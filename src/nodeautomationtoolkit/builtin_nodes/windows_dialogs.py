from __future__ import annotations

from nodeautomationtoolkit.core.definition import node


def _qt_widgets():
    from PySide6.QtWidgets import QApplication, QFileDialog

    application = QApplication.instance()
    if application is None:
        raise RuntimeError("Системне вікно вибору файла доступне лише під час запуску застосунку")
    return QFileDialog


@node(
    name="Відкрити файл",
    category="Windows",
    description="Показує стандартне системне вікно Windows і повертає вибраний шлях.",
    type_id="builtin.windows.open_file",
)
def open_file_dialog(
    title: str = "Виберіть файл",
    file_filter: str = "Усі файли (*.*)",
    initial_folder: str = "",
) -> str:
    file_dialog = _qt_widgets()
    path, _selected_filter = file_dialog.getOpenFileName(
        None,
        title,
        initial_folder,
        file_filter,
    )
    return path


@node(
    name="Відкрити декілька файлів",
    category="Windows",
    description="Показує стандартне системне вікно Windows для вибору кількох файлів.",
    type_id="builtin.windows.open_files",
)
def open_files_dialog(
    title: str = "Виберіть файли",
    file_filter: str = "Усі файли (*.*)",
    initial_folder: str = "",
) -> list[str]:
    file_dialog = _qt_widgets()
    paths, _selected_filter = file_dialog.getOpenFileNames(
        None,
        title,
        initial_folder,
        file_filter,
    )
    return paths


@node(
    name="Вибрати папку",
    category="Windows",
    description="Показує стандартне системне вікно Windows для вибору папки.",
    type_id="builtin.windows.select_folder",
)
def select_folder_dialog(
    title: str = "Виберіть папку",
    initial_folder: str = "",
) -> str:
    file_dialog = _qt_widgets()
    return file_dialog.getExistingDirectory(None, title, initial_folder)


@node(
    name="Зберегти файл як",
    category="Windows",
    description="Показує стандартне системне вікно Windows і повертає шлях збереження.",
    type_id="builtin.windows.save_file",
)
def save_file_dialog(
    title: str = "Зберегти файл як",
    suggested_name: str = "",
    file_filter: str = "Усі файли (*.*)",
) -> str:
    file_dialog = _qt_widgets()
    path, _selected_filter = file_dialog.getSaveFileName(
        None,
        title,
        suggested_name,
        file_filter,
    )
    return path
