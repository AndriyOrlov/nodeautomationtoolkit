"""Вікно «Зразки та шаблони»: лише файли-зразки, за якими генеруються документи.

Виконавець, папки результатів і засвідчувачі сюди свідомо не входять: вони
змінюються від наказу до наказу й тому стоять на самих вкладках генерації.
"""

from __future__ import annotations

from typing import Callable

from PySide6.QtWidgets import QDialog, QGridLayout, QHBoxLayout, QTabWidget, QVBoxLayout, QWidget

from .widgets import button, label, line_edit


class SamplesDialog(QDialog):
    def __init__(self, shell, parent: QWidget | None = None):
        super().__init__(parent)
        self._shell = shell
        self.setObjectName("SamplesDialog")
        self.setWindowTitle("Зразки та шаблони")
        self.resize(880, 440)
        self.finished.connect(self.deleteLater)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        intro = label(
            "Тут лише зразки, за якими генеруються документи. Виконавець, папки результатів "
            "і засвідчувачі — на вкладках генерації: вони змінюються від наказу до наказу.",
            "MutedLabel",
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        tabs = QTabWidget()
        tabs.setObjectName("ResultTabs")
        tabs.addTab(
            self._page(
                [
                    ("Словник (Excel):", shell.excel_path, shell.select_excel, None),
                    ("Зразок витягу:", shell.template_path, shell.select_template, shell.edit_template),
                ],
                "Зразок витягу — DOCX із тегами {{кому}}, {{куди}}, {{зміст}}, {{пункти}} "
                "та реквізитами підписанта й засвідчувача.",
                [("ⓘ Теги шаблону", shell.show_template_tags)],
            ),
            "  📄 Витяги  ",
        )
        tabs.addTab(
            self._page(
                [
                    (
                        "Заготовка примірника:",
                        shell.p2_back_page_path,
                        shell.select_p2_back_page,
                        shell.edit_p2_back_page,
                    ),
                ],
                "Заготовка примірника — повноцінний зразок (шапка, теги, {{зміст}}), "
                "у який переноситься текст наказу.",
                [("ⓘ Правила примірника 2", shell.show_p2_info)],
            ),
            "  📑 Примірники 2/3  ",
        )
        tabs.addTab(
            self._page(
                [
                    (
                        "Зразок супроводу:",
                        shell.message_cover_template_path,
                        shell.select_message_cover_template,
                        None,
                    ),
                    (
                        "Зразок зі змістом:",
                        shell.message_content_template_path,
                        shell.select_message_content_template,
                        None,
                    ),
                ],
                "Два зразки: титульне повідомлення й повідомлення зі змістом наказу.",
                [("ⓘ Теги повідомлень", shell.show_message_tags)],
            ),
            "  💬 Повідомлення  ",
        )
        layout.addWidget(tabs, 1)

        bottom = QHBoxLayout()
        bottom.addStretch(1)
        bottom.addWidget(button("💾 Зберегти та закрити", "run", self.accept))
        layout.addLayout(bottom)

    @staticmethod
    def _page(
        rows: list[tuple[str, object, Callable, Callable | None]],
        note: str,
        extra_buttons: list[tuple[str, Callable]],
    ) -> QWidget:
        page = QWidget()
        grid = QGridLayout(page)
        grid.setContentsMargins(12, 12, 12, 12)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(8)
        for row, (caption, variable, pick, open_file) in enumerate(rows):
            grid.addWidget(label(caption, "FieldLabel"), row, 0)
            grid.addWidget(line_edit(variable), row, 1)
            actions = QHBoxLayout()
            actions.setSpacing(4)
            actions.addWidget(button("📂 Обрати", "secondary", pick))
            if open_file is not None:
                actions.addWidget(button("✎ Відкрити у Word", "ghost", open_file))
            # Без розпірки єдина кнопка рядка розтягувалась на всю колонку.
            actions.addStretch(1)
            grid.addLayout(actions, row, 2)
        grid.setColumnStretch(1, 1)

        note_label = label(note, "MutedLabel")
        note_label.setWordWrap(True)
        grid.addWidget(note_label, len(rows), 0, 1, 3)

        extras = QHBoxLayout()
        for caption, slot in extra_buttons:
            extras.addWidget(button(caption, "sky", slot))
        extras.addStretch(1)
        grid.addLayout(extras, len(rows) + 1, 0, 1, 3)
        grid.setRowStretch(len(rows) + 2, 1)
        return page

    def done(self, result: int) -> None:
        # Esc, хрестик і кнопка ведуть сюди — налаштування зберігаються завжди.
        self._shell.save_config()
        self._shell.log("Зразки та шаблони збережено.")
        super().done(result)
