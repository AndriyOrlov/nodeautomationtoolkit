"""Жодна функція не має читати імені, якого ніде не звʼязано.

Цей клас дефектів у проєкті найдорожчий: файл компілюється, тести на окремих
модулях проходять, а падає воно вже у користувача — посеред пакета, після
всієї виконаної роботи, повідомленням `name '…' is not defined`.

Приклади, на яких це вже ловилося:

* `_slash_to_lines`, `executor_start_from_bookmark`, `clean_redundant_blanks` —
  викликались із чужої області видимості;
* `preview_delay` — присвоєння потрапило в `run_extracts` замість
  `run_generate_copies`, і **жоден** із 13 наказів не сформувався;
* `DataTable` у COM-гілці витягів — імпорт стояв у сусідніх функціях,
  але не в тій, що ним користується.

`py_compile` такого не бачить: синтаксис правильний. Готовий лінтер поставити
не можна — правило 1.1 забороняє мережу, — тому обхід свій, у
`tests/_undefined_names.py`.
"""

import pathlib

import pytest

from _undefined_names import find_undefined_names

ROOT = pathlib.Path(__file__).resolve().parent.parent

# Файли, які виконують справжню роботу. Тести сюди не входять: у них є
# фікстури pytest, які звʼязуються не через звичайне присвоєння.
CHECKED = [
    ROOT / "generate_extracts.py",
    *sorted((ROOT / "src" / "nodeautomationtoolkit").rglob("*.py")),
    *sorted((ROOT / "scripts").rglob("*.py")),
]


# Відомі НЕВИПРАВЛЕНІ дефекти. Вони справжні, але лежать поза генератором
# наказів, і що саме мало бути на місці цих імен — не очевидно, тож вгадувати
# не можна. Список звужується, а не росте: якщо дефект виправлять, тест
# скаже прибрати рядок звідси; якщо зʼявиться новий — тест впаде.
KNOWN_UNFIXED = {
    # Діалог PyQt: `_run` читає `before_text`/`after_text`/`extra`, яких у
    # ньому ніхто не створює. Поруч є віджети `self._before`, `self._after`,
    # `self._extra` — схоже, при правці загубились виклики до них. До
    # генератора наказів цей файл не належить.
    "src/nodeautomationtoolkit/ui/order_analysis_dialog.py": {
        ("OrderAnalysisDialog._run", "before_text"),
        ("OrderAnalysisDialog._run", "after_text"),
        ("OrderAnalysisDialog._run", "extra"),
    },
    # COM-гілка витягів: `DataTable` не імпортовано, тож гілка завжди падає
    # `NameError`, збій ковтається, і роботу робить запасний шлях на
    # python-docx. Дописати імпорт НЕ МОЖНА: це вмикає мертву гілку, і
    # `test_signatory_slash_line_breaks` одразу падає. Гілку треба лагодити
    # цілком (разом із полями сторінки, розд. 12.2) або прибирати.
    "src/nodeautomationtoolkit/builtin_nodes/word_batch.py": {
        ("_try_create_extracts_via_word_com", "DataTable"),
    },
}


def _describe(problems) -> str:
    return "\n".join(
        f"  {where} читає «{name}» (рядок {lineno})" for lineno, where, name in problems
    )


@pytest.mark.parametrize("path", CHECKED, ids=lambda p: str(p.relative_to(ROOT)))
def test_every_name_a_function_reads_is_bound_somewhere(path):
    relative = path.relative_to(ROOT).as_posix()
    problems = find_undefined_names(path)
    known = KNOWN_UNFIXED.get(relative, set())

    fresh = [item for item in problems if (item[1], item[2]) not in known]
    assert not fresh, (
        f"{relative}: імена без звʼязування — це `NameError` у користувача, "
        f"а не помилка компіляції:\n{_describe(fresh)}"
    )

    still_broken = {(where, name) for _, where, name in problems}
    assert known <= still_broken, (
        f"{relative}: ці дефекти вже виправлено — приберіть їх із KNOWN_UNFIXED: "
        f"{sorted(known - still_broken)}"
    )


def test_the_check_actually_catches_a_missing_name(tmp_path):
    """Перевірка не має бути порожньою: підкладаємо саме той дефект."""
    broken = tmp_path / "broken.py"
    broken.write_text(
        "def outer():\n"
        "    preview_delay = 1.5\n"
        "\n"
        "def inner():\n"
        "    return preview_delay\n",
        encoding="utf-8",
    )

    problems = find_undefined_names(broken)

    assert [(where, name) for _, where, name in problems] == [("inner", "preview_delay")]


def test_a_closure_over_the_enclosing_function_is_not_a_problem():
    """Вкладена функція бачить імена зовнішньої — це нормально."""
    import tempfile

    with tempfile.TemporaryDirectory() as folder:
        path = pathlib.Path(folder) / "ok.py"
        path.write_text(
            "def outer():\n"
            "    value = 1\n"
            "    def inner():\n"
            "        return value\n"
            "    return inner()\n",
            encoding="utf-8",
        )

        assert find_undefined_names(path) == []
