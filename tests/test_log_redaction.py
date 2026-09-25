"""Знеособлення журналу: структура лишається, назви частин і ПІБ — зникають.

Журнал збою треба комусь показувати, а в ньому світяться шифри, скорочення
частин, прізвища й повні шляхи на робочому диску. Кнопка «🔒 Копіювати
знеособлено» проганяє текст через `redact_sensitive_text`.

Значення тут ВИГАДАНІ, лише формою схожі на справжні (правило «публічний
репозиторій — без реальних даних»).
"""

import importlib.util
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))


def _load_generator():
    spec = importlib.util.spec_from_file_location(
        "generate_extracts_redaction_tests", PROJECT_ROOT / "generate_extracts.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


generator = _load_generator()
redact = generator.redact_sensitive_text

LOG = """[1/4] Генеруємо витяг для: 199омбр А9998
  ↧ 199омбр А9998: виконавця опущено вниз сторінки (+39 ентер(ів)).
  ✓ 199омбр А9998: 5 стор., 10.1 с.
[2/4] Генеруємо витяг для: 99 АК А9997
  ✓ 99 АК А9997: 3 стор., 6.8 с.
[3/4] Генеруємо витяг для: Тестівський ОТЦК та СП
Підписант оригіналу наказу переноситься у витяг: полковник Петро ВИГАДАНКО.
Збережено файл: E:\\Робота\\Накази\\2026\\Витяги наказу № 999.docx
📄 Загальна кількість сторінок у документі: 14 стор.
Зчитано 116 записів з Excel. Скорочень з колонки C: 107.
"""


def test_ciphers_units_names_and_paths_disappear():
    safe = redact(LOG)

    for secret in ("А9998", "А9997", "199омбр", "99 АК", "ВИГАДАНКО", "Тестівський",
                   "E:\\Робота", "Витяги наказу № 999.docx"):
        assert secret not in safe, f"у знеособленому журналі лишилось: {secret}"


def test_structure_and_statistics_survive():
    safe = redact(LOG)

    for kept in ("Генеруємо витяг для", "виконавця опущено вниз сторінки", "+39 ентер(ів)",
                 "5 стор.", "10.1 с.", "14 стор.", "116 записів", "107"):
        assert kept in safe, f"знеособлення зʼїло корисне: {kept}"


def test_same_unit_gets_the_same_label():
    safe = redact(LOG)

    first_line = safe.splitlines()[0]
    label = first_line.split(": ")[-1].strip()
    assert label, first_line
    # та сама частина згадується в журналі тричі — мітка має бути однакова
    assert safe.count(label) >= 3, safe


def test_different_units_get_different_labels():
    safe = redact(LOG)
    labels = {line.split(": ")[-1].strip() for line in safe.splitlines()
              if line.startswith("[") and "Генеруємо витяг" in line}
    assert len(labels) == 3, labels


def test_empty_and_clean_text_is_unchanged():
    assert redact("") == ""
    assert redact("Готовий до роботи.") == "Готовий до роботи."


@pytest.mark.parametrize("name", [
    "Петро Вигаданко", "Вигаданко Петро Іванович",
    "ВИГАДАНКО Петро Іванович", "Петро Іванович ВИГАДАНКО",
    "Петро ВИГАДАНКО", "Анна-Марія Вигаданко",
])
def test_full_name_is_removed_in_both_orders_and_cases(name):
    safe = redact(f"Підписант: полковник {name}.")
    for word in name.split():
        assert word not in safe
    assert "Підписант: полковник" in safe


@pytest.mark.parametrize("name", [
    "Тестівського обласного територіального центру комплектування та соціальної підтримки",
    "Тестівському обласному ТЦК та СП", "ТЕСТІВСЬКОГО ОБЛАСНОГО ТЦК ТА СП",
])
def test_tck_name_is_removed_in_declined_forms(name):
    safe = redact(f"Адресат: {name}.")
    assert "тестів" not in safe.casefold()
    assert "Адресат:" in safe


def test_routing_version_after_a_path_stays_readable():
    """Шлях прибирається, а те, що стоїть після « · », лишається.

    Раніше правило шляху з'їдало рядок до кінця, і в знеособленому журналі не
    було видно навіть версії модуля маршрутизації — єдиної ознаки того, що
    програму перезапущено після оновлення.
    """
    line = (
        "Модуль маршрутизації: C:\\Users\\Тест\\Нова папка\\nat\\src\\recipient_mapping.py"
        " · версія 2026-09-15-v12-internal-context-only"
    )
    assert redact(line) == "Модуль маршрутизації: <шлях> · версія 2026-09-15-v12-internal-context-only"


def test_excel_reference_keeps_date_and_size_but_not_the_path():
    safe = redact("Еталон Excel: E:/Робота/Нова папка/словник.xlsx · змінено 15.09.2026 12:11:52 · 5470 байт.")
    assert "Робота" not in safe and "словник" not in safe
    assert safe.startswith("Еталон Excel: <шлях> · змінено 15.09.2026 12:11:52")


def test_path_without_separator_is_removed_to_the_end_of_line():
    safe = redact("Збережено файл: E:\\Робота\\Накази\\Витяги наказу № 999.docx\nДалі звичайний рядок.")
    assert safe.splitlines() == ["Збережено файл: <шлях>", "Далі звичайний рядок."]
