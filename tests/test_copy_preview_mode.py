"""Режим превʼю: покроковий показ збірки примірника.

Що саме перевіряється:

* коли режим ВИМКНЕНО — не пишеться нічого й не чекається нічого, тобто
  звичайний прогін не платить за цей режим ані рядком журналу, ані часом;
* коли УВІМКНЕНО — кроки нумеруються «N/8», пауза йде ПІСЛЯ дії (щоб було
  видно її результат, а не напис), а хибне значення паузи не зриває роботу.
"""

import pytest

from nodeautomationtoolkit.builtin_nodes.copy_generator import (
    PREVIEW_DEFAULT_DELAY,
    PREVIEW_TOTAL_STEPS,
    PreviewSteps,
)


class _Recorder:
    """Збирає журнал і паузи в одну стрічку подій, щоб бачити їх ПОРЯДОК."""

    def __init__(self):
        self.events: list[tuple[str, object]] = []

    def log(self, message):
        self.events.append(("log", message))

    def sleep(self, seconds):
        self.events.append(("sleep", seconds))

    @property
    def logs(self):
        return [value for kind, value in self.events if kind == "log"]

    @property
    def sleeps(self):
        return [value for kind, value in self.events if kind == "sleep"]


def _steps(recorder, **kwargs):
    kwargs.setdefault("delay", 1.5)
    kwargs.setdefault("enabled", True)
    return PreviewSteps(log=recorder.log, sleeper=recorder.sleep, **kwargs)


# --------------------------------------------------------------------------
# Вимкнений режим нічого не коштує
# --------------------------------------------------------------------------


def test_disabled_mode_writes_nothing_and_waits_for_nothing():
    recorder = _Recorder()
    steps = _steps(recorder, enabled=False)

    with steps.step("відкриваю наказ"):
        steps.detail("абзаци 12–147")

    assert recorder.events == []


def test_disabled_mode_still_runs_the_work():
    """Показник не має впливати на те, що робиться всередині кроку."""
    recorder = _Recorder()
    steps = _steps(recorder, enabled=False)
    done = []

    with steps.step("крок"):
        done.append("зроблено")

    assert done == ["зроблено"]


def test_default_preview_object_is_disabled():
    """`build_copy_document` без `preview` має поводитись як завжди."""
    assert PreviewSteps().enabled is False


# --------------------------------------------------------------------------
# Увімкнений режим
# --------------------------------------------------------------------------


def test_steps_are_numbered_out_of_the_total():
    recorder = _Recorder()
    steps = _steps(recorder)

    with steps.step("відкриваю наказ"):
        pass
    with steps.step("шукаю межі тіла та підписанта"):
        pass

    assert recorder.logs == [
        f"  ▶ Крок 1/{PREVIEW_TOTAL_STEPS}: відкриваю наказ",
        f"  ▶ Крок 2/{PREVIEW_TOTAL_STEPS}: шукаю межі тіла та підписанта",
    ]


def test_pause_comes_after_the_work_not_before():
    """Сенс режиму — побачити РЕЗУЛЬТАТ кроку у вікні Word."""
    recorder = _Recorder()
    steps = _steps(recorder)

    with steps.step("переношу зміст"):
        recorder.events.append(("work", None))

    kinds = [kind for kind, _ in recorder.events]
    assert kinds == ["log", "work", "sleep"]


def test_detail_belongs_to_the_current_step():
    recorder = _Recorder()
    steps = _steps(recorder)

    with steps.step("шукаю межі тіла та підписанта"):
        steps.detail("абзаци 12–147")

    assert recorder.logs[1] == "     → абзаци 12–147"


def test_the_configured_delay_is_used():
    recorder = _Recorder()
    steps = _steps(recorder, delay=0.25)

    with steps.step("крок"):
        pass

    assert recorder.sleeps == [0.25]


def test_zero_delay_skips_waiting_but_keeps_the_log():
    """Нуль — це «показуй кроки, але не гальмуй»."""
    recorder = _Recorder()
    steps = _steps(recorder, delay=0)

    with steps.step("крок"):
        pass

    assert recorder.sleeps == []
    assert len(recorder.logs) == 1


def test_a_failed_step_does_not_pause():
    """Збій має видно спливти нагору, а не чекати ще півтори секунди."""
    recorder = _Recorder()
    steps = _steps(recorder)

    with pytest.raises(ValueError):
        with steps.step("крок"):
            raise ValueError("не вдалося визначити межі тіла наказу")

    assert recorder.sleeps == []


# --------------------------------------------------------------------------
# Поле паузи заповнює користувач — воно може містити будь-що
# --------------------------------------------------------------------------


@pytest.mark.parametrize("raw", ["", "   ", "півтори", None, "abc"])
def test_broken_delay_falls_back_instead_of_breaking_generation(raw):
    steps = PreviewSteps(delay=raw, enabled=True)

    assert steps.delay == PREVIEW_DEFAULT_DELAY


def test_delay_arrives_as_text_from_the_entry_field():
    assert PreviewSteps(delay="2.5", enabled=True).delay == 2.5


def test_negative_delay_is_clamped_to_zero():
    assert PreviewSteps(delay=-5, enabled=True).delay == 0.0


# --------------------------------------------------------------------------
# Нумерація не має тривати наскрізно через пакет
# --------------------------------------------------------------------------


def test_each_order_starts_counting_from_one():
    """У пакеті на кожен наказ створюється СВІЙ показник."""
    recorder = _Recorder()

    for _ in range(2):
        steps = _steps(recorder)
        with steps.step("відкриваю наказ"):
            pass

    assert recorder.logs == [f"  ▶ Крок 1/{PREVIEW_TOTAL_STEPS}: відкриваю наказ"] * 2


# --------------------------------------------------------------------------
# Загальна кількість кроків має відповідати реальній збірці
# --------------------------------------------------------------------------


def test_total_matches_the_steps_actually_in_build_copy_document():
    """Інакше в журналі зʼявиться «Крок 9/8».

    `PREVIEW_TOTAL_STEPS` — це знаменник у написі, тож він мусить дорівнювати
    кількості кроків у самій збірці. Рахуємо їх у коді функції, щоб доданий
    колись девʼятий крок одразу провалив тест, а не збив нумерацію мовчки.
    """
    import inspect

    from nodeautomationtoolkit.builtin_nodes import copy_generator

    source = inspect.getsource(copy_generator.build_copy_document)

    assert source.count("steps.step(") == PREVIEW_TOTAL_STEPS
