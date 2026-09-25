"""Приклад наказу з навмисними помилками — щоб перевірити макрос у своєму Word.

Усе вигадане. Запуск: python scripts/word_macro/make_example.py
Очікуване — у word_macro/README.md, розділ «Перевірка, що все працює».
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from docx import Document
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[2]
OUTPUT = ROOT / "word_macro" / "Приклад наказу для перевірки.docx"


def fake_ipn(birth: date, serial: int = 123) -> str:
    days = (birth - date(1899, 12, 31)).days
    digits = f"{days:05d}{serial:03d}1"
    weights = (-1, 5, 7, 9, 4, 6, 10, 5, 7)
    return digits + str(sum(int(d) * w for d, w in zip(digits, weights, strict=True)) % 11 % 10)


def main() -> int:
    good = fake_ipn(date(1988, 3, 1))
    typo = good[:-1] + str((int(good[-1]) + 1) % 10)
    paragraphs = [
        "НАКАЗ",
        "§ 1",
        "Відповідно до пункту 1 Положення нижчепойменованих офіцерів ЗВІЛЬНИТИ з займаних "
        "посад і ПРИЗНАЧИТИ:",
        "1. Капітана ТЕСТЕНКА Олега Васильовича, командира механізованої роти 901 механізованого "
        "батальйону - НАЧАЛЬНИКОМ ШТАБУУ 901 МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ.",
        "02.03.1988 р.н., освіта: вища.",
        f"{good}.",
        "§ 2",
        "Відповідно до підпункту «б» пункту 2 частини п'ятої статті 26 Закону України "
        "нижчепойменованих офіцерів ЗВІЛЬНИТИ З ВІЙСЬКОВОЇ СЛУЖБИ:",
        "3. Полковника ПРИКЛАДЕНКА Івана Петровича, командира бригади.",
        "1985 р.н.",
        f"{typo}.",
        "",
        "Командир військової частини А0001",
        "полковник                    Петро ТЕСТОВИЙ",
    ]
    document = Document()
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(14)
    for text in paragraphs:
        document.add_paragraph(text)
    document.save(OUTPUT)
    print(OUTPUT)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
