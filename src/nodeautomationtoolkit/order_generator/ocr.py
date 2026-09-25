"""Розпізнавання сканів українською — локальний Tesseract.

Плани й подання часто приходять сканами. Tesseract запускається як звичайна
програма на цій машині: жодних хмарних сервісів (PROJECT_RULES 1.1).

Потрібне:
- сам Tesseract (`tesseract.exe`) і мовний файл `ukr.traineddata`;
- для PDF — `pdftoppm` (Poppler) або `mutool`, щоб зробити зі сторінок зображення.

Шлях береться з `NAT_TESSERACT`, з PATH або зі звичайних місць встановлення.
Якщо нічого не знайдено, функції кидають `OcrUnavailable` з підказкою.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
DEFAULT_LANGUAGES = "ukr+eng"
_WINDOWS_PATHS = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    r"C:\Tesseract-OCR\tesseract.exe",
)
_INSTALL_HINT = (
    "Tesseract не знайдено. Встановіть його локально (наприклад, "
    "https://github.com/UB-Mannheim/tesseract/wiki), додайте мову «ukr» "
    "і, якщо він не в PATH, укажіть шлях у змінній середовища NAT_TESSERACT."
)


class OcrUnavailable(RuntimeError):
    """Tesseract або потрібні складники не встановлені."""


def find_tesseract() -> str | None:
    """Шлях до tesseract.exe або None."""
    from_env = os.environ.get("NAT_TESSERACT", "").strip('" ')
    if from_env and Path(from_env).is_file():
        return from_env
    found = shutil.which("tesseract")
    if found:
        return found
    for candidate in _WINDOWS_PATHS:
        if Path(candidate).is_file():
            return candidate
    return None


def _run(command: list[str]) -> str:
    result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if result.returncode != 0:
        raise OcrUnavailable(f"{Path(command[0]).name}: {result.stderr.strip()[:200]}")
    return result.stdout


def languages() -> list[str]:
    """Мови, які вміє встановлений Tesseract («ukr», «eng» …)."""
    executable = find_tesseract()
    if not executable:
        return []
    try:
        output = _run([executable, "--list-langs"])
    except OcrUnavailable:
        return []
    return [line.strip() for line in output.splitlines()[1:] if line.strip()]


def is_available(language: str = "ukr") -> bool:
    return bool(find_tesseract()) and language in languages()


def _check_language(language: str) -> str:
    executable = find_tesseract()
    if not executable:
        raise OcrUnavailable(_INSTALL_HINT)
    installed = languages()
    wanted = [part for part in language.split("+") if part]
    missing = [part for part in wanted if part not in installed]
    if missing == wanted:
        raise OcrUnavailable(
            f"Для Tesseract немає мов {', '.join(missing)}; встановлені: {', '.join(installed) or '—'}"
        )
    return "+".join(part for part in wanted if part in installed) or "eng"


def ocr_image(path: str | Path, language: str = DEFAULT_LANGUAGES) -> str:
    """Текст із зображення (скан сторінки)."""
    executable = find_tesseract()
    if not executable:
        raise OcrUnavailable(_INSTALL_HINT)
    language = _check_language(language)
    with tempfile.TemporaryDirectory(prefix="nat_ocr_") as folder:
        base = Path(folder) / "page"
        _run([executable, str(path), str(base), "-l", language, "--psm", "6"])
        return (base.with_suffix(".txt")).read_text(encoding="utf-8", errors="replace")


def _pdf_pages(path: Path, folder: Path) -> list[Path]:
    """Сторінки PDF як зображення: pdftoppm або mutool."""
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        _run([pdftoppm, "-r", "300", "-png", str(path), str(folder / "page")])
    else:
        mutool = shutil.which("mutool")
        if not mutool:
            raise OcrUnavailable(
                "Для розпізнавання PDF потрібен pdftoppm (Poppler) або mutool (MuPDF) на цій машині"
            )
        _run([mutool, "draw", "-r", "300", "-o", str(folder / "page-%d.png"), str(path)])
    return sorted(folder.glob("page*.png"), key=lambda item: _page_number(item.name))


def _page_number(name: str) -> int:
    match = re.search(r"(\d+)", name)
    return int(match.group(1)) if match else 0


def ocr_pdf(path: str | Path, language: str = DEFAULT_LANGUAGES, max_pages: int = 50) -> str:
    """Текст зі сканованого PDF (посторінково)."""
    path = Path(path)
    with tempfile.TemporaryDirectory(prefix="nat_ocr_pdf_") as folder:
        pages = _pdf_pages(path, Path(folder))[:max_pages]
        return "\n".join(ocr_image(page, language) for page in pages)


def ocr_file(path: str | Path, language: str = DEFAULT_LANGUAGES) -> str:
    path = Path(path)
    if path.suffix.casefold() == ".pdf":
        return ocr_pdf(path, language)
    return ocr_image(path, language)
