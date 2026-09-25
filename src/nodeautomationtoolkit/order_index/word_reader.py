"""Текст старих форматів (.doc, .rtf) через Word — лише локально.

- Відкривається **тимчасова копія**, а не оригінал (PROJECT_RULES 1.3): після
  примусово закритого Word оригінал із `~$`-замком відкривається з зависанням,
  копія — миттєво.
- Розширення копії береться за вмістом: `.doc`, що насправді є `.docx`,
  читається без Word.
- Власний прихований екземпляр Word (`DispatchEx`), макроси вимкнено,
  документи з паролем не відкриваються (фіктивний пароль → помилка, а не діалог).
- Текст береться одним зверненням `Content.Text` і ділиться на абзаци за `\\r`
  та кінцями клітинок `\\x07`, щоб таблиця не склеювалась в один рядок.
"""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import time
from pathlib import Path

from .extractor import read_docx_paragraphs

_OOXML_SIGNATURE = b"PK\x03\x04"
_OLE_SIGNATURE = b"\xd0\xcf\x11\xe0"
_RTF_SIGNATURE = b"{\\rtf"
_REJECTED_BY_CALLEE = -2147418111
_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def sniff_suffix(path: str | Path) -> str:
    with open(path, "rb") as handle:
        head = handle.read(8)
    if head.startswith(_OOXML_SIGNATURE):
        return ".docx"
    if head.startswith(_OLE_SIGNATURE):
        return ".doc"
    if head.startswith(_RTF_SIGNATURE):
        return ".rtf"
    return Path(path).suffix.lower()


def split_word_text(text: str) -> list[str]:
    text = (text or "").replace("\x0b", " ").replace("\x0c", "\r").replace("\x07", "\r")
    return [re.sub(r"[\x00-\x08\x0e-\x1f]", "", part) for part in text.split("\r")]


class WordUnavailable(RuntimeError):
    pass


class WordTextReader:
    """Один прихований Word на весь прогін; закривається в `close()`."""

    def __init__(self):
        self._word = None
        self._com_initialized = False
        self._temp = tempfile.mkdtemp(prefix="nat_order_index_")

    def read_paragraphs(self, path: str | Path) -> list[str]:
        suffix = sniff_suffix(path)
        if suffix == ".docx":
            return read_docx_paragraphs(path)
        copy = Path(self._temp) / f"source{suffix}"
        shutil.copy2(path, copy)
        try:
            return split_word_text(self._read_with_word(copy))
        finally:
            try:
                os.remove(copy)
            except OSError:
                pass

    def _application(self):
        if self._word is None:
            try:
                import pythoncom
                import win32com.client
            except ImportError as error:
                raise WordUnavailable("pywin32 не встановлено") from error
            pythoncom.CoInitialize()
            self._com_initialized = True
            try:
                word = win32com.client.DispatchEx("Word.Application")
            except Exception as error:
                raise WordUnavailable(f"Word не запускається: {error}") from error
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                word.AutomationSecurity = _MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:
                pass
            self._word = word
        return self._word

    def _read_with_word(self, copy: Path) -> str:
        word = self._application()
        document = _retry(
            lambda: word.Documents.Open(
                FileName=str(copy),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument="nat-index-no-password",
                Visible=False,
                NoEncodingDialog=True,
            )
        )
        try:
            return str(_retry(lambda: document.Content.Text) or "")
        finally:
            try:
                document.Close(SaveChanges=0)
            except Exception:
                pass

    def close(self) -> None:
        if self._word is not None:
            try:
                self._word.Quit(SaveChanges=0)
            except Exception:
                pass
            self._word = None
        if self._com_initialized:
            import pythoncom

            pythoncom.CoUninitialize()
            self._com_initialized = False
        shutil.rmtree(self._temp, ignore_errors=True)

    def __enter__(self) -> WordTextReader:
        return self

    def __exit__(self, *_exc) -> None:
        self.close()


def _retry(call, attempts: int = 6):
    for attempt in range(attempts):
        try:
            return call()
        except Exception as error:
            hresult = getattr(error, "hresult", None) or (error.args[0] if error.args else None)
            if hresult != _REJECTED_BY_CALLEE or attempt == attempts - 1:
                raise
            time.sleep(0.5)
    return None
