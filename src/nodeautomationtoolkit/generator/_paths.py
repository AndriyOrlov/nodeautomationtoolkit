"""Тека програми — там лежить `config.json` генератора.

Раніше це була тека `generate_extracts.py` (`os.path.dirname(__file__)`). Після
винесення логіки в пакет фасад `generate_extracts.py` записує сюди саме свою
теку, тож розташування конфігу не змінилося — і в зібраному EXE теж.
"""
from pathlib import Path

#: Типове значення — корінь проєкту (src/nodeautomationtoolkit/generator → 3 рівні вгору).
APP_DIR = str(Path(__file__).resolve().parents[3])
