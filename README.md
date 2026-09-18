# Генератор витягів, примірників і повідомлень за наказами

Локальна програма для Windows: з наказу по особовому складу (DOCX) створює
примірники № 2, розрахунок розсилки й адресні витяги, а також шифровані повідомлення
про прийняття кадрових рішень. Усе працює лише на цьому комп'ютері — без мережі.

Докладна інструкція для користувача — кнопка **📖 Інструкція** у вікні програми
(`src/nodeautomationtoolkit/generator_qt/instruction.md`).

## Запуск

- `start_generator_qt.bat` — основне вікно (PySide6). Бере Python з теки `python`
  поруч (portable), з `.venv` або з системи; зміни в коді підхоплюються після
  перезапуску.
- `build_generator_qt.bat` — збірка `dist\GeneratorVytyagivQt\GeneratorVytyagivQt.exe`
  (PyInstaller). Кожна збірка підвищує версію на 0.1.

Потрібен установлений Microsoft Word: документи формуються через нього.

## Розробка

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev,windows]"
python -m pytest
```

Portable-Python без встановлення — `scripts/portable/README.md`.

## Структура

- `generate_extracts.py` — логіка генератора (клас `App`): маршрутизація, робота з Word,
  витяги, примірники, повідомлення.
- `generate_extracts_qt.py` + `src/nodeautomationtoolkit/generator_qt` — вікно PySide6.
- `src/nodeautomationtoolkit/builtin_nodes` — модулі логіки: таблиця частин і
  маршрутизація (`recipient_mapping`), повідомлення (`message_order`), примірники
  (`copy_generator`), порівняння документів, теги шаблонів.
- `src/nodeautomationtoolkit/personnel` — довідники (CSV): звання, посади, відмінки,
  рівні шпк, перевірка РНОКПП.
- `src/nodeautomationtoolkit/order_index` — локальний індекс попередніх наказів
  (посади, підрозділи, пункти про осіб); лежить лише на комп'ютері користувача.
- `src/nodeautomationtoolkit/order_generator` — генератор наказів по особовому складу:
  індексація → генерація → перевірка → збірка (`pipeline.py`). Види наказів:
  призначення (переміщення), звільнення з військової служби, присвоєння військових
  звань. Ще в роботі: у вікні програми він за кнопкою **🔒 Накази** й тимчасовим
  паролем.
- `scripts/e2e_*` — наскрізні перевірки на вигаданих наказах; `scripts/diagnostics` —
  знеособлений звіт.
- `tests` — тести.

Правила проєкту — `PROJECT_RULES.md`, робочі нотатки й домовленості — `AGENT.md`.
