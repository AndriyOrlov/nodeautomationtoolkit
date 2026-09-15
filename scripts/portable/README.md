# Portable-Python поруч із програмою

**Уже зібрано.** Тека `python\` у корені проєкту містить самодостатній
інтерпретатор із усіма залежностями. `start_generator.bat` бере саме його, тож
програма запускається з флешки на будь-якому ПК — **без установлення Python**.

| | |
| --- | --- |
| версія | Python **3.14.2**, x64 (embeddable) |
| розмір | ≈ 54 МБ, ~2300 файлів |
| у git | **не потрапляє** — `python/` у `.gitignore` |

## Що всередині

| що | навіщо |
| --- | --- |
| embeddable-збірка Python 3.14.2 | сам інтерпретатор |
| `tkinter\`, `tcl\`, `_tkinter.pyd`, `tcl86t.dll`, `tk86t.dll`, `zlib1.dll` | **інтерфейс.** В embeddable-збірці tkinter НЕМАЄ — докладено вручну |
| `pywin32` | керування Word через COM |
| `ttkbootstrap` | тема інтерфейсу; без неї GUI падає (AGENT.md 8.4) |
| `openpyxl` | Excel-словник адресатів |
| `python-docx` | скрипти перевірки в `scripts\e2e_extracts\` |

## Що потрібно на цільовому ПК

- **Microsoft Word** — програма керує ним через COM, замінити нічим.
- Більше нічого: ані Python, ані бібліотек.

Шляхи до шаблонів і словника в `config.json` лишаються тими, що були — їх
доведеться вказати заново, якщо на новому ПК вони в іншому місці. Це очікувано:
переносний тут **інтерпретатор**, а не ваші зразки.

## Перевірка на новому ПК

```bat
python\python.exe -c "import tkinter, ttkbootstrap, openpyxl, docx, win32com.client; print('усе на місці')"
start_generator.bat
```

## Як зібрати заново

Якщо теку `python\` загублено або треба інша версія:

```bat
rem 1. embeddable-збірка тієї самої версії й розрядності, що й у розробці
curl -o py.zip https://www.python.org/ftp/python/3.14.2/python-3.14.2-embed-amd64.zip
powershell -c "Expand-Archive py.zip python -Force"

rem 2. увімкнути site-packages: у python\python314._pth розкоментувати рядок
rem    #import site   ->   import site

rem 3. ДОКЛАСТИ tkinter — в embeddable його немає
rem    з повної інсталяції (напр. C:\Python314) скопіювати в python\:
rem      Lib\tkinter\      -> python\tkinter\
rem      tcl\              -> python\tcl\
rem      DLLs\_tkinter.pyd, DLLs\tcl86t.dll, DLLs\tk86t.dll, DLLs\zlib1.dll -> python\

rem 4. pip і залежності
curl -o python\get-pip.py https://bootstrap.pypa.io/get-pip.py
python\python.exe python\get-pip.py
python\python.exe -m pip install pywin32 openpyxl python-docx ttkbootstrap
del python\get-pip.py
```

### Дві пастки, на яких легко згоріти

1. **tkinter в embeddable-збірці відсутній.** Без кроку 3 програма падає на
   `import tkinter` — і жодне повідомлення про це не натякає.
2. **`zlib1.dll` теж треба.** Без нього `_tkinter.pyd` не вантажиться з
   `ImportError: DLL load failed`, хоча `tcl86t.dll` і `tk86t.dll` уже на місці:
   вони самі залежать від `zlib1.dll`, а в embeddable його немає.
