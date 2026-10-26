"""Ресурс версії для EXE повідомлень (вкладка «Подробиці» у властивостях файла).

EXE без опису й версії антивіруси та SmartScreen оцінюють суворіше, тож
`build_messages_qt.bat` вшиває цей ресурс у файл через `--version-file`.

Використання:
    python scripts/messages_exe/make_version_info.py <файл-результат>
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from bump_version import read_version  # noqa: E402

EXE_NAME = "PovidomlenniaKadrovi.exe"
DESCRIPTION = "Повідомлення про прийняття кадрових рішень"
PRODUCT = "Node Automation Toolkit — повідомлення"
AUTHOR = "Andrii Orlov"


def version_info_text(version: str) -> str:
    numbers = tuple(int(part) for part in version.split(".")) + (0,)
    strings = {
        "CompanyName": AUTHOR,
        "FileDescription": DESCRIPTION,
        "FileVersion": version,
        "InternalName": EXE_NAME.removesuffix(".exe"),
        "LegalCopyright": f"© {AUTHOR}",
        "OriginalFilename": EXE_NAME,
        "ProductName": PRODUCT,
        "ProductVersion": version,
    }
    entries = ",\n".join(
        f"          StringStruct({key!r}, {value!r})" for key, value in strings.items()
    )
    return f"""VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={numbers},
    prodvers={numbers},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '042204B0',
        [
{entries}
        ])
    ]),
    VarFileInfo([VarStruct('Translation', [0x0422, 1200])])
  ]
)
"""


def main(argv: list[str]) -> int:
    if len(argv) != 1:
        print(__doc__)
        return 2
    target = Path(argv[0])
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(version_info_text(read_version()), encoding="utf-8")
    print(target)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
