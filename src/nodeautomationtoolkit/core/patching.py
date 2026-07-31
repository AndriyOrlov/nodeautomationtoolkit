from __future__ import annotations

import json
import os
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

PATCH_FORMAT = "nodeautomationtoolkit-patch-v1"


@dataclass(frozen=True, slots=True)
class InstalledPatch:
    version: str
    target_dir: Path
    restart_required: bool = True


def install_patch(archive: Path, app_data_dir: Path) -> InstalledPatch:
    archive = Path(archive)
    if not archive.is_file():
        raise FileNotFoundError(f"Патч не знайдено: {archive}")

    with zipfile.ZipFile(archive) as package:
        names = package.namelist()
        _validate_members(names)
        if "patch.json" not in names:
            raise ValueError("У патчі немає patch.json")
        manifest = json.loads(package.read("patch.json").decode("utf-8"))
        version = _validate_manifest(manifest)
        if not any(name.startswith("nodeautomationtoolkit/") for name in names):
            raise ValueError("У патчі немає пакета nodeautomationtoolkit")

        versions_dir = app_data_dir / "patches" / "versions"
        target_dir = versions_dir / version
        versions_dir.mkdir(parents=True, exist_ok=True)
        if not target_dir.exists():
            temporary = Path(tempfile.mkdtemp(prefix=f".{version}-", dir=versions_dir))
            try:
                package.extractall(temporary)
                if not (temporary / "nodeautomationtoolkit" / "__init__.py").is_file():
                    raise ValueError("Патч не містить повного Python-пакета")
                os.replace(temporary, target_dir)
            except Exception:
                shutil.rmtree(temporary, ignore_errors=True)
                raise

    active_file = app_data_dir / "patches" / "active_patch.txt"
    active_file.parent.mkdir(parents=True, exist_ok=True)
    temporary_active = active_file.with_suffix(".tmp")
    temporary_active.write_text(version, encoding="utf-8")
    os.replace(temporary_active, active_file)
    return InstalledPatch(version=version, target_dir=target_dir)


def _validate_manifest(manifest: object) -> str:
    if not isinstance(manifest, dict) or manifest.get("format") != PATCH_FORMAT:
        raise ValueError("Невідомий формат патча")
    version = str(manifest.get("version", "")).strip()
    if not version or any(character not in "0123456789.-_" for character in version):
        raise ValueError("Некоректна версія патча")
    return version


def _validate_members(names: list[str]) -> None:
    for name in names:
        path = PurePosixPath(name)
        if path.is_absolute() or ".." in path.parts or "\\" in name:
            raise ValueError(f"Небезпечний шлях у патчі: {name}")
        if path.parts and path.parts[0] not in {"patch.json", "nodeautomationtoolkit"}:
            raise ValueError(f"Зайвий файл у патчі: {name}")
