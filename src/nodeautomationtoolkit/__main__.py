from __future__ import annotations

import importlib.abc
import importlib.util
import os
import sys
from pathlib import Path


class _InstalledPatchFinder(importlib.abc.MetaPathFinder):
    def __init__(self, patch_root: Path) -> None:
        self.package_root = patch_root / "nodeautomationtoolkit"

    def find_spec(self, fullname: str, path=None, target=None):
        del path, target
        if fullname == "nodeautomationtoolkit":
            candidate = self.package_root / "__init__.py"
            package_locations = [str(self.package_root)]
        elif fullname.startswith("nodeautomationtoolkit."):
            relative = fullname.split(".")[1:]
            package_dir = self.package_root.joinpath(*relative)
            if (package_dir / "__init__.py").is_file():
                candidate = package_dir / "__init__.py"
                package_locations = [str(package_dir)]
            else:
                candidate = package_dir.with_suffix(".py")
                package_locations = None
        else:
            return None
        if not candidate.is_file():
            return None
        return importlib.util.spec_from_file_location(
            fullname,
            candidate,
            submodule_search_locations=package_locations,
        )


def _patch_locations() -> list[Path]:
    if override := os.environ.get("NODEAUTOMATIONTOOLKIT_APP_DATA"):
        return [Path(override)]
    if roaming := os.environ.get("APPDATA"):
        root = Path(roaming)
        return [
            root / "Node Automation Toolkit",
            root / "DEADSUE.ART" / "Node Automation Toolkit",
        ]
    data_home = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))
    return [data_home / "Node Automation Toolkit"]


def _activate_installed_patch() -> None:
    for app_data in _patch_locations():
        active_file = app_data / "patches" / "active_patch.txt"
        try:
            version = active_file.read_text(encoding="utf-8").strip()
        except OSError:
            continue
        patch_root = app_data / "patches" / "versions" / version
        if (patch_root / "nodeautomationtoolkit" / "__init__.py").is_file():
            sys.meta_path.insert(0, _InstalledPatchFinder(patch_root))
            return


_activate_installed_patch()

from nodeautomationtoolkit.app import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
