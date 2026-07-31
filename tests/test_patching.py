from __future__ import annotations

import json
import zipfile

import pytest

from nodeautomationtoolkit.core.patching import PATCH_FORMAT, install_patch


def test_installs_offline_patch_and_activates_version(tmp_path):
    archive = tmp_path / "update.natpatch.zip"
    manifest = {"format": PATCH_FORMAT, "version": "0.3.1"}
    with zipfile.ZipFile(archive, "w") as package:
        package.writestr("patch.json", json.dumps(manifest))
        package.writestr("nodeautomationtoolkit/__init__.py", '__version__ = "0.3.1"\n')

    result = install_patch(archive, tmp_path / "app-data")

    assert result.version == "0.3.1"
    assert (result.target_dir / "nodeautomationtoolkit" / "__init__.py").is_file()
    active_file = tmp_path / "app-data" / "patches" / "active_patch.txt"
    assert active_file.read_text(encoding="utf-8") == "0.3.1"


def test_rejects_patch_path_traversal(tmp_path):
    archive = tmp_path / "unsafe.zip"
    manifest = {"format": PATCH_FORMAT, "version": "0.3.1"}
    with zipfile.ZipFile(archive, "w") as package:
        package.writestr("patch.json", json.dumps(manifest))
        package.writestr("../outside.py", "bad")

    with pytest.raises(ValueError, match="Небезпечний шлях"):
        install_patch(archive, tmp_path / "app-data")
