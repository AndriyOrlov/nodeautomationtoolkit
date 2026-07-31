from pathlib import Path


def test_packaged_entrypoint_uses_absolute_import():
    entrypoint = Path("src/nodeautomationtoolkit/__main__.py").read_text(encoding="utf-8")

    assert "from nodeautomationtoolkit.app import main" in entrypoint
    assert "from .app import main" not in entrypoint
