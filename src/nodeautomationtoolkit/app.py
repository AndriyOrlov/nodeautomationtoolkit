from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication

from .core.registry import NodeRegistry
from .ui.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Node Automation Toolkit")
    app.setOrganizationName("DEADSUE.ART")

    registry = NodeRegistry()
    plugin_dir = Path.cwd() / "plugins" / "nodes"
    registry.reload(plugin_dir)

    window = MainWindow(registry=registry, plugin_dir=plugin_dir)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

