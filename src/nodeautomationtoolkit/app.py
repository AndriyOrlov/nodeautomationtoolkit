import encodings.ascii  # noqa: F401
import encodings.cp1251  # noqa: F401
import encodings.idna  # noqa: F401
import encodings.latin_1  # noqa: F401
import encodings.mbcs  # noqa: F401
import encodings.utf_8  # noqa: F401
import json
import os
import sys
from pathlib import Path

from PySide6.QtCore import QStandardPaths
from PySide6.QtWidgets import QApplication

from .core.executor import GraphExecutor
from .core.models import GraphModel
from .core.registry import NodeRegistry
from .ui.main_window import MainWindow


def _run_cli_scenario(path_str: str, registry: NodeRegistry) -> int:
    scenario_path = Path(path_str).expanduser()
    if not scenario_path.is_file():
        print(f"❌ Помилка: файл сценарію не знайдено: {scenario_path}", file=sys.stderr)
        return 1

    print(f"🚀 Виконання сценарію: {scenario_path.name}...")
    try:
        raw_data = json.loads(scenario_path.read_text(encoding="utf-8"))
        graph = GraphModel.from_dict(raw_data)
        executor = GraphExecutor(registry)
        exec_result = executor.execute(graph)
        print("✅ Сценарій успішно виконано!")
        for node_id, values in exec_result.values.items():
            if isinstance(values, dict) and "summary" in values:
                print(f"  • [{node_id}]: {values['summary']}")
        return 0
    except Exception as err:
        print(f"❌ Помилка виконання сценарію: {err}", file=sys.stderr)
        return 1


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Node Automation Toolkit")
    app.setOrganizationName("DEADSUE.ART")

    registry = NodeRegistry()
    app_data = Path(
        QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation)
    )
    os.environ["NAT_APP_DATA_DIR"] = str(app_data)
    plugin_dir = app_data / "plugins" / "nodes"
    registry.reload(plugin_dir)

    args = sys.argv[1:]
    if "--run" in args or "-r" in args:
        flag = "--run" if "--run" in args else "-r"
        idx = args.index(flag)
        if idx + 1 < len(args):
            return _run_cli_scenario(args[idx + 1], registry)

    window = MainWindow(registry=registry, plugin_dir=plugin_dir)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
