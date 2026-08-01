from __future__ import annotations

import importlib
import importlib.util
import sys
from collections.abc import Iterable
from pathlib import Path
from types import ModuleType

from .definition import NodeDefinition


class NodeRegistry:
    def __init__(self) -> None:
        self._definitions: dict[str, NodeDefinition] = {}
        self.errors: list[str] = []

    def register(self, definition: NodeDefinition) -> None:
        if definition.type_id in self._definitions:
            raise ValueError(f"Дублікат type_id ноди: {definition.type_id}")
        self._definitions[definition.type_id] = definition

    def get(self, type_id: str) -> NodeDefinition:
        try:
            return self._definitions[type_id]
        except KeyError as error:
            raise KeyError(f"Нода не встановлена: {type_id}") from error

    def all(self) -> list[NodeDefinition]:
        return sorted(self._definitions.values(), key=lambda item: (item.category, item.name))

    def categories(self) -> dict[str, list[NodeDefinition]]:
        result: dict[str, list[NodeDefinition]] = {}
        for definition in self.all():
            result.setdefault(definition.category, []).append(definition)
        return result

    def reload(self, plugin_dir: Path | None = None) -> None:
        self._definitions.clear()
        self.errors.clear()
        self._load_builtin_modules()
        if plugin_dir is not None:
            self._load_plugin_directory(plugin_dir)

    def _load_builtin_modules(self) -> None:
        modules = [
            "nodeautomationtoolkit.builtin_nodes.files",
            "nodeautomationtoolkit.builtin_nodes.text",
            "nodeautomationtoolkit.builtin_nodes.text_analysis",
            "nodeautomationtoolkit.builtin_nodes.recipient_mapping",
            "nodeautomationtoolkit.builtin_nodes.order_ai",
            "nodeautomationtoolkit.builtin_nodes.logic",
            "nodeautomationtoolkit.builtin_nodes.windows_dialogs",
            "nodeautomationtoolkit.builtin_nodes.word",
            "nodeautomationtoolkit.builtin_nodes.word_batch",
            "nodeautomationtoolkit.builtin_nodes.output",
        ]
        for module_name in modules:
            module = importlib.import_module(module_name)
            self._register_module(module)

    def _load_plugin_directory(self, plugin_dir: Path) -> None:
        plugin_dir.mkdir(parents=True, exist_ok=True)
        for file_path in sorted(plugin_dir.glob("*.py")):
            if file_path.name.startswith("_"):
                continue
            module_name = f"nat_user_plugin_{file_path.stem}"
            try:
                spec = importlib.util.spec_from_file_location(module_name, file_path)
                if spec is None or spec.loader is None:
                    raise ImportError("Не вдалося створити завантажувач")
                module = importlib.util.module_from_spec(spec)
                sys.modules[module_name] = module
                spec.loader.exec_module(module)
                self._register_module(module)
            except Exception as error:  # noqa: BLE001 - plugin boundary
                self.errors.append(f"{file_path.name}: {error}")

    def _register_module(self, module: ModuleType) -> None:
        for definition in self._definitions_in_module(module):
            try:
                self.register(definition)
            except ValueError as error:
                self.errors.append(f"{module.__name__}: {error}")

    @staticmethod
    def _definitions_in_module(module: ModuleType) -> Iterable[NodeDefinition]:
        for value in vars(module).values():
            definition = getattr(value, "__nat_node_definition__", None)
            if isinstance(definition, NodeDefinition):
                yield definition
