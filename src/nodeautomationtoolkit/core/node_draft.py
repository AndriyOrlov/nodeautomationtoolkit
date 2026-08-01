from __future__ import annotations

import ast
import re
from dataclasses import dataclass, field
from pathlib import Path

from pydantic import BaseModel, Field, field_validator


class NodeTestCase(BaseModel):
    arguments: dict = Field(default_factory=dict)
    expected: object | None = None


class NodeDraft(BaseModel):
    name: str = Field(min_length=1, max_length=80)
    category: str = Field(default="Згенеровані", min_length=1, max_length=60)
    description: str = Field(default="", max_length=500)
    function_name: str = Field(min_length=1, max_length=80)
    code: str = Field(min_length=1, max_length=25_000)
    requirements: list[str] = Field(default_factory=list, max_length=20)
    tests: list[NodeTestCase] = Field(default_factory=list, max_length=20)

    @field_validator("function_name")
    @classmethod
    def valid_function_name(cls, value: str) -> str:
        if not value.isidentifier():
            raise ValueError("function_name має бути коректним Python-ідентифікатором")
        return value


@dataclass(slots=True)
class NodeCodeReview:
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    permissions: set[str] = field(default_factory=set)

    @property
    def installable(self) -> bool:
        return not self.errors


SAFE_IMPORT_ROOTS = {
    "collections",
    "datetime",
    "decimal",
    "functools",
    "itertools",
    "json",
    "math",
    "nodeautomationtoolkit",
    "re",
    "statistics",
    "typing",
}

BLOCKED_IMPORT_ROOTS = {
    "ctypes",
    "ftplib",
    "http",
    "requests",
    "socket",
    "subprocess",
    "telnetlib",
    "urllib",
    "winreg",
}

BLOCKED_CALL_NAMES = {"__import__", "compile", "eval", "exec"}
BLOCKED_ATTRIBUTE_CALLS = {
    ("os", "popen"),
    ("os", "remove"),
    ("os", "removedirs"),
    ("os", "rmdir"),
    ("os", "system"),
    ("shutil", "rmtree"),
}


def review_node_code(code: str, *, allow_filesystem: bool = True) -> NodeCodeReview:
    review = NodeCodeReview()
    try:
        tree = ast.parse(code)
    except SyntaxError as error:
        review.errors.append(f"Синтаксична помилка: рядок {error.lineno}: {error.msg}")
        return review

    decorated_functions = []
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            if any(_decorator_name(item) == "node" for item in node.decorator_list):
                decorated_functions.append(node.name)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            roots = _import_roots(node)
            for root in roots:
                if root in BLOCKED_IMPORT_ROOTS:
                    review.errors.append(f"Заборонений імпорт: {root}")
                elif root not in SAFE_IMPORT_ROOTS:
                    review.errors.append(
                        f"Неперевірена бібліотека '{root}'. "
                        "Додайте її як окремий перевірений пакет нод."
                    )
        elif isinstance(node, ast.Call):
            if isinstance(node.func, ast.Name) and node.func.id in BLOCKED_CALL_NAMES:
                review.errors.append(f"Заборонений динамічний виклик: {node.func.id}")
            attribute = _attribute_call(node.func)
            if attribute in BLOCKED_ATTRIBUTE_CALLS:
                review.errors.append(f"Заборонена системна операція: {'.'.join(attribute)}")
            if isinstance(node.func, ast.Name) and node.func.id == "open":
                review.permissions.add("filesystem")
                if not allow_filesystem:
                    review.errors.append("AI-нода не може відкривати або записувати файли")

    if len(decorated_functions) != 1:
        review.errors.append("Модуль має містити рівно одну функцію з декоратором @node")
    if "filesystem" in review.permissions:
        review.warnings.append("Нода просить доступ до локальної файлової системи")
    return review


def install_node_draft(
    draft: NodeDraft,
    plugin_dir: Path,
    *,
    allow_filesystem: bool = True,
) -> Path:
    review = review_node_code(draft.code, allow_filesystem=allow_filesystem)
    if not review.installable:
        raise ValueError("Ноду не встановлено:\n" + "\n".join(review.errors))
    plugin_dir.mkdir(parents=True, exist_ok=True)
    safe_name = re.sub(r"[^a-zA-Z0-9_]+", "_", draft.function_name).strip("_").lower()
    if not safe_name:
        raise ValueError("Не вдалося створити безпечну назву файла")
    target = plugin_dir / f"generated_{safe_name}.py"
    if target.exists():
        raise FileExistsError(f"Нода вже існує: {target.name}")
    target.write_text(draft.code.rstrip() + "\n", encoding="utf-8")
    return target


def _decorator_name(node: ast.expr) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Call):
        return _decorator_name(node.func)
    if isinstance(node, ast.Attribute):
        return node.attr
    return ""


def _import_roots(node: ast.Import | ast.ImportFrom) -> set[str]:
    if isinstance(node, ast.ImportFrom):
        return {node.module.split(".", 1)[0]} if node.module else set()
    return {alias.name.split(".", 1)[0] for alias in node.names}


def _attribute_call(node: ast.expr) -> tuple[str, str] | None:
    if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name):
        return node.value.id, node.attr
    return None
