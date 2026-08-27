# -*- coding: utf-8 -*-
"""Пошук імен, які функція ЧИТАЄ, але ніде не звʼязує — майбутній `NameError`.

Навіщо власний обхід AST: `py_compile` таких помилок не бачить (синтаксис
правильний), а pyflakes у проєкт не поставити — правило 1.1 забороняє мережу.

Це найдорожчий клас дефектів у цьому проєкті: код компілюється, тести на
модулях проходять, і все падає вже у користувача посеред пакета. Так було з
`_slash_to_lines`, `executor_start_from_bookmark`, `clean_redundant_blanks`,
і так само — з `preview_delay`, коли присвоєння потрапило не в ту функцію.
"""
from __future__ import annotations

import ast
import builtins
import importlib
import io


def _star_names(module_name: str) -> set[str]:
    """Публічні імена модуля, який імпортовано зірочкою."""
    if not module_name:
        return set()
    try:
        module = importlib.import_module(module_name)
    except Exception:
        return set()
    exported = getattr(module, "__all__", None)
    if exported is not None:
        return set(exported)
    return {name for name in dir(module) if not name.startswith("_")}


def _bound_names(node) -> set[str]:
    """Імена, звʼязані В МЕЖАХ цієї області видимості (без вкладених функцій)."""
    names: set[str] = set()

    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda)):
        a = node.args
        for arg in [*a.posonlyargs, *a.args, *a.kwonlyargs]:
            names.add(arg.arg)
        if a.vararg:
            names.add(a.vararg.arg)
        if a.kwarg:
            names.add(a.kwarg.arg)

    def handle(n):
        if isinstance(n, ast.Name) and isinstance(n.ctx, (ast.Store, ast.Del)):
            names.add(n.id)
        elif isinstance(n, (ast.Import, ast.ImportFrom)):
            for alias in n.names:
                if alias.name == "*":
                    # `from x import *` — імена відомі лише самому модулю,
                    # тому питаємо його напряму.
                    names.update(_star_names(n.module))
                    continue
                names.add((alias.asname or alias.name).split(".")[0])
        elif isinstance(n, (ast.Global, ast.Nonlocal)):
            names.update(n.names)
        elif isinstance(n, ast.ExceptHandler) and n.name:
            names.add(n.name)

    def walk(n):
        for child in ast.iter_child_nodes(n):
            # У вкладену функцію/клас не заходимо: там своя область.
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                names.add(child.name)
                continue
            if isinstance(child, ast.Lambda):
                continue
            handle(child)
            walk(child)

    walk(node)
    return names


def _loaded_names(node):
    """Імена, які ця область ЧИТАЄ (без вкладених функцій).

    Вкладені області пропускаються НА РІВНІ ДИТИНИ, тому обхід починається з
    самого вузла: якщо піти по `node.body`, то вкладений `def` як окремий
    оператор буде розкритий, і його локальні імена приїдуть сюди чужими.
    """
    out = []

    def walk(n):
        for child in ast.iter_child_nodes(n):
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef, ast.Lambda)):
                continue
            if isinstance(child, ast.Name) and isinstance(child.ctx, ast.Load):
                out.append((child.id, child.lineno))
            walk(child)

    walk(node)
    return out


def find_undefined_names(path):
    tree = ast.parse(io.open(path, encoding="utf-8").read(), filename=str(path))
    module_names = _bound_names(tree) | set(dir(builtins)) | {"__file__", "__name__", "__doc__"}

    problems = []

    def visit(node, enclosing: set[str], where: str):
        # Клас не створює області для вкладених функцій — його імена
        # доступні лише через self, тож у ланцюг не додаємо.
        if isinstance(node, ast.ClassDef):
            for child in node.body:
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    visit(child, enclosing, f"{where}.{child.name}")
                elif isinstance(child, ast.ClassDef):
                    visit(child, enclosing, f"{where}.{child.name}")
            return

        own = _bound_names(node)
        visible = enclosing | own
        for name, lineno in _loaded_names(node):
            if name not in visible:
                problems.append((lineno, where, name))
        # Вкладені функції — по черзі, з нашими іменами у ланцюгу.
        def descend(n):
            for child in ast.iter_child_nodes(n):
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    visit(child, visible, f"{where}.{child.name}")
                elif isinstance(child, ast.ClassDef):
                    visit(child, visible, f"{where}.{child.name}")
                elif isinstance(child, ast.Lambda):
                    visit(child, visible, f"{where}.<lambda>")
                else:
                    descend(child)
        descend(node)

    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            visit(node, module_names, node.name)
        elif isinstance(node, ast.ClassDef):
            visit(node, module_names, node.name)
    return problems
