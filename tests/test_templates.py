from nodeautomationtoolkit.core.executor import GraphExecutor
from nodeautomationtoolkit.core.registry import NodeRegistry
from nodeautomationtoolkit.core.templates import build_word_smoke_graph


def test_word_smoke_graph_is_executable(monkeypatch):
    registry = NodeRegistry()
    registry.reload()
    captured = []
    monkeypatch.setattr(
        registry.get("builtin.windows.open_file"),
        "function",
        lambda **_kwargs: "order.docx",
    )
    monkeypatch.setattr(
        registry.get("builtin.word.read_docx"),
        "function",
        lambda path: {
            "document": None,
            "file_name": path,
            "paragraphs": None,
            "text": "Тестовий текст наказу",
        },
    )
    monkeypatch.setattr(
        registry.get("builtin.output.show_result"),
        "function",
        lambda value, title="": captured.append((title, value)) or value,
    )

    result = GraphExecutor(registry).execute(build_word_smoke_graph())

    assert "show-word-text" in result.order
    assert captured == [("Текст Word-документа", "Тестовий текст наказу")]
