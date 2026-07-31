import json

from nodeautomationtoolkit.core.local_llm import LocalLlmClient, LocalLlmConfig


def test_parses_structured_node_response(monkeypatch):
    response = {
        "choices": [
            {
                "message": {
                    "content": json.dumps(
                        {
                            "name": "Подвоїти",
                            "category": "Тест",
                            "description": "",
                            "function_name": "double",
                            "code": "from nodeautomationtoolkit import node\n\n"
                            "@node(name='Подвоїти')\n"
                            "def double(value: int) -> int:\n    return value * 2\n",
                            "requirements": [],
                            "tests": [],
                        },
                        ensure_ascii=False,
                    )
                }
            }
        ]
    }

    client = LocalLlmClient(LocalLlmConfig(model="local-model"))
    monkeypatch.setattr(client, "_request", lambda endpoint, body: response)

    draft = client.generate_node("Створи ноду подвоєння")

    assert draft.function_name == "double"

