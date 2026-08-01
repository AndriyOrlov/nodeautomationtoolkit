import json

from nodeautomationtoolkit.core.local_llm import (
    LocalLlmClient,
    LocalLlmConfig,
    LocalLlmProvider,
)


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


def test_openai_uses_responses_without_tools_or_storage(monkeypatch):
    captured = {}
    response = {"output_text": json.dumps({"value": "готово"}, ensure_ascii=False)}
    client = LocalLlmClient(
        LocalLlmConfig(
            provider=LocalLlmProvider.OPENAI,
            base_url="https://api.openai.com/v1/",
            model="gpt-5.6",
            api_key="secret",
        )
    )

    def fake_request(endpoint, body):
        captured.update({"endpoint": endpoint, "body": body})
        return response

    monkeypatch.setattr(client, "_request", fake_request)
    result = client.generate_structured(
        system_prompt="Тільки план",
        user_prompt="Створи граф",
        schema_name="result",
        schema={
            "type": "object",
            "properties": {"value": {"type": "string"}},
            "required": ["value"],
            "additionalProperties": False,
        },
    )

    assert result == {"value": "готово"}
    assert captured["endpoint"] == "responses"
    assert captured["body"]["store"] is False
    assert "tools" not in captured["body"]
    assert "files" not in json.dumps(captured["body"])
