from __future__ import annotations

import json
from dataclasses import dataclass
from enum import StrEnum
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen

from .node_draft import NodeDraft


class LocalLlmProvider(StrEnum):
    LM_STUDIO = "LM Studio"
    OLLAMA = "Ollama"
    CUSTOM = "OpenAI-сумісний"


DEFAULT_BASE_URLS = {
    LocalLlmProvider.LM_STUDIO: "http://127.0.0.1:1234/v1/",
    LocalLlmProvider.OLLAMA: "http://127.0.0.1:11434/v1/",
    LocalLlmProvider.CUSTOM: "http://127.0.0.1:1234/v1/",
}


@dataclass(frozen=True, slots=True)
class LocalLlmConfig:
    provider: LocalLlmProvider = LocalLlmProvider.LM_STUDIO
    base_url: str = DEFAULT_BASE_URLS[LocalLlmProvider.LM_STUDIO]
    model: str = ""
    api_key: str = "local"
    timeout_seconds: int = 120

    def normalized_base_url(self) -> str:
        return self.base_url.rstrip("/") + "/"


SYSTEM_PROMPT = """Ти створюєш одну локальну Python-ноду для Node Automation Toolkit.

Поверни лише JSON за наданою схемою. Поле code має містити повний Python-модуль без
Markdown. Модуль імпортує декоратор так:

from nodeautomationtoolkit import node

і містить рівно одну функцію з @node(...). Використовуй type hints. Нода має бути
детермінованою, невеликою та зрозумілою. Не використовуй мережу, subprocess, shell,
eval, exec, compile, ctypes, приховані імпорти або автоматичне встановлення пакетів.
Не читай документи чи інші файли, якщо користувач прямо не описав це як вхід ноди.
Не додавай реальних службових даних. Для прикладів використовуй синтетичні значення.

Якщо задачу безпечно реалізувати неможливо, все одно поверни валідну чернетку з
поясненням обмеження у description та мінімальним безпечним кодом.
"""


class LocalLlmError(RuntimeError):
    pass


class LocalLlmClient:
    """Small dependency-free client for localhost OpenAI-compatible servers."""

    def __init__(self, config: LocalLlmConfig) -> None:
        self.config = config

    def list_models(self) -> list[str]:
        payload = self._request("models", method="GET")
        return [str(item["id"]) for item in payload.get("data", []) if item.get("id")]

    def generate_node(self, request_text: str) -> NodeDraft:
        if not request_text.strip():
            raise ValueError("Опишіть, яку ноду потрібно створити")
        if not self.config.model.strip():
            raise ValueError("Оберіть локальну модель")

        data = self.generate_structured(
            system_prompt=SYSTEM_PROMPT,
            user_prompt=request_text.strip(),
            schema_name="node_draft",
            schema=NodeDraft.model_json_schema(),
        )
        try:
            return NodeDraft.model_validate(data)
        except (TypeError, ValueError) as error:
            raise LocalLlmError("Локальна модель повернула некоректну відповідь") from error

    def generate_structured(
        self,
        *,
        system_prompt: str,
        user_prompt: str,
        schema_name: str,
        schema: dict[str, Any],
    ) -> dict[str, Any]:
        if not self.config.model.strip():
            raise ValueError("Оберіть локальну модель")
        body = {
            "model": self.config.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.1,
            "stream": False,
            "response_format": {
                "type": "json_schema",
                "json_schema": {
                    "name": schema_name,
                    "strict": True,
                    "schema": schema,
                },
            },
        }
        payload = self._request("chat/completions", body)
        try:
            content = payload["choices"][0]["message"]["content"]
            data = json.loads(content)
        except (KeyError, IndexError, TypeError, json.JSONDecodeError) as error:
            raise LocalLlmError("Локальна модель повернула некоректний JSON") from error
        if not isinstance(data, dict):
            raise LocalLlmError("Очікувався JSON-об'єкт")
        return data

    def _request(
        self,
        endpoint: str,
        body: dict[str, Any] | None = None,
        *,
        method: str = "POST",
    ) -> dict[str, Any]:
        url = urljoin(self.config.normalized_base_url(), endpoint)
        headers = {"Accept": "application/json"}
        data = None
        if body is not None:
            headers["Content-Type"] = "application/json"
            data = json.dumps(body).encode("utf-8")
        if self.config.api_key:
            headers["Authorization"] = f"Bearer {self.config.api_key}"
        request = Request(url, data=data, headers=headers, method=method)
        try:
            with urlopen(request, timeout=self.config.timeout_seconds) as response:  # noqa: S310
                raw = response.read().decode("utf-8")
        except HTTPError as error:
            details = error.read().decode("utf-8", errors="replace")
            raise LocalLlmError(
                f"Помилка локального LLM-сервера ({error.code}): {details}"
            ) from error
        except URLError as error:
            raise LocalLlmError(
                "Локальний LLM-сервер недоступний. Запустіть сервер у LM Studio або Ollama."
            ) from error
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as error:
            raise LocalLlmError("LLM-сервер повернув не JSON") from error
        if not isinstance(payload, dict):
            raise LocalLlmError("Некоректний формат відповіді LLM-сервера")
        return payload
