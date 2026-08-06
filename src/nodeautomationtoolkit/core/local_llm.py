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
    EMBEDDED = "Вбудована Qwen3 4B"
    GEMINI = "Google Gemini API"
    OPENAI = "OpenAI API"
    LM_STUDIO = "LM Studio"
    OLLAMA = "Ollama"
    CUSTOM = "OpenAI-сумісний"


DEFAULT_BASE_URLS = {
    LocalLlmProvider.EMBEDDED: "http://127.0.0.1:11439/v1/",
    LocalLlmProvider.GEMINI: "https://generativelanguage.googleapis.com/v1beta/openai/",
    LocalLlmProvider.OPENAI: "https://api.openai.com/v1/",
    LocalLlmProvider.LM_STUDIO: "http://127.0.0.1:1234/v1/",
    LocalLlmProvider.OLLAMA: "http://127.0.0.1:11434/v1/",
    LocalLlmProvider.CUSTOM: "http://127.0.0.1:1234/v1/",
}

PROVIDER_API_KEY_URLS = {
    LocalLlmProvider.GEMINI: "https://aistudio.google.com/app/apikey",
    LocalLlmProvider.OPENAI: "https://platform.openai.com/api-keys",
    LocalLlmProvider.LM_STUDIO: "https://lmstudio.ai/",
    LocalLlmProvider.OLLAMA: "https://ollama.com/",
    LocalLlmProvider.EMBEDDED: "",
    LocalLlmProvider.CUSTOM: "",
}

PROVIDER_PRESET_MODELS = {
    LocalLlmProvider.GEMINI: [
        "gemini-3.0-flash",
        "gemini-3.0-pro",
        "gemini-2.5-flash",
        "gemini-2.5-pro",
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-1.5-pro",
    ],
    LocalLlmProvider.OPENAI: [
        "gpt-5.6",
        "gpt-5.0",
        "gpt-4o",
        "gpt-4o-mini",
        "o3-mini",
        "gpt-4-turbo",
        "gpt-3.5-turbo",
    ],
    LocalLlmProvider.EMBEDDED: [
        "embedded-qwen3-4b",
    ],
    LocalLlmProvider.LM_STUDIO: [
        "qwen2.5-7b-instruct",
        "qwen2.5-coder-7b-instruct",
        "phi-4-mini-instruct",
        "llama-3.2-3b-instruct",
    ],
    LocalLlmProvider.OLLAMA: [
        # ── Оптимальні для 6 GB VRAM ───────────────────────────────────────────
        "qwen2.5:7b",            # ~5.5 GB VRAM · Найкраща для аналізу наказів та JSON
        "qwen2.5:3b",            # ~2 GB VRAM   · Дуже швидка, менш точна
        "phi4-mini",             # ~2.5 GB VRAM · Microsoft, відмінно для структурованих відповідей
        "llama3.2:3b",           # ~2 GB VRAM   · Хороший универсальний варіант
        # ── Для генерації Python-нод ──────────────────────────────────────────
        "qwen2.5-coder:7b",      # ~5.5 GB VRAM · Спеціалізований на коді
        "qwen2.5-coder:3b",      # ~2 GB VRAM   · Легша версія для Python-нод
        # ── Старіші (для сумісності) ──────────────────────────────────────────
        "mistral",
        "llama3.2",
    ],
    LocalLlmProvider.CUSTOM: [
        "default",
    ],
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


def load_llm_settings(provider: LocalLlmProvider | None = None) -> dict[str, str]:
    from PySide6.QtCore import QSettings

    settings = QSettings("DEADSUE.ART", "NodeAutomationToolkit")
    if provider is None:
        provider_name = settings.value("ai/last_provider", LocalLlmProvider.EMBEDDED.value)
        try:
            provider = LocalLlmProvider(str(provider_name))
        except ValueError:
            provider = LocalLlmProvider.EMBEDDED

    saved_url = settings.value(f"ai/base_url_{provider.value}", DEFAULT_BASE_URLS[provider])
    saved_model = settings.value(f"ai/model_{provider.value}", "")
    saved_key = settings.value(f"ai/api_key_{provider.value}", "")

    return {
        "provider": provider.value,
        "base_url": str(saved_url) if saved_url else DEFAULT_BASE_URLS[provider],
        "model": str(saved_model) if saved_model else "",
        "api_key": str(saved_key) if saved_key else "",
    }


def save_llm_settings(provider_value: str, base_url: str, model: str, api_key: str) -> None:
    from PySide6.QtCore import QSettings

    settings = QSettings("DEADSUE.ART", "NodeAutomationToolkit")
    settings.setValue("ai/last_provider", provider_value)
    settings.setValue(f"ai/base_url_{provider_value}", base_url)
    settings.setValue(f"ai/model_{provider_value}", model)
    if provider_value != LocalLlmProvider.EMBEDDED.value:
        settings.setValue(f"ai/api_key_{provider_value}", api_key)


SYSTEM_PROMPT = """Ти створюєш одну локальну Python-ноду для Node Automation Toolkit.

Поверни лише JSON за наданою схемою. Поле code має містити повний Python-модуль без
Markdown. Модуль імпортує декоратор так:

from nodeautomationtoolkit import node

і містить рівно одну функцію з @node(...). Використовуй type hints. Нода має бути
детермінованою, невеликою та зрозумілою. Не використовуй мережу, subprocess, shell,
eval, exec, compile, ctypes, приховані імпорти або автоматичне встановлення пакетів.
Не читай і не записуй файли. Не використовуй open, pathlib, os або shutil. Файлові
операції виконують лише перевірені вбудовані ноди; твоя нода перетворює лише значення,
передані через її входи.
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
            raise ValueError("Оберіть модель")

        data = self.generate_structured(
            system_prompt=SYSTEM_PROMPT,
            user_prompt=request_text.strip(),
            schema_name="node_draft",
            schema=NodeDraft.model_json_schema(),
        )
        try:
            return NodeDraft.model_validate(data)
        except (TypeError, ValueError) as error:
            raise LocalLlmError("Модель повернула некоректну відповідь") from error

    def generate_structured(
        self,
        *,
        system_prompt: str,
        user_prompt: str,
        schema_name: str,
        schema: dict[str, Any],
    ) -> dict[str, Any]:
        if not self.config.model.strip():
            raise ValueError("Оберіть модель")

        if self.config.provider == LocalLlmProvider.OPENAI:
            return self._generate_responses(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                schema_name=schema_name,
                schema=schema,
            )

        system_instruction = (
            f"{system_prompt}\n\nВАЖЛИВО: Дай відповідь ВИКЛЮЧНО у форматі валідного JSON за цією схемою:\n"
            f"{json.dumps(schema, ensure_ascii=False)}"
        )

        body: dict[str, Any] = {
            "model": self.config.model,
            "messages": [
                {"role": "system", "content": system_instruction},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.1,
            "stream": False,
        }

        if self.config.provider == LocalLlmProvider.GEMINI:
            body["response_format"] = {"type": "json_object"}
        else:
            body["response_format"] = {
                "type": "json_schema",
                "json_schema": {
                    "name": schema_name,
                    "strict": True,
                    "schema": schema,
                },
            }

        payload = self._request("chat/completions", body)
        try:
            content = payload["choices"][0]["message"]["content"]
            clean_content = content.strip()
            if clean_content.startswith("```"):
                lines = clean_content.splitlines()
                if lines[0].startswith("```"):
                    lines = lines[1:]
                if lines and lines[-1].startswith("```"):
                    lines = lines[:-1]
                clean_content = "\n".join(lines).strip()
            data = json.loads(clean_content)
        except (KeyError, IndexError, TypeError, json.JSONDecodeError) as error:
            raise LocalLlmError(f"LLM модель повернула некоректний JSON: {error}") from error
        if not isinstance(data, dict):
            raise LocalLlmError("Очікувався JSON-об'єкт")
        return data

    def _generate_responses(
        self,
        *,
        system_prompt: str,
        user_prompt: str,
        schema_name: str,
        schema: dict[str, Any],
    ) -> dict[str, Any]:
        body = {
            "model": self.config.model,
            "instructions": system_prompt,
            "input": user_prompt,
            "store": False,
            "text": {
                "format": {
                    "type": "json_schema",
                    "name": schema_name,
                    "strict": True,
                    "schema": schema,
                }
            },
        }
        payload = self._request("responses", body)
        text = payload.get("output_text")
        if not isinstance(text, str):
            parts = []
            for output in payload.get("output", []):
                for content in output.get("content", []):
                    if content.get("type") == "output_text" and isinstance(
                        content.get("text"), str
                    ):
                        parts.append(content["text"])
            text = "".join(parts)
        try:
            data = json.loads(text)
        except (TypeError, json.JSONDecodeError) as error:
            raise LocalLlmError("OpenAI API повернув некоректний JSON") from error
        if not isinstance(data, dict):
            raise LocalLlmError("Очікувався JSON-об'єкт")
        return data

    def fetch_available_models(self) -> list[str]:
        if self.config.provider == LocalLlmProvider.EMBEDDED:
            from .embedded_llm import MODEL_ALIAS
            return [MODEL_ALIAS]
        try:
            payload = self._request("models", method="GET")
            data = payload.get("data", [])
            models = []
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, dict) and "id" in item:
                        model_id = str(item["id"])
                        if model_id.startswith("models/"):
                            model_id = model_id[7:]
                        models.append(model_id)
                    elif isinstance(item, str):
                        models.append(item)
            if models:
                return sorted(list(dict.fromkeys(models)))
        except Exception:
            pass
        return PROVIDER_PRESET_MODELS.get(self.config.provider, ["default"])

    def _request(
        self,
        endpoint: str,
        body: dict[str, Any] | None = None,
        *,
        method: str = "POST",
    ) -> dict[str, Any]:
        base_url = self.config.normalized_base_url()
        if self.config.provider == LocalLlmProvider.EMBEDDED:
            from .embedded_llm import ensure_embedded_server

            base_url = ensure_embedded_server()
        url = urljoin(base_url, endpoint)
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
            raise LocalLlmError(f"Помилка LLM API ({error.code}): {details}") from error
        except URLError as error:
            if self.config.provider == LocalLlmProvider.EMBEDDED:
                raise LocalLlmError(
                    "Вбудована модель недоступна. Перевірте її через кнопку 'Локальна модель'."
                ) from error
            raise LocalLlmError(
                "LLM-сервер недоступний. Перевірте адресу, інтернет або локальний сервер."
            ) from error
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as error:
            raise LocalLlmError("LLM API повернув не JSON") from error
        if not isinstance(payload, dict):
            raise LocalLlmError("Некоректний формат відповіді LLM-сервера")
        return payload
