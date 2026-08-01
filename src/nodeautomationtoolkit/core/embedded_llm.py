from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import threading
import time
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

MODEL_NAME = "Qwen3 4B Q4_K_M"
MODEL_ALIAS = "embedded-qwen3-4b"
MODEL_FILENAME = "Qwen3-4B-Q4_K_M.gguf"
MODEL_URL = (
    "https://huggingface.co/ggml-org/Qwen3-4B-GGUF/resolve/main/Qwen3-4B-Q4_K_M.gguf?download=true"
)
MODEL_SIZE = 2_497_280_640
MODEL_SHA256 = "ab27b9bfa375a178d6cba48f3ad892b94b7739659dcc7aae8058ce0ffed6b328"

LLAMA_BUILD = "b10217"
RUNTIME_FILENAME = f"llama-{LLAMA_BUILD}-bin-win-vulkan-x64.zip"
RUNTIME_URL = (
    f"https://github.com/ggml-org/llama.cpp/releases/download/{LLAMA_BUILD}/{RUNTIME_FILENAME}"
)
RUNTIME_SIZE = 34_089_052
RUNTIME_SHA256 = "957320cb0bca241ec8249d16d4b137f0d27df08ce7740345f36313fa21d77b2b"

SERVER_PORT = 11_439
SERVER_API_KEY = "nat-local"
ProgressCallback = Callable[[str, int, int], None]


def default_embedded_llm_dir() -> Path:
    configured = os.environ.get("NAT_APP_DATA_DIR", "").strip()
    if configured:
        return Path(configured) / "embedded-llm"
    if os.name == "nt":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
        return root / "DEADSUE.ART" / "Node Automation Toolkit" / "embedded-llm"
    return Path.home() / ".local" / "share" / "nodeautomationtoolkit" / "embedded-llm"


@dataclass(frozen=True, slots=True)
class EmbeddedLlmStatus:
    runtime_installed: bool
    model_installed: bool
    runtime_path: str
    model_path: str

    @property
    def ready(self) -> bool:
        return self.runtime_installed and self.model_installed


class EmbeddedLlmInstaller:
    def __init__(self, base_dir: Path | None = None) -> None:
        self.base_dir = (base_dir or default_embedded_llm_dir()).resolve()
        self.runtime_dir = self.base_dir / "runtime"
        self.model_dir = self.base_dir / "models"
        self.model_path = self.model_dir / MODEL_FILENAME

    def server_path(self) -> Path:
        direct = self.runtime_dir / "llama-server.exe"
        if direct.is_file():
            return direct
        matches = list(self.runtime_dir.rglob("llama-server.exe"))
        return matches[0] if matches else direct

    def status(self) -> EmbeddedLlmStatus:
        server = self.server_path()
        return EmbeddedLlmStatus(
            runtime_installed=server.is_file(),
            model_installed=self.model_path.is_file()
            and self.model_path.stat().st_size == MODEL_SIZE,
            runtime_path=str(server),
            model_path=str(self.model_path),
        )

    def install(self, progress: ProgressCallback | None = None) -> EmbeddedLlmStatus:
        if os.name != "nt":
            raise RuntimeError("Вбудована модель зараз підтримує Windows 10/11 x64")
        self.base_dir.mkdir(parents=True, exist_ok=True)
        self.runtime_dir.mkdir(parents=True, exist_ok=True)
        self.model_dir.mkdir(parents=True, exist_ok=True)
        if not self.server_path().is_file():
            archive = self.base_dir / RUNTIME_FILENAME
            self._download(
                RUNTIME_URL,
                archive,
                RUNTIME_SIZE,
                RUNTIME_SHA256,
                "Рушій llama.cpp",
                progress,
            )
            self._extract_runtime(archive)
        if not self.status().model_installed:
            self._download(
                MODEL_URL,
                self.model_path,
                MODEL_SIZE,
                MODEL_SHA256,
                f"Модель {MODEL_NAME}",
                progress,
            )
        status = self.status()
        if not status.ready:
            raise RuntimeError("Локальну модель встановлено не повністю")
        return status

    def _download(
        self,
        url: str,
        target: Path,
        expected_size: int,
        expected_sha256: str,
        label: str,
        progress: ProgressCallback | None,
    ) -> None:
        if target.is_file() and target.stat().st_size == expected_size:
            if progress:
                progress(label, expected_size, expected_size)
            return
        partial = target.with_suffix(target.suffix + ".part")
        downloaded = partial.stat().st_size if partial.is_file() else 0
        headers = {"User-Agent": "NodeAutomationToolkit/0.4"}
        if downloaded:
            headers["Range"] = f"bytes={downloaded}-"
        request = Request(url, headers=headers)
        try:
            response = urlopen(request, timeout=60)  # noqa: S310
        except (HTTPError, URLError) as error:
            raise RuntimeError(f"Не вдалося завантажити {label}: {error}") from error
        mode = "ab" if downloaded and getattr(response, "status", 200) == 206 else "wb"
        if mode == "wb":
            downloaded = 0
        with response, partial.open(mode) as stream:
            while chunk := response.read(1024 * 1024):
                stream.write(chunk)
                downloaded += len(chunk)
                if progress:
                    progress(label, downloaded, expected_size)
        if partial.stat().st_size != expected_size:
            raise RuntimeError(
                f"{label} завантажено не повністю: {partial.stat().st_size} із {expected_size} байт"
            )
        digest = hashlib.sha256()
        with partial.open("rb") as stream:
            while chunk := stream.read(8 * 1024 * 1024):
                digest.update(chunk)
        if digest.hexdigest() != expected_sha256:
            partial.unlink(missing_ok=True)
            raise RuntimeError(f"Контрольна сума {label} не збігається")
        os.replace(partial, target)

    def _extract_runtime(self, archive: Path) -> None:
        with tempfile.TemporaryDirectory(prefix="nat-llama-") as temporary:
            staging = Path(temporary)
            with zipfile.ZipFile(archive) as package:
                root = staging.resolve()
                for member in package.infolist():
                    destination = (staging / member.filename).resolve()
                    if root not in destination.parents and destination != root:
                        raise RuntimeError("Небезпечний шлях у ZIP рушія")
                package.extractall(staging)
            shutil.copytree(staging, self.runtime_dir, dirs_exist_ok=True)
        if not self.server_path().is_file():
            raise RuntimeError("У пакеті llama.cpp немає llama-server.exe")


_process: subprocess.Popen | None = None
_process_lock = threading.Lock()


def embedded_base_url() -> str:
    return f"http://127.0.0.1:{SERVER_PORT}/v1/"


def _server_healthy() -> bool:
    try:
        request = Request(
            f"http://127.0.0.1:{SERVER_PORT}/health",
            headers={"Authorization": f"Bearer {SERVER_API_KEY}"},
        )
        with urlopen(request, timeout=2) as response:  # noqa: S310
            payload = json.loads(response.read().decode("utf-8"))
        return payload.get("status") == "ok"
    except (OSError, ValueError, URLError):
        return False


def ensure_embedded_server(base_dir: Path | None = None) -> str:
    global _process
    if _server_healthy():
        return embedded_base_url()
    with _process_lock:
        if _server_healthy():
            return embedded_base_url()
        installer = EmbeddedLlmInstaller(base_dir)
        status = installer.status()
        if not status.ready:
            raise RuntimeError(
                "Локальна модель не встановлена. Натисніть 'Локальна модель' → 'Встановити'."
            )
        command = [
            status.runtime_path,
            "--model",
            status.model_path,
            "--alias",
            MODEL_ALIAS,
            "--host",
            "127.0.0.1",
            "--port",
            str(SERVER_PORT),
            "--ctx-size",
            "16384",
            "--n-gpu-layers",
            "all",
            "--jinja",
            "--reasoning",
            "off",
            "--offline",
            "--no-webui",
            "--api-key",
            SERVER_API_KEY,
            "--log-disable",
        ]
        creationflags = 0x08000000 if os.name == "nt" else 0
        _process = subprocess.Popen(
            command,
            cwd=str(Path(status.runtime_path).parent),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags,
        )
        deadline = time.monotonic() + 240
        while time.monotonic() < deadline:
            if _process.poll() is not None:
                raise RuntimeError("Локальна модель не запустилась")
            if _server_healthy():
                return embedded_base_url()
            time.sleep(0.5)
        _process.terminate()
        raise RuntimeError("Локальна модель не встигла запуститися")
