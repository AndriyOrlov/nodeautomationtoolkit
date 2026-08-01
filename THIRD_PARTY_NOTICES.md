# Third-party notices

## Qwen3 4B

- Model: `Qwen/Qwen3-4B`
- GGUF quantization: `ggml-org/Qwen3-4B-GGUF`, `Q4_K_M`
- License: Apache License 2.0
- Source: https://huggingface.co/Qwen/Qwen3-4B
- GGUF: https://huggingface.co/ggml-org/Qwen3-4B-GGUF

The model is not stored in this repository. The user explicitly installs it from
Hugging Face using the in-app local model manager.

## llama.cpp

- Project: `ggml-org/llama.cpp`
- License: MIT
- Source: https://github.com/ggml-org/llama.cpp

The application downloads the pinned Windows Vulkan runtime only after the user
presses the install button. Inference binds to `127.0.0.1`, disables the Web UI and
built-in tools, and forces llama.cpp offline mode.
