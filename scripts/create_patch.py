from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a small offline NAT source patch")
    parser.add_argument("--version", required=True)
    parser.add_argument("--output", type=Path, required=True)
    options = parser.parse_args()

    root = Path(__file__).resolve().parents[1]
    source_package = root / "src" / "nodeautomationtoolkit"
    manifest = json.loads((root / "patch.json").read_text(encoding="utf-8"))
    manifest["version"] = options.version
    options.output.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(options.output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "patch.json",
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        )
        for path in sorted(source_package.rglob("*")):
            if not path.is_file() or "__pycache__" in path.parts:
                continue
            archive.write(path, path.relative_to(root / "src").as_posix())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
