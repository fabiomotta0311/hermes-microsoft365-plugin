#!/usr/bin/env python3
"""Emit a deterministic CycloneDX SBOM from the committed uv.lock."""
from __future__ import annotations

import argparse
import hashlib
import json
import tomllib
from pathlib import Path


def component(package: dict) -> dict:
    name = package["name"]
    version = package["version"]
    hashes = []
    for artifact in [package.get("sdist"), *(package.get("wheels") or [])]:
        if artifact and artifact.get("hash", "").startswith("sha256:"):
            hashes.append({"alg": "SHA-256", "content": artifact["hash"].split(":", 1)[1]})
    result = {
        "type": "library",
        "bom-ref": f"pkg:pypi/{name}@{version}",
        "name": name,
        "version": version,
        "purl": f"pkg:pypi/{name}@{version}",
    }
    if hashes:
        result["hashes"] = hashes
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--lock", type=Path, default=Path("uv.lock"))
    parser.add_argument("--project", type=Path, default=Path("pyproject.toml"))
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    lock_bytes = args.lock.read_bytes()
    lock = tomllib.loads(lock_bytes.decode("utf-8"))
    project = tomllib.loads(args.project.read_text(encoding="utf-8"))["project"]
    packages = sorted(lock.get("package", []), key=lambda item: (item["name"], item["version"]))
    serial = hashlib.sha256(lock_bytes).hexdigest()[:32]
    serial = f"{serial[:8]}-{serial[8:12]}-{serial[12:16]}-{serial[16:20]}-{serial[20:]}"
    bom = {
        "bomFormat": "CycloneDX",
        "specVersion": "1.5",
        "serialNumber": "urn:uuid:" + serial,
        "version": 1,
        "metadata": {
            "tools": [{"vendor": "Hermes Microsoft 365 Plugin", "name": "generate_sbom.py", "version": "1"}],
            "component": {"type": "application", "name": project["name"], "version": project["version"]},
        },
        "components": [component(package) for package in packages],
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(bom, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(f"wrote {args.output} with {len(packages)} locked components")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
