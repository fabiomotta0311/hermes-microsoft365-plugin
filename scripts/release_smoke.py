#!/usr/bin/env python3
"""Validate a built wheel in a clean, dependency-free virtual environment."""
from __future__ import annotations

import argparse
import importlib.metadata
import re
import subprocess
import sys
import tempfile
import venv
from pathlib import Path

DISTRIBUTION = "hermes-microsoft365-plugin"
VERSION_RE = re.compile(r"^Version: (?P<version>.+)$", re.MULTILINE)
MANIFEST_RE = re.compile(r"^version:\s*(?P<version>[^\s#]+)", re.MULTILINE)
TRAIN = {
    "msgraph-sdk": "1.64.0",
    "msgraph-core": "1.5.1",
    "azure-identity": "1.26.0",
    "microsoft-kiota-abstractions": "1.14.2",
    "microsoft-kiota-authentication-azure": "1.12.3",
    "microsoft-kiota-http": "1.14.2",
    "microsoft-kiota-serialization-form": "1.12.3",
    "microsoft-kiota-serialization-json": "1.14.2",
    "microsoft-kiota-serialization-multipart": "1.12.3",
    "microsoft-kiota-serialization-text": "1.12.3",
}


def run(*args: str, cwd: str | None = None) -> None:
    subprocess.run(args, check=True, cwd=cwd)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("wheel", type=Path)
    args = parser.parse_args()
    wheel = args.wheel.resolve()
    if wheel.suffix != ".whl" or not wheel.is_file():
        raise SystemExit(f"not a wheel: {wheel}")

    with tempfile.TemporaryDirectory(prefix="hermes-m365-wheel-smoke-") as temp:
        env = Path(temp) / "venv"
        venv.EnvBuilder(with_pip=True, clear=True).create(env)
        python = env / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")
        run(str(python), "-m", "pip", "install", "--disable-pip-version-check", str(wheel))
        run(str(python), "-m", "pip", "check")
        probe = """
import importlib.metadata as metadata
from pathlib import Path
import sys
import microsoft365

assert Path(microsoft365.__file__).resolve().is_relative_to(Path(sys.prefix).resolve())
version = metadata.version("hermes-microsoft365-plugin")
assert version == __VERSION__
assert callable(microsoft365.register)
manifest = Path(microsoft365.__file__).with_name("plugin.yaml").read_text(encoding="utf-8")
assert "version: " + version in manifest
for package, expected in __TRAIN__.items():
    actual = metadata.version(package)
    assert actual == expected, (package, actual, expected)
    print(f"{package}=={actual}")
print(f"installed {metadata.distribution('hermes-microsoft365-plugin').metadata['Name']} {version}")
""".replace("__VERSION__", repr(_wheel_version(wheel))).replace("__TRAIN__", repr(TRAIN))
        run(str(python), "-I", "-c", probe, cwd=temp)
    return 0


def _wheel_version(wheel: Path) -> str:
    import zipfile

    with zipfile.ZipFile(wheel) as archive:
        metadata = next(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.endswith(".dist-info/METADATA")
        )
        manifest = next(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name == "microsoft365/plugin.yaml"
        )
    metadata_match = VERSION_RE.search(metadata)
    manifest_match = MANIFEST_RE.search(manifest)
    if not metadata_match or not manifest_match or metadata_match["version"] != manifest_match["version"]:
        raise SystemExit("wheel metadata and plugin manifest versions do not agree")
    return metadata_match["version"]


if __name__ == "__main__":
    raise SystemExit(main())
