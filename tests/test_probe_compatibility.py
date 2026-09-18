from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PLUGIN_DIR = ROOT / "microsoft365"


def test_package_imports_under_catalog_probe_module_name() -> None:
    probe = """
import importlib.util
import sys
from pathlib import Path

sys.path[:] = [
    entry for entry in sys.path
    if not (entry and (Path(entry) / "microsoft365").is_dir())
]
plugin_dir = sys.argv[1]
module_name = "hermes_validate_probe_plugin"
spec = importlib.util.spec_from_file_location(
    module_name,
    plugin_dir + "/__init__.py",
    submodule_search_locations=[plugin_dir],
)
module = importlib.util.module_from_spec(spec)
module.__path__ = [plugin_dir]
sys.modules[spec.name] = module
spec.loader.exec_module(module)
assert callable(module.register)
"""
    result = subprocess.run(
        [sys.executable, "-c", probe, str(PLUGIN_DIR)],
        cwd=Path(__file__).resolve().parent,
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr or result.stdout
