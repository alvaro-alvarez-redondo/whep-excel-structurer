from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent
SRC_ROOT = REPO_ROOT / "workflow" / "src"
DEFAULT_CONFIG = REPO_ROOT / "workflow" / "config" / "example.units.yml"


_REQUIRED_PACKAGES = ["iia_excel_reorg", "pandas", "numpy"]


def _ensure_dependencies() -> None:
    """Install the package and all its dependencies when running from source."""
    missing = [pkg for pkg in _REQUIRED_PACKAGES if importlib.util.find_spec(pkg) is None]
    if not missing:
        return

    workflow_dir = REPO_ROOT / "workflow"
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "-e", str(workflow_dir)]
    )


def main() -> None:
    """Run the workbook reorganization workflow from VS Code's Run button.

    When no command-line arguments are provided, this wrapper automatically uses
    the repository's example configuration file so the project can be launched
    directly as a single script from the repository root.
    """
    sys.path.insert(0, str(SRC_ROOT))
    _ensure_dependencies()

    from iia_excel_reorg.cli import main as cli_main

    if len(sys.argv) == 1 and DEFAULT_CONFIG.exists():
        sys.argv.extend(["--config", str(DEFAULT_CONFIG)])

    cli_main()


if __name__ == "__main__":
    main()
