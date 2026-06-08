from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent
SRC_ROOT = REPO_ROOT / "workflow" / "src"


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
    """Run the independent footnote harmonization pipeline from repo root.

    This wrapper enables one-file execution (for example via VS Code Run button)
    without requiring package installation first.
    """
    sys.path.insert(0, str(SRC_ROOT))
    _ensure_dependencies()

    from iia_excel_reorg.footnote_pipeline import main as footnote_main

    footnote_main()


if __name__ == "__main__":
    main()

