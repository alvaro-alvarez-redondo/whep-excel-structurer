from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent
SRC_ROOT = REPO_ROOT / "workflow" / "src"


_REQUIRED_PACKAGES = ["iia_excel_reorg", "openpyxl", "pandas", "numpy"]


def _ensure_dependencies() -> None:
    """Install the package and all its dependencies when running from source."""
    missing = [pkg for pkg in _REQUIRED_PACKAGES if importlib.util.find_spec(pkg) is None]
    if not missing:
        # Also ensure fast Excel engine is available
        if importlib.util.find_spec("calamine") is None:
            workflow_dir = REPO_ROOT / "workflow"
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "-e", str(workflow_dir) + "[fast]"]
            )
        return

    workflow_dir = REPO_ROOT / "workflow"
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "-e", str(workflow_dir) + "[fast]"]
    )


def main() -> None:
    """Run the pre-processing pipeline from VS Code's Run button.

    This wrapper enables one-file execution without requiring package
    installation first.
    """
    sys.path.insert(0, str(SRC_ROOT))
    _ensure_dependencies()

    from iia_excel_reorg.preprocess_pipeline import main as preprocess_main

    preprocess_main()


if __name__ == "__main__":
    main()
