"""Central directory layout constants and helpers for all pipelines."""

from __future__ import annotations

from pathlib import Path

# Root data directory (sibling to the repository root)
DATA_DIR = Path("data")

# ---------------------------------------------------------------------------
# Pre-processing pipeline (human review preparation)
# ---------------------------------------------------------------------------
PREPROCESS_DIR = DATA_DIR / "preprocess"
PREPROCESS_INPUT_DIR = PREPROCESS_DIR / "00_input"
PREPROCESS_OUTPUT_DIR = PREPROCESS_DIR / "01_output"
PREPROCESS_COUNTRY_LABEL_PATTERNS_PATH = (
    DATA_DIR / "country_label_patterns.xlsx"
)

# ---------------------------------------------------------------------------
# Main transformation pipeline
# ---------------------------------------------------------------------------
TRANSFORM_DIR = DATA_DIR / "transform"
TRANSFORM_INPUT_DIR = TRANSFORM_DIR / "00_input"
TRANSFORM_OUTPUT_DIR = TRANSFORM_DIR / "01_output"
TRANSFORM_LISTS_DIR = TRANSFORM_DIR / "lists"

# ---------------------------------------------------------------------------
# Shared resources
# ---------------------------------------------------------------------------
FOOTNOTE_MAPPING_PATH = DATA_DIR / "footnote_mapping.xlsx"

# Legacy aliases for backward-compatible CLI defaults
DEFAULT_LEGACY_INPUT_DIR = Path("data/raw_inputs")
DEFAULT_LEGACY_OUTPUT_DIR = Path("data/10-raw_imports")
DEFAULT_LEGACY_LISTS_DIR = Path("data/lists")


def resolve_project_root(caller_path: Path | None = None) -> Path:
    """Return the project root directory (repository root).

    When *caller_path* is provided, walks up three levels from it to reach the
    repo root.  Otherwise returns the current working directory.
    """
    if caller_path is not None:
        return caller_path.resolve().parents[2]
    return Path.cwd()


def ensure_preprocess_dirs(project_root: Path | None = None) -> tuple[Path, Path]:
    """Create pre-processing source and prepared directories if missing.

    Returns ``(source_dir, prepared_dir)``.
    """
    root = project_root or Path.cwd()
    source = root / PREPROCESS_SOURCE_DIR
    prepared = root / PREPROCESS_PREPARED_DIR
    source.mkdir(parents=True, exist_ok=True)
    prepared.mkdir(parents=True, exist_ok=True)
    return source, prepared


def ensure_transform_dirs(
    project_root: Path | None = None,
) -> tuple[Path, Path, Path]:
    """Create transformation input, output, and lists directories if missing.

    Returns ``(input_dir, output_dir, lists_dir)``.
    """
    root = project_root or Path.cwd()
    inp = root / TRANSFORM_INPUT_DIR
    out = root / TRANSFORM_OUTPUT_DIR
    lists_ = root / TRANSFORM_LISTS_DIR
    inp.mkdir(parents=True, exist_ok=True)
    out.mkdir(parents=True, exist_ok=True)
    lists_.mkdir(parents=True, exist_ok=True)
    return inp, out, lists_
