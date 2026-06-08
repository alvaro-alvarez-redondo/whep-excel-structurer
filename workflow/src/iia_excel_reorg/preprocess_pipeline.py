"""Pre-processing pipeline: prepare raw Excel files for human review."""

from __future__ import annotations

import argparse
import shutil
import sys
import time
import traceback
from collections import deque
from pathlib import Path
from typing import Callable, TypeAlias

from .core.preprocessor import process_workbook
from .paths import (
    PREPROCESS_COUNTRY_LABEL_PATTERNS_PATH,
    PREPROCESS_INPUT_DIR,
    PREPROCESS_OUTPUT_DIR,
)
from .utils.text import format_elapsed

WorkbookEntry: TypeAlias = tuple[Path, Path]
WorkbookAction: TypeAlias = Callable[[WorkbookEntry], None]
_EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm"})


def build_parser() -> argparse.ArgumentParser:
    """Build and return the CLI argument parser."""
    parser = argparse.ArgumentParser(
        description=(
            "Pre-process raw Excel workbooks for human review. "
            "Mirrors the source folder structure into the prepared output."
        ),
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=str(PREPROCESS_INPUT_DIR),
        help=(
            "Excel workbook file or directory containing workbook files. "
            "Defaults to the 'data/preprocess/00_input/' folder."
        ),
    )
    parser.add_argument(
        "output_dir",
        nargs="?",
        default=str(PREPROCESS_OUTPUT_DIR),
        help=(
            "Directory where prepared workbooks will be written. "
            "Defaults to 'data/preprocess/01_output/'."
        ),
    )
    parser.add_argument(
        "--country-label-patterns",
        default=str(PREPROCESS_COUNTRY_LABEL_PATTERNS_PATH),
        help=(
            "Excel workbook with letter_dictionary, country_patterns, and "
            "description sheets. Defaults to data/country_label_patterns.xlsx; "
            "a preset workbook is generated there when missing."
        ),
    )
    return parser


def _iter_workbooks_recursive(root: Path) -> list[Path]:
    """Walk *root* recursively and return all Excel workbook paths."""
    return sorted(
        candidate
        for candidate in root.rglob("*")
        if candidate.is_file() and candidate.suffix.lower() in _EXCEL_SUFFIXES
    )


def _compute_relative_entries(
    input_root: Path, output_root: Path
) -> list[WorkbookEntry]:
    """Return ``(source_path, target_path)`` pairs preserving folder structure."""
    workbook_paths = _iter_workbooks_recursive(input_root)
    return [
        (
            workbook_path,
            output_root / workbook_path.relative_to(input_root),
        )
        for workbook_path in workbook_paths
    ]


def _ensure_workspace(input_path: Path, output_root: Path) -> None:
    """Create or reset the output workspace directory."""
    if not input_path.exists() and input_path.suffix == "":
        input_path.mkdir(parents=True, exist_ok=True)
    if output_root.exists():
        shutil.rmtree(output_root)
    output_root.mkdir(parents=True, exist_ok=True)


def _render_progress_bar(
    label: str,
    current: int,
    total: int,
    width: int = 24,
) -> str:
    """Return a single-line progress bar string."""
    normalized_total = max(total, 1)
    completed = min(width, int(width * current / normalized_total))
    percent = int(100 * current / normalized_total)
    bar = "█" * completed + "·" * (width - completed)
    return f"{label:<21} │{bar}│ {percent:>3}% ({current}/{normalized_total})"


def _run_progress(
    label: str,
    items: list[WorkbookEntry],
    action: WorkbookAction,
) -> None:
    """Run *action* on each item in *items* while updating a progress bar."""
    total = len(items)
    sys.stdout.write(_render_progress_bar(label, 0, total))
    sys.stdout.flush()

    def _step(index_item: tuple[int, WorkbookEntry]) -> None:
        index, item = index_item
        action(item)
        sys.stdout.write("\r" + _render_progress_bar(label, index, total))
        sys.stdout.flush()

    deque(map(_step, enumerate(items, start=1)), maxlen=0)
    print()


def _prepare_entry(entry: WorkbookEntry) -> None:
    """Create output subdirectory for *entry* if needed."""
    _, target_path = entry
    target_path.parent.mkdir(parents=True, exist_ok=True)


def _process_entry(entry: WorkbookEntry, country_label_patterns_path: Path) -> None:
    """Process a single workbook and write it to the target path."""
    source_path, target_path = entry
    try:
        process_workbook(
            source_path,
            target_path,
            country_label_patterns_path=country_label_patterns_path,
        )
    except Exception:
        print(f"\nError processing {source_path}:")
        traceback.print_exc()


def main() -> None:
    """Entry point for the ``iia-prepare`` command-line tool."""
    start_time = time.perf_counter()
    parser = build_parser()
    args = parser.parse_args()
    input_path = Path(args.input)
    output_root = Path(args.output_dir)
    country_label_patterns_path = Path(args.country_label_patterns)
    _ensure_workspace(input_path, output_root)

    if input_path.is_file():
        workbook_entries: list[WorkbookEntry] = [
            (input_path, output_root / input_path.name)
        ]
    else:
        workbook_entries = _compute_relative_entries(input_path, output_root)

    if not workbook_entries:
        print(f"No Excel workbooks found in: {input_path}")
        print(
            "Created workspace folders if needed. Add source Excel files there "
            "and run again."
        )
        return

    _run_progress("Preparing folders", workbook_entries, _prepare_entry)
    _run_progress(
        "Preparing excels",
        workbook_entries,
        lambda entry: _process_entry(entry, country_label_patterns_path),
    )

    elapsed = time.perf_counter() - start_time
    print(f"Done in {format_elapsed(elapsed)}")


if __name__ == "__main__":
    main()
