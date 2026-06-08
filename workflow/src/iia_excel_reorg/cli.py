"""Command-line workflow for reorganizing historical Excel workbooks."""

from __future__ import annotations
import argparse
import itertools
import re
import shutil
import sys
import time
from collections import defaultdict, deque
from pathlib import Path
from typing import Callable, TypeAlias
from .config import load_config
from .core.transformer import (
    DocumentIndex,
    GeographyIndex,
    MissingUnitCountryDocumentIndex,
    NonYearHeaderDocumentIndex,
    ProductIndex,
    UnitFootnoteDocumentIndex,
    transform_workbook,
)
from .paths import TRANSFORM_INPUT_DIR, TRANSFORM_LISTS_DIR, TRANSFORM_OUTPUT_DIR
from .utils.naming import sanitize_name
from .utils.text import derive_product_from_document, format_elapsed
from .xlsx_io import read_workbook

WorkbookEntry: TypeAlias = tuple[Path, Path]
WorkbookAction: TypeAlias = Callable[[WorkbookEntry], None]
TxtAction: TypeAlias = Callable[[], None]
_EXTRACTED_PAGES_RE = re.compile(
    r"^extracted_pages_(?P<year>\d{4})_\d{2}$",
    re.IGNORECASE,
)
_EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm"})
DEFAULT_INPUT_DIR = TRANSFORM_INPUT_DIR
DEFAULT_OUTPUT_DIR = TRANSFORM_OUTPUT_DIR
HEMISPHERE_INDEX_FILENAME = "unique_hemisphere_values.txt"
CONTINENT_INDEX_FILENAME = "unique_continent_values.txt"
COUNTRY_INDEX_FILENAME = "unique_country_values.txt"
PRODUCT_INDEX_FILENAME = "unique_product_values.txt"
FINAL_DOCUMENTS_INDEX_FILENAME = "final_docs.txt"
UNIT_FOOTNOTE_DOCUMENT_INDEX_FILENAME = "final_docs_with_unit_footnotes.txt"
MISSING_UNIT_COUNTRY_DOCUMENT_INDEX_FILENAME = (
    "final_docs_with_missing_country_units.txt"
)
NON_YEAR_HEADER_DOCUMENT_INDEX_FILENAME = (
    "final_docs_with_extra_non_year_columns.txt"
)
DUPLICATE_ORIGINAL_DOCUMENTS_FILENAME = "duplicated_original_documents.txt"
LISTS_DIR = TRANSFORM_LISTS_DIR


class DuplicateOriginalDocumentIndex:
    """Track original source documents that share the same filename."""

    def __init__(self) -> None:
        self._paths_by_name: dict[str, set[str]] = defaultdict(set)

    def add_document(self, path: Path, *, root: Path) -> None:
        """Record *path* relative to *root* for duplicate-name reporting."""
        if path.is_absolute():
            base_root = root if root.is_dir() else root.parent
            relative_path = path.relative_to(base_root)
        else:
            relative_path = path
        self._paths_by_name[path.name].add(relative_path.as_posix())

    def write_txt(self, path: str | Path) -> Path:
        """Write only duplicated original document names and their folders to *path*.

        Vectorized: the per-document ``append``/``extend`` loop is replaced by
        ``itertools.chain.from_iterable`` over a filtered generator expression.
        """
        output_path = Path(path)
        body = list(
            itertools.chain.from_iterable(
                [name, *(f"  {m}" for m in paths), ""]
                for name in sorted(self._paths_by_name)
                for paths in (sorted(self._paths_by_name[name]),)
                if len(paths) >= 2
            )
        )
        lines = ["[documents]", *body]
        output_path.write_text(
            "\n".join(lines + ([] if lines[-1] == "" else [""])),
            encoding="utf-8",
        )
        return output_path


def build_parser() -> argparse.ArgumentParser:
    """Build and return the CLI argument parser."""
    parser = argparse.ArgumentParser(
        description=(
            "Reorganize historical Excel workbooks into a standardized structure."
        ),
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=str(DEFAULT_INPUT_DIR),
        help=(
            "Excel workbook file or directory containing workbook files. "
            "Defaults to the 'data/transform/00_input/' folder in the current directory. "
            'Quote the path when it contains spaces: "data/transform/00_input".'
        ),
    )
    parser.add_argument(
        "output_dir",
        nargs="?",
        default=str(DEFAULT_OUTPUT_DIR),
        help=(
            "Directory where transformed workbooks will be written. "
            "Defaults to 'data/transform/01_output/' in the current directory."
        ),
    )
    parser.add_argument(
        "--config",
        help=(
            "Path to YAML configuration for categories, aliases, filters, "
            "and unit overrides."
        ),
    )
    return parser


def _compute_output_subdir(workbook_path: Path) -> Path:
    """Return the relative output subdirectory for *workbook_path*.

    Vectorized: the sequential ``enumerate`` scan for the ``extracted_pages_*``
    segment is replaced by a ``next()`` call over a generator expression.
    """
    parts = workbook_path.parts
    idx, match = next(
        (
            (idx, m)
            for idx, part in enumerate(parts)
            if (m := _EXTRACTED_PAGES_RE.match(part)) is not None
        ),
        (None, None),
    )
    if match is None:
        return Path(".")
    year = match.group("year")
    parent_dir = Path(f"iia_extracted_pages_{year}")
    intermediate = parts[idx + 1 : -1]
    if intermediate:
        child_dir = sanitize_name(f"iia_{intermediate[0]}_{year}")
        return parent_dir / child_dir
    if idx > 0:
        topic = parts[idx - 1]
        child_dir = sanitize_name(f"iia_{topic}_{year}")
        return parent_dir / child_dir
    return parent_dir


def _iter_workbooks(path: Path) -> list[Path]:
    """Return Excel workbooks under *path* using a non-recursive scan."""
    if path.is_file():
        return [path]
    return sorted(
        candidate
        for candidate in path.iterdir()
        if candidate.is_file() and candidate.suffix.lower() in _EXCEL_SUFFIXES
    )


def _iter_workbooks_structured(root: Path) -> list[WorkbookEntry]:
    """Walk *root* recursively and return ``(workbook_path, output_subdir)`` pairs."""
    workbook_paths = sorted(
        candidate
        for candidate in root.rglob("*")
        if candidate.is_file() and candidate.suffix.lower() in _EXCEL_SUFFIXES
    )
    return [
        (workbook_path, _compute_output_subdir(workbook_path))
        for workbook_path in workbook_paths
    ]


def _extract_sheet_names(path: Path) -> str:
    """Return a semicolon-delimited list of sheet names for *path*."""
    try:
        workbook = read_workbook(path)
        return "; ".join(sheet.name.title() for sheet in workbook.sheets)
    except Exception:
        return "Could not read sheets"


def _ensure_workspace(input_path: Path, output_root: Path) -> None:
    """Create or reset the input/output workspace directories."""
    if not input_path.exists() and input_path.suffix == "":
        input_path.mkdir(parents=True, exist_ok=True)
    if output_root.exists():
        shutil.rmtree(output_root)
    output_root.mkdir(parents=True, exist_ok=True)
    LISTS_DIR.mkdir(parents=True, exist_ok=True)


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
    """Run *action* on each item in *items* while updating a progress bar.

    Vectorized: the sequential ``for`` loop is replaced by a ``map``-driven
    side-effect pipeline consumed through ``collections.deque``.
    """
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


def _run_txt_progress(label: str, actions: list[tuple[str, TxtAction]]) -> None:
    """Run TXT generation actions while updating a dedicated progress bar.

    Vectorized: the sequential ``for`` loop is replaced by a ``map``-driven
    side-effect pipeline consumed through ``collections.deque``.
    """
    total = len(actions)
    sys.stdout.write(_render_progress_bar(label, 0, total))
    sys.stdout.flush()

    def _step(index_action: tuple[int, tuple[str, TxtAction]]) -> None:
        index, (_, action) = index_action
        action()
        sys.stdout.write("\r" + _render_progress_bar(label, index, total))
        sys.stdout.flush()

    deque(map(_step, enumerate(actions, start=1)), maxlen=0)
    print()


def main() -> None:
    """Entry point for the ``iia-excel-reorg`` command-line tool."""
    start_time = time.perf_counter()
    parser = build_parser()
    args = parser.parse_args()
    input_path = Path(args.input)
    output_root = Path(args.output_dir)
    _ensure_workspace(input_path, output_root)
    config = load_config(args.config)
    if input_path.is_file():
        workbook_entries: list[WorkbookEntry] = [(input_path, Path("."))]
    else:
        workbook_entries = _iter_workbooks_structured(input_path)
        if not workbook_entries:
            workbook_entries = [
                (path, Path(".")) for path in _iter_workbooks(input_path)
            ]
    if not workbook_entries:
        print(f"No Excel workbooks found in: {input_path}")
        print(
            "Created workspace folders if needed. Add source Excel files there "
            "and run again."
        )
        return
    geography_index = GeographyIndex()
    product_index = ProductIndex()
    document_index = DocumentIndex()
    duplicate_original_document_index = DuplicateOriginalDocumentIndex()
    unit_footnote_document_index = UnitFootnoteDocumentIndex()
    missing_unit_country_document_index = MissingUnitCountryDocumentIndex()
    non_year_header_document_index = NonYearHeaderDocumentIndex()

    def prepare_output(entry: WorkbookEntry) -> None:
        """Create the output subdirectory for *entry* if needed."""
        _, output_subdir = entry
        (output_root / output_subdir).mkdir(parents=True, exist_ok=True)

    def transform_entry(entry: WorkbookEntry) -> None:
        """Transform a single workbook and write it into the output tree."""
        workbook_path, output_subdir = entry
        output_dir = output_root / output_subdir
        sheet_names = _extract_sheet_names(workbook_path)
        duplicate_original_document_index.add_document(workbook_path, root=input_path)
        output_name = (
            f"{sanitize_name(config.canonical_name_for_document(workbook_path))}.xlsx"
        )
        output_path = output_dir / output_name
        entry_missing_unit_country_document_index = MissingUnitCountryDocumentIndex()
        transform_workbook(
            workbook_path,
            output_path,
            config=config,
            geography_index=geography_index,
            unit_footnote_document_index=unit_footnote_document_index,
            missing_unit_country_document_index=(
                entry_missing_unit_country_document_index
            ),
            non_year_header_document_index=non_year_header_document_index,
        )
        if entry_missing_unit_country_document_index.documents:
            missing_unit_country_document_index.add_document_sheet_names(
                workbook_path.name,
                sheet_names,
            )
        document_index.add_document(output_path.name)
        product_index.add_product(derive_product_from_document(output_path.name))

    _run_progress("Reorganizing folders", workbook_entries, prepare_output)
    _run_progress("Reorganizing excels", workbook_entries, transform_entry)
    _run_txt_progress(
        "Generating txt lists",
        [
            (
                HEMISPHERE_INDEX_FILENAME,
                lambda: geography_index.write_dimension_txt(
                    LISTS_DIR / HEMISPHERE_INDEX_FILENAME,
                    label="hemispheres",
                ),
            ),
            (
                CONTINENT_INDEX_FILENAME,
                lambda: geography_index.write_dimension_txt(
                    LISTS_DIR / CONTINENT_INDEX_FILENAME,
                    label="continents",
                ),
            ),
            (
                COUNTRY_INDEX_FILENAME,
                lambda: geography_index.write_dimension_txt(
                    LISTS_DIR / COUNTRY_INDEX_FILENAME,
                    label="countries",
                ),
            ),
            (
                PRODUCT_INDEX_FILENAME,
                lambda: product_index.write_txt(LISTS_DIR / PRODUCT_INDEX_FILENAME),
            ),
            (
                FINAL_DOCUMENTS_INDEX_FILENAME,
                lambda: document_index.write_txt(LISTS_DIR / FINAL_DOCUMENTS_INDEX_FILENAME),
            ),
            (
                UNIT_FOOTNOTE_DOCUMENT_INDEX_FILENAME,
                lambda: unit_footnote_document_index.write_txt(
                    LISTS_DIR / UNIT_FOOTNOTE_DOCUMENT_INDEX_FILENAME
                ),
            ),
            (
                MISSING_UNIT_COUNTRY_DOCUMENT_INDEX_FILENAME,
                lambda: missing_unit_country_document_index.write_txt(
                    LISTS_DIR / MISSING_UNIT_COUNTRY_DOCUMENT_INDEX_FILENAME
                ),
            ),
            (
                NON_YEAR_HEADER_DOCUMENT_INDEX_FILENAME,
                lambda: non_year_header_document_index.write_txt(
                    LISTS_DIR / NON_YEAR_HEADER_DOCUMENT_INDEX_FILENAME
                ),
            ),
            (
                DUPLICATE_ORIGINAL_DOCUMENTS_FILENAME,
                lambda: duplicate_original_document_index.write_txt(
                    LISTS_DIR / DUPLICATE_ORIGINAL_DOCUMENTS_FILENAME
                ),
            ),
        ],
    )
    elapsed = time.perf_counter() - start_time
    print(f"Done in {format_elapsed(elapsed)}")


if __name__ == "__main__":
    main()
