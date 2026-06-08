"""Independent pipeline for harmonizing footnotes in transformed workbooks."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .io.xlsx import SheetData, WorkbookData, read_workbook, write_workbook
from .paths import FOOTNOTE_MAPPING_PATH, TRANSFORM_OUTPUT_DIR

PIPELINE_NAME = "Footnote Harmonization Pipeline"
DEFAULT_INPUT_DIR = TRANSFORM_OUTPUT_DIR
DEFAULT_TEMPLATE_PATH = FOOTNOTE_MAPPING_PATH
_EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm"})
_ORIGINAL_FOOTNOTE_HEADER = "original footnote"
_CLEANED_FOOTNOTE_HEADER = "cleaned footnote"


def _normalize_header(value: str | int | float | None) -> str:
    """Return a case-folded header label."""
    if value is None:
        return ""
    return str(value).strip().casefold()


def _split_footnotes(value: str | int | float | None) -> list[str]:
    """Split a cell value into normalized footnote tokens."""
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    return [token.strip() for token in text.split(";") if token.strip()]


def _join_footnotes(values: list[str]) -> str:
    """Join footnote tokens while preserving order and removing duplicates."""
    seen: set[str] = set()
    ordered = [value for value in values if not (value in seen or seen.add(value))]
    return "; ".join(ordered)


def _iter_workbooks(root: Path) -> list[Path]:
    """Return all Excel workbooks under *root* recursively."""
    return sorted(
        candidate
        for candidate in root.rglob("*")
        if candidate.is_file() and candidate.suffix.lower() in _EXCEL_SUFFIXES
    )


def _find_footnotes_column(sheet) -> int | None:
    """Return the column index of the ``footnotes`` header, if present."""
    return next(
        (
            column
            for column in range(1, sheet.max_column + 1)
            if _normalize_header(sheet.get_cell(1, column).value) == "footnotes"
        ),
        None,
    )


def collect_unique_footnotes(input_dir: str | Path = DEFAULT_INPUT_DIR) -> list[str]:
    """Collect sorted unique footnotes from every workbook in *input_dir*."""
    root = Path(input_dir)
    unique: set[str] = set()
    for workbook_path in _iter_workbooks(root):
        workbook = read_workbook(workbook_path)
        for sheet in workbook.sheets:
            footnote_col = _find_footnotes_column(sheet)
            if footnote_col is None:
                continue
            for row in range(2, sheet.max_row + 1):
                unique.update(_split_footnotes(sheet.get_cell(row, footnote_col).value))
    return sorted(unique)


def generate_mapping_template(
    input_dir: str | Path = DEFAULT_INPUT_DIR,
    template_path: str | Path = DEFAULT_TEMPLATE_PATH,
) -> Path:
    """Write an Excel mapping template populated with unique original footnotes."""
    output_path = Path(template_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    original_footnotes = collect_unique_footnotes(input_dir)
    sheet_rows = [
        [_ORIGINAL_FOOTNOTE_HEADER.title(), _CLEANED_FOOTNOTE_HEADER.title()],
        *[[footnote, ""] for footnote in original_footnotes],
    ]
    workbook = WorkbookData(sheets=[])
    mapping_sheet = SheetData(name="footnote_mapping")
    for row_index, values in enumerate(sheet_rows, start=1):
        mapping_sheet.set_row(row_index, values)
    workbook.sheets.append(mapping_sheet)
    return write_workbook(output_path, workbook)


def load_mapping_template(path: str | Path) -> dict[str, str]:
    """Load ``original -> cleaned`` mappings from an Excel mapping template."""
    workbook = read_workbook(path)
    if not workbook.sheets:
        return {}
    sheet = workbook.sheets[0]
    header_positions = {
        _normalize_header(sheet.get_cell(1, column).value): column
        for column in range(1, sheet.max_column + 1)
    }
    original_col = header_positions.get(_ORIGINAL_FOOTNOTE_HEADER)
    cleaned_col = header_positions.get(_CLEANED_FOOTNOTE_HEADER)
    if original_col is None or cleaned_col is None:
        raise ValueError(
            "Mapping template must include 'Original Footnote' and "
            "'Cleaned Footnote' columns."
        )
    mapping: dict[str, str] = {}
    for row in range(2, sheet.max_row + 1):
        original = str(sheet.get_cell(row, original_col).value or "").strip()
        cleaned = str(sheet.get_cell(row, cleaned_col).value or "").strip()
        if original and cleaned:
            mapping[original] = cleaned
    return mapping


def _rewrite_workbook_footnotes(workbook_path: Path, mapping: dict[str, str]) -> bool:
    """Apply *mapping* to one workbook and return whether any cell changed."""
    workbook = read_workbook(workbook_path)
    changed = False
    for sheet in workbook.sheets:
        footnote_col = _find_footnotes_column(sheet)
        if footnote_col is None:
            continue
        for row in range(2, sheet.max_row + 1):
            cell = sheet.get_cell(row, footnote_col)
            parts = _split_footnotes(cell.value)
            if not parts:
                continue
            remapped = [mapping.get(part, part) for part in parts]
            cleaned_value = _join_footnotes(remapped)
            if cleaned_value != str(cell.value or "").strip():
                sheet.set_cell(row, footnote_col, cleaned_value, fill_rgb=cell.fill_rgb)
                changed = True
    if changed:
        write_workbook(workbook_path, workbook)
    return changed


def apply_mapping_in_place(
    input_dir: str | Path = DEFAULT_INPUT_DIR,
    mapping_template_path: str | Path = DEFAULT_TEMPLATE_PATH,
) -> list[Path]:
    """Rewrite footnotes in all workbooks under *input_dir* using mapping template."""
    mapping = load_mapping_template(mapping_template_path)
    root = Path(input_dir)
    changed_paths = [
        workbook_path
        for workbook_path in _iter_workbooks(root)
        if _rewrite_workbook_footnotes(workbook_path, mapping)
    ]
    return changed_paths


def build_parser() -> argparse.ArgumentParser:
    """Build CLI parser for the independent footnote harmonization pipeline."""
    parser = argparse.ArgumentParser(
        description=(
            f"{PIPELINE_NAME}: generate a footnote mapping template and/or apply "
            "cleaned footnotes in place."
        ),
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    generate_parser = subparsers.add_parser(
        "generate-template",
        help="Scan transformed outputs and create a footnote mapping template workbook.",
    )
    generate_parser.add_argument("input_dir", nargs="?", default=str(DEFAULT_INPUT_DIR))
    generate_parser.add_argument(
        "template_path",
        nargs="?",
        default=str(DEFAULT_TEMPLATE_PATH),
    )

    apply_parser = subparsers.add_parser(
        "apply-mapping",
        help="Apply an edited mapping template to all source workbooks in place.",
    )
    apply_parser.add_argument("input_dir", nargs="?", default=str(DEFAULT_INPUT_DIR))
    apply_parser.add_argument(
        "template_path",
        nargs="?",
        default=str(DEFAULT_TEMPLATE_PATH),
    )
    return parser


def main() -> None:
    """CLI entrypoint for the independent footnote harmonization pipeline."""
    parser = build_parser()
    cli_args = sys.argv[1:]
    if not cli_args:
        if DEFAULT_TEMPLATE_PATH.exists():
            args = parser.parse_args(
                ["apply-mapping", str(DEFAULT_INPUT_DIR), str(DEFAULT_TEMPLATE_PATH)]
            )
        else:
            args = parser.parse_args(
                ["generate-template", str(DEFAULT_INPUT_DIR), str(DEFAULT_TEMPLATE_PATH)]
            )
    else:
        args = parser.parse_args(cli_args)
    if args.command == "generate-template":
        output_path = generate_mapping_template(args.input_dir, args.template_path)
        print(f"Generated mapping template: {output_path}")
        return
    if args.command == "apply-mapping":
        changed_paths = apply_mapping_in_place(args.input_dir, args.template_path)
        print(f"Updated {len(changed_paths)} workbook(s) in place.")
        return
    raise ValueError(f"Unknown command: {args.command}")
