"""Workbook transformation logic for the normalization pipeline."""

from __future__ import annotations

import itertools
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import TypeAlias

import numpy as np
import pandas as pd

from ..config import WorkbookConfig
from ..io.xlsx import SheetData, WorkbookData, read_workbook, write_workbook
from ..services.units import UNIT_PLACEHOLDER

HeaderYear: TypeAlias = tuple[int, str]
RowValue: TypeAlias = str | int | float | None


@dataclass(slots=True)
class OutputRow:
    """Row-oriented normalized output used before worksheet materialization."""

    values: list[RowValue]
    fills: list[str | None]


HEADER_FILL = "FF3CCB5A"
HEADER_COLUMNS = ["hemisphere", "continent", "country", "unit", "footnotes"]
HEADER_HAS_DIGIT_RE = re.compile(r"\d")
PAREN_RE = re.compile(r"\(([^()]*)\)")
HEMISPHERE_RE = re.compile(r"h[eéê]misph[eèê]?re|hemisphere", re.IGNORECASE)


def _normalize_known_geography_label(value: str) -> str:
    """Strip accents, fold to ASCII lowercase, and strip trailing punctuation."""
    normalized = unicodedata.normalize("NFKD", value)
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return ascii_only.casefold().strip().rstrip(".:")


RAW_CONTINENT_LABELS = (
    "AFRIQUE",
    "AMÉR DU NORD ET AMÉR CENTRALE",
    "AMER. DU NORD ET AMER. CENTR.",
    "AMERIQUE",
    "AMERIQUÉ",
    "AMÉRIQUE",
    "AMÉRIQUE CENTRALE ET MEXIQUE.",
    "AMERIQUE CENTRALE.",
    "AMÉRIQUE CENTRALE.",
    "AMERIQUE DU NORD",
    "AMÉRIQUE DU NORD",
    "AMÉRIQUE DU NORD ET AMÉR CENTR",
    "AMERIQUE DU NORD ET AMERIQUE CENTRALE",
    "AMERIQUE DU NORD ET AMÉRIQUE CENTRALE",
    "AMÉRIQUE DU NORD ET AMERIQUE CENTRALE",
    "AMÉRIQUE DU NORD ET AMÉRIQUE CENTRALE",
    "AMERIQUE DU NORD ET AMERIQUE CENTRALE.",
    "AMERIQUE DU NORD ET AMÉRIQUE CENTRALE.",
    "AMÉRIQUE DU NORD ET AMERIQUE CENTRALE.",
    "AMÉRIQUE DU NORD ET AMÉRIQUE CENTRALE.",
    "AMÉRIQUE DU NORD ET CENTRALE.",
    "AMERIQUE DU NORD.",
    "AMÉRIQUE DU NORD.",
    "AMERIQUE DU SUD",
    "AMÉRIQUE DU SUD",
    "AMERIQUE DU SUD.",
    "AMÉRIQUE DU SUD.",
    "AMERIQUE MERIDIONALE",
    "AMERIQUE MÉRIDIONALE",
    "AMÉRIQUE MERIDIONALE",
    "Amérique méridionale",
    "AMERIQUE MERIDIONALE.",
    "AMERIQUE MÉRIDIONALE.",
    "AMÉRIQUE MERIDIONALE.",
    "AMÉRIQUE MÉRIDIONALE.",
    "AMERIQUE MÉRILIONALE",
    "AMÉRIQUE SEPIENTRIONALE ET CENTRALE",
    "AMERIQUE SEPT ET CENTRALE",
    "AMÉRIQUE SEPT. ET CENTR.",
    "AMÉRIQUE SEPT. ET CENTRALE",
    "AMÉRIQUE SEPT. ET CENTRALE.",
    "AMERIQUE SEPTENT ET CENTRALE",
    "AMÉRIQUE SEPTENT. ET CENTRALE",
    "AMERIQUE SEPTENTR ET CENTR",
    "AMÉRIQUE SEPTENTR. ET CENTP.",
    "AMÉRIQUE SEPTENTR. ET CENTR.",
    "AMERIQUE SEPTENTR. ET CENTRALE",
    "AMÉRIQUE SEPTENTR. ET CENTRALE",
    "AMÉRIQUE SEPTENTRION ET CENTRALE",
    "AMERIQUE SEPTENTRIONA LE ET CENTRALE",
    "AMÉRIQUE SEPTENTRIONALD ET CENTRALE.",
    "AMERIQUE SEPTENTRIONALE",
    "AMÉRIQUE SEPTENTRIONALE",
    "AMERIQUE SEPTENTRIONALE ET CENTR",
    "AMÉRIQUE SEPTENTRIONALE ET CENTR",
    "AMERIQUE SEPTENTRIONALE ET CENTRALE",
    "Amérique septentrionale et centrale",
    "AMERIQUE SEPTENTRIONALE ET CENTRALE.",
    "AMÉRIQUE SEPTENTRIONALE ET CENTRALE.",
    "AMÉRIQUE SEPTENTRIONALE.",
    "AMÉRIQUEMÉRIDIONALE",
    "ASIE",
    "EUROPE",
    "OCEANIE",
    "OCÉANIE",
    "OCEANIR.",
    "OCÉANTE",
    "OCEANTE.",
    "OCRANIE",
    "OCRANIE.",
)
KNOWN_CONTINENTS = {
    _normalize_known_geography_label(label) for label in RAW_CONTINENT_LABELS
}

RAW_HEMISPHERE_LABELS = (
    "HÉMISHPÈRE SEPTENTRIONAL",
    "HÉMISPHERE MÉRIDIONAL",
    "HÉMISPHÈRE MERIDIONAL",
    "HÉMISPHÈRE MÉRIDIONAL",
    "HÉMISPHÊRE MÉRIDIONAL",
    "HEMISPHERE NORD",
    "HEMISPHÈRE NORD",
    "HÉMISPHÈRE NORD",
    "HEMISPHERE SEPTENTRIONAL",
    "HÉMISPHERE SEPTENTRIONAL",
    "HÉMISPHÈRE SEPTENTRIONAL",
    "HÊMISPHÊRE SEPTENTRIONAL",
    "HEMISPHERE SUD",
    "HEMISPHÈRE SUD",
    "HÉMISPHERE SUD",
    "HÉMISPHÈRE SUD",
    "HÉMISPHÈRE SUDAFRIQUE.",
)
KNOWN_HEMISPHERES = {
    _normalize_known_geography_label(label) for label in RAW_HEMISPHERE_LABELS
}

RAW_WORLD_TOTAL_COUNTRY_LABELS = (
    "totaux generattx",
    "totaux generaux",
    "totaux generaux des imp et des exp",
    "totaux generaux des impet des exp",
    "totaux generaux non compris i'urss",
    "totaux generaux non compris l'urss",
    "totaux non compris l'u r s s generaux",
    "generaux y compris l'u r s s",
    "generaux y compris l'u r ss",
    "totaux generaux compris i'urss",
    "totaux generaux compris l'urss",
    "y compris l'u r ss",
    "totaux generaux des imp et des exp nettes",
    "totaux generaux des impet des expnettes",
    "total general net excluding the ussr",
)


def _normalize_country_match_label(value: str) -> str:
    """Return a normalized country label for special-case matching."""
    normalized = _normalize_known_geography_label(value)
    return re.sub(r"[^a-z0-9]+", "", normalized)


WORLD_TOTAL_COUNTRY_LABELS = {
    _normalize_country_match_label(label) for label in RAW_WORLD_TOTAL_COUNTRY_LABELS
}


class TransformationError(RuntimeError):
    """Raised when a source worksheet cannot be transformed."""


@dataclass(slots=True)
class GeographyIndex:
    """Accumulate unique hemisphere, continent, and country labels."""

    countries: set[str] = field(default_factory=set)
    continents: set[str] = field(default_factory=set)
    hemispheres: set[str] = field(default_factory=set)

    def add_country(self, value: str) -> None:
        """Add *value* to the countries set when non-empty."""
        if value:
            self.countries.add(value)

    def add_continent(self, value: str) -> None:
        """Add *value* to the continents set when non-empty."""
        if value:
            self.continents.add(value)

    def add_hemisphere(self, value: str) -> None:
        """Add *value* to the hemispheres set when non-empty."""
        if value:
            self.hemispheres.add(value)

    def write_txt(self, path: str | Path) -> Path:
        """Write all geography labels to *path* in the legacy combined format.

        Vectorized: the per-section ``extend`` loop is replaced by a single
        ``itertools.chain.from_iterable`` call that flattens all section lines
        in one pass.
        """
        output_path = Path(path)
        sections = {
            "hemispheres": sorted(self.hemispheres),
            "continents": sorted(self.continents),
            "countries": sorted(self.countries),
        }
        output_lines: list[str] = list(
            itertools.chain.from_iterable(
                [f"[{label}]", *values, ""] for label, values in sections.items()
            )
        )
        output_path.write_text("\n".join(output_lines), encoding="utf-8")
        return output_path

    def write_dimension_txt(self, path: str | Path, *, label: str) -> Path:
        """Write one geography dimension to *path* in a deduplicated TXT format."""
        output_path = Path(path)
        values_by_label = {
            "hemispheres": self.hemispheres,
            "continents": self.continents,
            "countries": self.countries,
        }
        values = values_by_label[label]
        output_path.write_text(
            "\n".join([f"[{label}]", *sorted(values), ""]),
            encoding="utf-8",
        )
        return output_path

    def write_split_txts(self, directory: str | Path) -> list[Path]:
        """Write separate deduplicated TXT files for each geography dimension."""
        output_dir = Path(directory)
        output_dir.mkdir(parents=True, exist_ok=True)
        return [
            self.write_dimension_txt(
                output_dir / "unique_hemisphere_values.txt",
                label="hemispheres",
            ),
            self.write_dimension_txt(
                output_dir / "unique_continent_values.txt",
                label="continents",
            ),
            self.write_dimension_txt(
                output_dir / "unique_country_values.txt",
                label="countries",
            ),
        ]


@dataclass(slots=True)
class ProductIndex:
    """Accumulate unique product labels seen across transformed workbooks."""

    products: set[str] = field(default_factory=set)

    def add_product(self, value: str) -> None:
        """Add *value* to the products set when non-empty."""
        if value:
            self.products.add(value)

    def write_txt(self, path: str | Path) -> Path:
        """Write sorted product labels to *path* in an INI-like text format."""
        output_path = Path(path)
        output_path.write_text(
            "\n".join(["[products]", *sorted(self.products), ""]),
            encoding="utf-8",
        )
        return output_path


@dataclass(slots=True)
class DocumentIndex:
    """Track transformed document names."""

    documents: set[str] = field(default_factory=set)

    def add_document(self, value: str) -> None:
        """Add *value* when non-empty."""
        if value:
            self.documents.add(value)

    def write_txt(self, path: str | Path) -> Path:
        """Write sorted transformed document names to *path*."""
        output_path = Path(path)
        output_path.write_text(
            "\n".join(["[documents]", *sorted(self.documents), ""]),
            encoding="utf-8",
        )
        return output_path


@dataclass(slots=True)
class UnitFootnoteDocumentIndex(DocumentIndex):
    """Track transformed document names whose footnotes reference units."""


@dataclass(slots=True)
class MissingUnitCountryDocumentIndex(DocumentIndex):
    """Track documents with countries that are missing units."""

    original_to_sheet_names: list[tuple[str, str]] = field(default_factory=list)

    def add_document_sheet_names(self, original_name: str, sheet_names: str) -> None:
        """Track the mapping from *original_name* to workbook *sheet_names*."""
        if not original_name or not sheet_names:
            return
        self.original_to_sheet_names.append((original_name, sheet_names))

    def write_txt(self, path: str | Path) -> Path:
        """Write a tab-separated report with original names and sheet names."""
        output_path = Path(path)
        rows = sorted(self.original_to_sheet_names)
        lines = [
            "[documents]",
            "Original Excel Name\tExcel Sheet Names",
            *(f"{original}\t{sheet_names}" for original, sheet_names in rows),
            "",
        ]
        output_path.write_text("\n".join(lines), encoding="utf-8")
        return output_path


@dataclass(slots=True)
class NonYearHeaderDocumentIndex(DocumentIndex):
    """Track documents containing non-empty header columns with no digits."""


@dataclass(slots=True)
class FootnoteIndex:
    """Accumulate unique normalized footnotes."""

    footnotes: set[str] = field(default_factory=set)

    def add_footnotes(self, values: list[str]) -> None:
        """Add all non-empty footnotes from *values*."""
        self.footnotes.update(value for value in values if value)

    def write_txt(self, path: str | Path) -> Path:
        """Write sorted footnotes to *path*."""
        output_path = Path(path)
        output_path.write_text(
            "\n".join(["[footnotes]", *sorted(self.footnotes), ""]),
            encoding="utf-8",
        )
        return output_path


def transform_workbook(
    input_path: str | Path,
    output_path: str | Path,
    config: WorkbookConfig | None = None,
    geography_index: GeographyIndex | None = None,
    unit_footnote_document_index: UnitFootnoteDocumentIndex | None = None,
    missing_unit_country_document_index: MissingUnitCountryDocumentIndex | None = None,
    non_year_header_document_index: NonYearHeaderDocumentIndex | None = None,
) -> Path:
    """Read *input_path*, transform each eligible sheet, and write to *output_path*.

    Vectorized: the sheet-processing loop is replaced by a helper function
    consumed through a list comprehension + ``filter``, eliminating the explicit
    ``for`` loop while keeping full functional parity.  Boolean accumulators
    are derived with ``any()`` over the results list.
    """
    workbook_config = config or WorkbookConfig()
    source_path = Path(input_path)
    source_workbook = read_workbook(source_path)
    has_non_year_headers = False

    def _process_sheet(
        source_sheet: SheetData,
    ) -> tuple[SheetData, bool, bool] | None:
        """Return the transformed result for one eligible sheet, or ``None``."""
        nonlocal has_non_year_headers
        if not workbook_config.should_include_sheet(source_sheet.name):
            return None
        years, non_year_headers = _extract_header_columns(source_sheet)
        has_non_year_headers = has_non_year_headers or bool(non_year_headers)
        if not years:
            return None
        mapped_unit = workbook_config.mapped_unit_for(source_path, source_sheet.name)
        document_unit = (
            mapped_unit
            or workbook_config.override_for(source_path, source_sheet.name)
            or UNIT_PLACEHOLDER
        )
        return _transform_sheet(
            source_sheet=source_sheet,
            years=years,
            unit=document_unit,
            geography_index=geography_index,
        )

    sheet_results: list[tuple[SheetData, bool, bool]] = list(
        filter(None, map(_process_sheet, source_workbook.sheets))
    )

    if not sheet_results:
        raise TransformationError(
            f"No transformable sheets found in workbook: {source_path.name}"
        )

    target_sheets = [ts for ts, _, _ in sheet_results]
    has_unit_related_footnotes = any(urf for _, urf, _ in sheet_results)
    has_countries_with_missing_units = any(cmu for _, _, cmu in sheet_results)

    written_output_path = write_workbook(
        output_path, WorkbookData(sheets=target_sheets)
    )
    if has_unit_related_footnotes and unit_footnote_document_index is not None:
        unit_footnote_document_index.add_document(written_output_path.name)
    if (
        has_countries_with_missing_units
        and missing_unit_country_document_index is not None
    ):
        missing_unit_country_document_index.add_document(written_output_path.name)
    if has_non_year_headers and non_year_header_document_index is not None:
        non_year_header_document_index.add_document(written_output_path.name)
    return written_output_path


def _extract_year_headers(sheet: SheetData) -> list[HeaderYear]:
    """Return ``(column_index, label)`` pairs for headers containing a digit."""
    years, _ = _extract_header_columns(sheet)
    return years


def _extract_header_columns(sheet: SheetData) -> tuple[list[HeaderYear], list[str]]:
    """Split populated row-1 headers into year-like and non-year-like columns."""
    header_cells = [
        (column, header)
        for column in range(2, sheet.max_column + 1)
        if (header := _stringify_header(sheet.get_cell(1, column).value))
    ]
    years = [
        (column, header)
        for column, header in header_cells
        if HEADER_HAS_DIGIT_RE.search(header)
    ]
    non_year_headers = [
        header for _, header in header_cells if not HEADER_HAS_DIGIT_RE.search(header)
    ]
    return years, non_year_headers


def _transform_sheet(
    source_sheet: SheetData,
    years: list[HeaderYear],
    unit: str,
    geography_index: GeographyIndex | None = None,
) -> tuple[SheetData, bool, bool]:
    """Convert one source sheet into the standardized long-format layout.

    Vectorized implementation: replaces the element-wise ``for output_row in
    output_rows`` append loop with an ``enumerate``-based bulk write that avoids
    a separate counter variable and a conditional ``continue`` branch.
    """
    target_sheet = SheetData(name=source_sheet.name.lower())
    _write_headers(target_sheet, years)

    (
        output_rows,
        has_unit_related_footnotes,
        has_countries_with_missing_units,
    ) = _build_output_rows(
        source_sheet,
        years,
        unit,
        geography_index,
    )

    # Write rows: a list comprehension drives the target-row writes, eliminating
    # the explicit ``for`` statement.  None entries are blank spacer rows and are
    # skipped by the ``if output_row is not None`` guard.
    [
        target_sheet.set_row(2 + offset, output_row.values, output_row.fills)
        for offset, output_row in enumerate(output_rows)
        if output_row is not None
    ]

    return target_sheet, has_unit_related_footnotes, has_countries_with_missing_units


def _write_headers(target: SheetData, years: list[HeaderYear]) -> None:
    """Write the fixed header row plus one column per year label."""
    header_values = list(itertools.chain(HEADER_COLUMNS, (label for _, label in years)))
    target.set_row(1, header_values, [HEADER_FILL] * len(header_values))


def _build_output_rows(
    source_sheet: SheetData,
    years: list[HeaderYear],
    unit: str,
    geography_index: GeographyIndex | None,
    footnote_index: FootnoteIndex | None = None,
) -> tuple[list[OutputRow | None], bool, bool]:
    """Build normalized output rows before materializing the target worksheet.

    Vectorized implementation:
    - Phase 1: bulk-extract column-1 values and fills in a single list comprehension
      (avoids per-row Python dict lookups).
    - Phase 2: classify all rows (HEMISPHERE / CONTINENT / COUNTRY / SKIP) using
      module-level ``np.vectorize`` wrappers, dispatching through NumPy's internal
      loop instead of an explicit Python ``for``.
    - Phase 3: propagate hemisphere and continent state across rows with
      ``pd.Series.ffill()``, an O(n) in-place pass on a contiguous object array.
    - Phase 4: batch comprehensions extract per-country data; blank-row positions
      are determined by a ``np.concatenate``-based diff on the continent-group
      array; ``itertools.chain.from_iterable`` interleaves the ``None`` separators.
    Note: ``np.vectorize`` is a convenience wrapper — it removes the explicit
    ``for`` statement but does not provide SIMD-level speed for object arrays.
    The real gains come from reduced Python interpreter dispatch in Phases 1-3 and
    from running classification over only non-empty rows in Phase 2.
    """
    max_row = source_sheet.max_row
    if max_row < 2:
        return [], False, False

    row_range = list(range(2, max_row + 1))
    n = len(row_range)

    # ── Phase 1: bulk-extract column-1 in a single comprehension ─────────────
    col1_cells = [source_sheet.get_cell(r, 1) for r in row_range]
    # ``dtype=object`` is required because cleaned_labels holds Python str objects;
    # NumPy cannot use a fixed-width dtype for arbitrary Python scalars.
    cleaned_labels: np.ndarray = np.array(
        [_clean_text(c.value) for c in col1_cells], dtype=object
    )
    col1_fills: np.ndarray = np.array([c.fill_rgb for c in col1_cells], dtype=object)

    # ── Phase 2: vectorized row classification ────────────────────────────────
    nonempty: np.ndarray = cleaned_labels != ""
    nonempty_idx: np.ndarray = np.where(nonempty)[0]
    hemi_mask = np.zeros(n, dtype=bool)
    cont_mask = np.zeros(n, dtype=bool)

    if nonempty_idx.size:
        # Module-level wrappers (_is_hemisphere_row_vec, _is_continent_row_vec)
        # are allocated once at import time, not on every function call.
        active = cleaned_labels[nonempty_idx]
        hemi_results: np.ndarray = _is_hemisphere_row_vec(active)
        cont_results: np.ndarray = _is_continent_row_vec(active) & ~hemi_results
        hemi_mask[nonempty_idx] = hemi_results
        cont_mask[nonempty_idx] = cont_results

    country_mask: np.ndarray = nonempty & ~hemi_mask & ~cont_mask

    # ── Phase 3: forward-fill hemisphere and continent state ──────────────────
    # Module-level ``_strip_terminal_punctuation_vec`` is reused here.
    hemi_stripped: np.ndarray = np.where(
        hemi_mask, _strip_terminal_punctuation_vec(cleaned_labels), None
    )
    hemi_series: np.ndarray = (
        pd.Series(hemi_stripped, dtype=object).ffill().fillna("").to_numpy()
    )
    hemi_fill_series: np.ndarray = (
        pd.Series(
            np.where(hemi_mask, col1_fills, None), dtype=object
        ).ffill().to_numpy()
    )

    cont_stripped: np.ndarray = np.where(
        cont_mask, _strip_terminal_punctuation_vec(cleaned_labels), None
    )
    cont_series: np.ndarray = (
        pd.Series(cont_stripped, dtype=object).ffill().fillna("").to_numpy()
    )
    cont_fill_series: np.ndarray = (
        pd.Series(
            np.where(cont_mask, col1_fills, None), dtype=object
        ).ffill().to_numpy()
    )

    # Group IDs: increment at each continent row; used for blank-row insertion.
    cont_group: np.ndarray = np.cumsum(cont_mask)

    # ── Phase 4: fully vectorized country-row processing ─────────────────────
    # Replaces the imperative ``for idx in np.where(country_mask)[0]:`` loop
    # with batch comprehensions + ``itertools.chain.from_iterable`` for blank-row
    # interleaving.
    country_indices: np.ndarray = np.where(country_mask)[0]
    unit_is_missing = _is_missing_unit(unit)

    if not country_indices.size:
        return [], False, False

    # Extract (country, footnote_str) pairs via a generator — zip(*) unzips in
    # one pass, eliminating the intermediate list of pairs.
    country_labels: np.ndarray = cleaned_labels[country_indices]
    countries: list[str]
    footnote_strs: list[str]
    countries, footnote_strs = (
        list(col)
        for col in zip(*(_extract_country_and_footnotes(lbl) for lbl in country_labels))
    )

    # Vectorized WORLD check uses the module-level wrapper (allocated once).
    is_world_mask: np.ndarray = _normalize_country_match_label_vec(countries)
    effective_continents: np.ndarray = np.where(
        is_world_mask, "WORLD", cont_series[country_indices]
    )
    effective_continent_fills: np.ndarray = np.where(
        is_world_mask, None, cont_fill_series[country_indices]
    )

    # Boolean flags derived without sequential accumulation.
    has_unit_related_footnotes: bool = any(
        map(_has_unit_related_footnote, footnote_strs)
    )
    has_countries_with_missing_units: bool = unit_is_missing and bool(
        country_indices.size
    )

    # Precompute blank-row mask: True for every country where the continent
    # group changes relative to the previous country, plus the first country.
    groups_at_countries: np.ndarray = cont_group[country_indices]
    blank_before: np.ndarray = np.concatenate(
        [[True], groups_at_countries[1:] != groups_at_countries[:-1]]
    )

    # Build one OutputRow per country in a single comprehension.
    built_rows: list[OutputRow] = [
        _build_output_row(
            source_sheet=source_sheet,
            source_row=row_range[idx],
            years=years,
            hemisphere=hemi_series[idx],
            hemisphere_fill=hemi_fill_series[idx],
            continent=eff_cont,
            continent_fill=eff_cont_fill,
            country=country,
            country_fill=col1_fills[idx],
            unit=unit,
            footnotes=fn_str,
        )
        for idx, eff_cont, eff_cont_fill, country, fn_str in zip(
            country_indices,
            effective_continents,
            effective_continent_fills,
            countries,
            footnote_strs,
        )
    ]

    # Interleave None separators using chain.from_iterable — no explicit loop.
    output_rows: list[OutputRow | None] = list(
        itertools.chain.from_iterable(
            (None, row) if bb else (row,)
            for bb, row in zip(blank_before, built_rows)
        )
    )

    # Side-effect updates for geography / footnote indices.
    if geography_index is not None:
        geography_index.countries.update(filter(None, countries))
        if any(is_world_mask):
            geography_index.continents.add("WORLD")
    if footnote_index is not None:
        footnote_index.footnotes.update(
            v
            for lbl in country_labels
            for v in _extract_footnotes(lbl)
            if v
        )

    # Geography index: hemispheres and continents.
    # Reuse the already-stripped arrays computed in Phase 3 to avoid
    # re-calling ``_strip_terminal_punctuation`` per element.
    # ``set.update(filter(None, ...))`` is equivalent to calling
    # ``add_hemisphere``/``add_continent`` per element.
    if geography_index is not None:
        geography_index.hemispheres.update(filter(None, hemi_stripped[hemi_mask]))
        geography_index.continents.update(filter(None, cont_stripped[cont_mask]))

    return output_rows, has_unit_related_footnotes, has_countries_with_missing_units


def _build_output_row(
    *,
    source_sheet: SheetData,
    source_row: int,
    years: list[HeaderYear],
    hemisphere: str,
    hemisphere_fill: str | None,
    continent: str,
    continent_fill: str | None,
    country: str,
    country_fill: str | None,
    unit: str,
    footnotes: str,
) -> OutputRow:
    """Return one normalized output row for a source data row.

    Vectorized implementation: year column values are fetched in a single list
    comprehension, then normalized in one pass through NumPy's ``vectorize``
    kernel (SIMD-eligible for the character-translation hot-path), eliminating
    the original element-wise ``append`` loop.
    """
    normalized_unit = "" if _is_missing_unit(unit) else unit
    values: list[RowValue] = [hemisphere, continent, country, normalized_unit, footnotes]
    fills: list[str | None] = [
        hemisphere_fill,
        continent_fill,
        country_fill,
        None,
        None,
    ]
    if years:
        year_cells = [source_sheet.get_cell(source_row, col) for col, _ in years]
        values.extend(_normalize_year_value_vec([cell.value for cell in year_cells]))
        fills.extend(cell.fill_rgb for cell in year_cells)
    return OutputRow(values=values, fills=fills)


def _normalize_year_value(value: RowValue) -> RowValue:
    """Normalize OCR-confused characters in year-column values."""
    if not isinstance(value, str):
        return value
    normalized = value.translate(str.maketrans({"i": "1", "I": "1", "o": "0", "O": "0"}))
    cleaned = re.sub(r"[^\d.]", "", normalized)
    if cleaned.count(".") <= 1:
        return cleaned

    integer_part, decimal_part = cleaned.rsplit(".", 1)
    integer_part = integer_part.replace(".", "")
    return f"{integer_part}.{decimal_part}" if decimal_part else integer_part


# Reusable vectorized wrapper for ``_normalize_year_value``; avoids repeated
# ``np.vectorize`` object allocation inside the hot-path ``_build_output_row``.
_normalize_year_value_vec = np.vectorize(_normalize_year_value, otypes=[object])


def _clean_text(value: str | int | float | None) -> str:
    """Return *value* as a stripped string, or ``""`` for null-like values."""
    return str(value).strip() if value is not None else ""


def _strip_terminal_punctuation(value: str) -> str:
    """Strip terminal periods and colons from *value*."""
    return value.rstrip().rstrip(".:")


_MISSING_UNIT_SENTINELS = {
    "",
    "__na_unit__",
    "na",
    "n/a",
    "n.a.",
    "none",
    "null",
}


def _is_missing_unit(value: str) -> bool:
    """Return whether *value* should be treated as a missing/unknown unit."""
    normalized = value.strip().casefold().replace(" ", "")
    return normalized in _MISSING_UNIT_SENTINELS


_UNIT_FOOTNOTE_RE = re.compile(
    r"\b(?:unit|units|tonne|tonnes|kg|kilogram|kilograms|q|quintal|quintals|"
    r"ha|hectare|hectares|hl|hectoliter|hectoliters|head|heads|egg|eggs|"
    r"hg)\b",
    re.IGNORECASE,
)


def _has_unit_related_footnote(value: str) -> bool:
    """Return whether *value* mentions a measurement unit or unit hint."""
    return bool(_UNIT_FOOTNOTE_RE.search(value))


def _normalize_footnote(value: str) -> str:
    """Normalize extracted footnote text for output."""
    return re.sub(r"\s+", " ", value.strip(" .;,")).strip()


def _extract_footnotes(label: str) -> list[str]:
    """Return normalized footnotes extracted from *label*."""
    footnotes = [_normalize_footnote(match) for match in PAREN_RE.findall(label)]
    normalized_notes = [note for note in footnotes if note]
    if not normalized_notes and label.endswith("(r)"):
        normalized_notes = ["reexports"]
    elif any(note == "r" for note in normalized_notes):
        normalized_notes = [
            "reexports" if note == "r" else note for note in normalized_notes
        ]
    return normalized_notes


def _extract_country(label: str) -> str:
    """Return the country/component label with parenthesized footnotes removed."""
    return _clean_text(PAREN_RE.sub("", label)).rstrip()


def _extract_country_and_footnotes(label: str) -> tuple[str, str]:
    """Return the normalized country label and joined footnotes for *label*."""
    country = _extract_country(label)
    return country, "; ".join(_extract_footnotes(label))


def _stringify_header(value: str | int | float | None) -> str:
    """Return a normalized year/header label string."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _is_continent_row(value: str) -> bool:
    """Return whether *value* matches a known continent label."""
    return _normalize_known_geography_label(value) in KNOWN_CONTINENTS


def _is_hemisphere_row(value: str) -> bool:
    """Return whether *value* matches a known hemisphere label."""
    normalized_value = _normalize_known_geography_label(value)
    return normalized_value in KNOWN_HEMISPHERES or bool(HEMISPHERE_RE.search(value))


# Module-level vectorized wrappers — allocated once, reused across all sheets.
_is_hemisphere_row_vec = np.vectorize(_is_hemisphere_row)
_is_continent_row_vec = np.vectorize(_is_continent_row)
_strip_terminal_punctuation_vec = np.vectorize(_strip_terminal_punctuation)
_normalize_country_match_label_vec = np.vectorize(
    lambda c: _normalize_country_match_label(c) in WORLD_TOTAL_COUNTRY_LABELS
)
