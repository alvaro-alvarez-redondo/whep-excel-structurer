"""Core preprocessing logic for Excel workbooks.

Implements sheet renaming, column removal, row_flag value clearing, country
normalization, horizontal concatenation, metadata deduplication, and placeholder
reinsertion.  Optimized for large-scale batch execution.
"""

from __future__ import annotations

import importlib.util
import unicodedata
import re
from functools import lru_cache
from pathlib import Path

import pandas as pd

from ..paths import PREPROCESS_COUNTRY_LABEL_PATTERNS_PATH

SHEET_PREFIX = "a-r_"
COLUMNS_TO_REMOVE = {"unit", "footnotes"}
ROW_FLAGS_TO_REMOVE = {"subcategory_propagated", "continent_in_country_col"}
METADATA_FIELDS = ["continent", "country", "original_country", "row_flag"]
PLACEHOLDER_COLUMNS = ["unit", "footnotes"]
SPACER_PREFIX = "__spacer_"
OUTPUT_SHEET_NAME = "data"
LOWERCASE_EXCLUDED_COLS = {"continent", "country"}

# Columns whose values and names must NOT receive OCR corrections.
_OCR_EXCLUDED_COLS = frozenset(
    {"continent", "country", "original_country", "unit", "footnotes", "row_flag"}
)
# Characters that look like "1" or "5" in scanned/OCR text.
_OCR_TO_1_RE = r"[|il/\\]"
_OCR_TO_5_RE = r"s"
_OCR_TO_0_RE = r"o"
_OCR_TO_8_RE = r"b"
_OCR_TO_6_RE = r"g"
_OCR_TO_2_RE = r"z"


def _fold_value(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return ascii_only.casefold()


def _normalize_original_country_match_value(value: str) -> str:
    """Return letters-and-spaces-only text for original_country matching."""
    folded = _fold_value(value)
    letters_spaces_only = re.sub(r"[^a-z\s]+", " ", folded)
    return re.sub(r"\s+", " ", letters_spaces_only).strip()


CountryLabelPatterns = dict[str, tuple[tuple[str, str], ...]]
RowReconstructionInputs = tuple[str, ...]


def _is_enabled(value: object) -> bool:
    """Return whether a configuration row is enabled."""
    if pd.isna(value):
        return True
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().casefold()
    return text not in {"", "0", "false", "f", "no", "n", "off", "disabled"}


def _parse_variant_chars(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value)
    return "".join(ch for ch in text if not ch.isspace() and ch not in ",;|")


def _normalize_letter_substitutions(
    rows: tuple[tuple[object, object], ...],
) -> dict[str, str]:
    substitutions: dict[str, str] = {}
    for canonical, variants in rows:
        canonical_text = _fold_value(str(canonical)).strip()
        if not canonical_text:
            continue
        canonical_char = canonical_text[0]
        variant_chars = _parse_variant_chars(variants)
        variant_chars = _fold_value(variant_chars)
        chars = "".join(dict.fromkeys(canonical_char + variant_chars))
        substitutions[canonical_char] = chars
    return substitutions


def _pattern_from_canonical_label(
    canonical_label: str, letter_substitutions: dict[str, str]
) -> str:
    folded = _fold_value(canonical_label)
    folded = re.sub(r"\s+", " ", folded).strip()
    pieces: list[str] = []

    for char in folded:
        if char.isspace():
            pieces.append(r"\s+")
            continue

        variants = letter_substitutions.get(char)
        if variants:
            escaped_variants = "".join(re.escape(variant) for variant in variants)
            pieces.append(f"[{escaped_variants}]")
        else:
            pieces.append(re.escape(char))

    return "".join(pieces)


def _build_country_label_patterns(
    letter_rows: tuple[tuple[object, object], ...],
    country_rows: tuple[tuple[object, object, object], ...],
) -> CountryLabelPatterns:
    letter_substitutions = _normalize_letter_substitutions(letter_rows)
    grouped: dict[str, list[tuple[str, str]]] = {}

    for continent, canonical_input, correct_output in country_rows:
        if pd.isna(continent) or pd.isna(canonical_input) or pd.isna(correct_output):
            continue
        continent_key = _normalize_country_label_series(
            pd.Series([continent])
        ).iat[0]
        if not continent_key:
            continue
        pattern = _pattern_from_canonical_label(
            str(canonical_input), letter_substitutions
        )
        grouped.setdefault(continent_key, []).append((pattern, str(correct_output)))

    return {key: tuple(value) for key, value in grouped.items()}


def _default_country_label_patterns() -> CountryLabelPatterns:
    return _build_country_label_patterns(
        DEFAULT_OCR_LETTER_SUBSTITUTIONS,
        DEFAULT_COUNTRY_LABEL_RULES,
    )


def _country_label_patterns_from_excel(path: Path) -> CountryLabelPatterns:
    letter_df = pd.read_excel(path, sheet_name="letter_dictionary")
    country_df = pd.read_excel(path, sheet_name="country_patterns")

    letter_rows = tuple(
        letter_df[["canonical_char", "variants"]].itertuples(index=False, name=None)
    )

    if "enabled" in country_df.columns:
        enabled = country_df["enabled"].map(_is_enabled)
        country_df = country_df.loc[enabled]

    country_rows = tuple(
        country_df[["continent", "canonical_input", "correct_output"]].itertuples(
            index=False, name=None
        )
    )
    return _build_country_label_patterns(letter_rows, country_rows)


@lru_cache(maxsize=8)
def load_country_label_patterns(path: str | None = None) -> CountryLabelPatterns:
    """Load country label regexes generated from canonical Excel rules."""
    pattern_path = (
        Path(path) if path is not None else PREPROCESS_COUNTRY_LABEL_PATTERNS_PATH
    )
    if not pattern_path.exists():
        return _default_country_label_patterns()
    return _country_label_patterns_from_excel(pattern_path)


def _row_reconstruction_inputs_from_excel(path: Path) -> RowReconstructionInputs:
    """Load enabled row-reconstruction input strings from the config workbook."""
    try:
        row_df = pd.read_excel(path, sheet_name="row_reconstruction")
    except ValueError:
        return ()

    if "input" not in row_df.columns:
        return ()

    if "enabled" in row_df.columns:
        row_df = row_df.loc[row_df["enabled"].map(_is_enabled)]

    inputs: list[str] = []
    seen: set[str] = set()
    for value in row_df["input"]:
        if pd.isna(value):
            continue
        text = str(value).strip()
        normalized = _normalize_country_label_series(pd.Series([text])).iat[0]
        if not text or not normalized or normalized in seen:
            continue
        inputs.append(text)
        seen.add(normalized)
    return tuple(inputs)


@lru_cache(maxsize=8)
def load_row_reconstruction_inputs(path: str | None = None) -> RowReconstructionInputs:
    """Load enabled row-reconstruction inputs from the config workbook."""
    pattern_path = (
        Path(path) if path is not None else PREPROCESS_COUNTRY_LABEL_PATTERNS_PATH
    )
    if not pattern_path.exists():
        return ()
    return _row_reconstruction_inputs_from_excel(pattern_path)


DEFAULT_OCR_LETTER_SUBSTITUTIONS = (
    ("a", "aeou"),
    ("e", "aeou"),
    ("o", "aeou"),
    ("u", "aeou"),
    ("2", "2z"),
)

DEFAULT_COUNTRY_LABEL_RULES = (
    ("asia", "22 provinces", "China 22 provinces"),
    ("asia", "manchuria", "China Manchuria"),
    ("asia", "taiwan", "China Taiwan"),
    ("europe", "saar", "France Saar"),
    ("europe", "bizone", "Germany Bizone"),
    ("europe", "western", "Germany Western"),
    ("europe", "eastern", "Germany Eastern"),
    ("europe", "french zone", "Germany French Zone"),
    ("europe", "soviet zone", "Germany Soviet Zone"),
    ("europe", "soviet", "Germany Soviet Zone"),
    ("europe", "berlin", "Germany Berlin"),
)


def _get_fast_engine() -> str | None:
    """Return the fastest available Excel engine name, or ``None``."""
    if importlib.util.find_spec("calamine") is not None:
        return "calamine"
    if importlib.util.find_spec("fastexcel") is not None:
        return "fastexcel"
    return None


def _read_all_sheets(path: Path) -> dict[str, pd.DataFrame]:
    """Read every sheet from *path* in a single pass using the fastest engine.

    Falls back to ``openpyxl`` when no fast engine is installed.
    """
    engine = _get_fast_engine()
    kwargs: dict[str, object] = {}
    if engine is not None:
        kwargs["engine"] = engine

    return pd.read_excel(path, sheet_name=None, **kwargs)


def remove_sheet_prefix(sheet_name: str, prefix: str = SHEET_PREFIX) -> str:
    """Remove *prefix* from the start of *sheet_name* when present."""
    if sheet_name.startswith(prefix):
        return sheet_name[len(prefix) :]
    return sheet_name


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Apply column removal, row_flag value clearing, and country normalization.

    Operates on a copy so the original DataFrame is never mutated.
    """
    df = df.copy()

    cols_present = [c for c in df.columns if c in COLUMNS_TO_REMOVE]
    if cols_present:
        df = df.drop(columns=cols_present)

    if "row_flag" in df.columns:
        df.loc[df["row_flag"].isin(ROW_FLAGS_TO_REMOVE), "row_flag"] = None

    if "country" in df.columns:
        col = df["country"].astype(str)
        col = col.str.replace(r"[()]", "", regex=True)
        col = col.str.replace(r"- ", "", regex=True)
        col = col.str.strip()
        df["country"] = col

    return df


def concatenate_sheets_horizontally(dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """Horizontally concatenate DataFrames with one spacer column between each.

    All DataFrames are reindexed to the same row count (the maximum) so row
    alignment is preserved even when tables have different lengths.
    """
    if not dfs:
        return pd.DataFrame()

    max_len = max(len(df) for df in dfs)

    # Fast path: all DFs already have the same length
    if all(len(df) == max_len for df in dfs):
        result = dfs[0]
        for idx, df in enumerate(dfs[1:], start=1):
            spacer = pd.DataFrame({f"{SPACER_PREFIX}{idx}__": [""] * max_len})
            result = pd.concat([result, spacer, df], axis=1)
        return result

    # Slow path: reindex to max length
    aligned = [df.reindex(range(max_len)).reset_index(drop=True) for df in dfs]
    result = aligned[0]
    for idx, df in enumerate(aligned[1:], start=1):
        spacer = pd.DataFrame({f"{SPACER_PREFIX}{idx}__": [""] * max_len})
        result = pd.concat([result, spacer, df], axis=1)

    return result


def apply_ocr_corrections(df: pd.DataFrame) -> pd.DataFrame:
    """Fix common OCR misreads in column names and cell values.

    Replaces OCR-confused characters in non-excluded columns, including:
    ``| i I l / \\`` → ``1``, ``S s`` → ``5``, ``o`` → ``0``,
    ``b`` → ``8``, ``g`` → ``6``, ``z`` → ``2``.

    Corrections are skipped for the protected metadata columns
    (``continent``, ``country``, ``original_country``, ``unit``,
    ``footnotes``, ``row_flag``).
    """
    df = df.copy()

    # 1. Fix column names
    new_names = []
    for col in df.columns:
        if col in _OCR_EXCLUDED_COLS:
            new_names.append(col)
        else:
            fixed = str(col).lower()
            fixed = re.sub(_OCR_TO_1_RE, "1", fixed)
            fixed = re.sub(_OCR_TO_5_RE, "5", fixed)
            fixed = re.sub(_OCR_TO_0_RE, "0", fixed)
            fixed = re.sub(_OCR_TO_8_RE, "8", fixed)
            fixed = re.sub(_OCR_TO_6_RE, "6", fixed)
            fixed = re.sub(_OCR_TO_2_RE, "2", fixed)
            new_names.append(fixed)
    df.columns = new_names

    # 2. Fix cell values in non-excluded columns
    for idx, col in enumerate(df.columns):
        if col in _OCR_EXCLUDED_COLS:
            continue
        series = df.iloc[:, idx]
        if series.dtype == object:
            s = series.astype(str).str.lower()
            s = s.str.replace(_OCR_TO_1_RE, "1", regex=True)
            s = s.str.replace(_OCR_TO_5_RE, "5", regex=True)
            s = s.str.replace(_OCR_TO_0_RE, "0", regex=True)
            s = s.str.replace(_OCR_TO_8_RE, "8", regex=True)
            s = s.str.replace(_OCR_TO_6_RE, "6", regex=True)
            s = s.str.replace(_OCR_TO_2_RE, "2", regex=True)
            df.iloc[:, idx] = s

    return df


def deduplicate_metadata_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Remove redundant copies of metadata columns that are identical row-wise.

    For each metadata field, compare only columns with the same name.
    Retain the first occurrence; remove only subsequent duplicates.
    Preserve non-identical columns even if they share the same name.
    """
    cols_to_drop: set[int] = set()

    for field in METADATA_FIELDS:
        positions = [i for i, c in enumerate(df.columns) if c == field]
        if len(positions) <= 1:
            continue

        kept_positions: list[int] = []

        for pos in positions:
            current_col = df.iloc[:, pos]
            is_duplicate = False

            # Compare against every previously-kept column with the same name
            for kept_pos in kept_positions:
                if df.iloc[:, kept_pos].equals(current_col):
                    is_duplicate = True
                    break

            if is_duplicate:
                cols_to_drop.add(pos)
            else:
                kept_positions.append(pos)

    if not cols_to_drop:
        return df

    # Select by position (iloc) because df.drop() drops by *name*, and after
    # horizontal concatenation we have duplicate column names.  Dropping by name
    # would remove every column with that name, not just the targeted position.
    cols_to_keep = [i for i in range(len(df.columns)) if i not in cols_to_drop]
    return df.iloc[:, cols_to_keep]


def insert_placeholder_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Insert empty *unit* and *footnotes* columns immediately right of *original_country*.

    If *original_country* is absent the placeholders are inserted at the front
    of the DataFrame as a graceful fallback.
    """
    df = df.copy()
    target = "original_country"

    if target in df.columns:
        insert_pos = list(df.columns).index(target) + 1
    else:
        insert_pos = 0

    for placeholder in reversed(PLACEHOLDER_COLUMNS):
        df.insert(insert_pos, placeholder, pd.Series([None] * len(df)))

    return df


def uniquify_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure column names are unique by appending _2, _3, ... to duplicates."""
    counts: dict[str, int] = {}
    new_columns: list[str] = []
    changed = False

    for col in df.columns:
        base = str(col)
        count = counts.get(base, 0) + 1
        counts[base] = count
        if count == 1:
            new_columns.append(base)
        else:
            changed = True
            new_columns.append(f"{base}_{count}")

    if not changed:
        return df

    df = df.copy()
    df.columns = new_columns
    return df


def insert_blank_rows_before_continents(df: pd.DataFrame) -> pd.DataFrame:
    """Insert an empty row before each continent start based on the continent column.

    The check only uses the ``continent`` column and ignores other columns.
    """
    if "continent" not in df.columns or df.empty:
        return df

    continent = df["continent"].where(df["continent"].notna(), "")
    continent = continent.astype(str).str.strip()
    prev_continent = continent.shift(1).fillna("")
    is_new_continent = (continent != "") & (continent != prev_continent)

    if not is_new_continent.any():
        return df

    empty_row = pd.DataFrame([[None] * len(df.columns)], columns=df.columns)
    pieces: list[pd.DataFrame] = []
    for idx, row in df.iterrows():
        if is_new_continent.loc[idx]:
            pieces.append(empty_row)
        pieces.append(row.to_frame().T)

    return pd.concat(pieces, ignore_index=True)


def replace_duplicate_country_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Replace duplicate country entries with "Total" when original_country is "Total".

    Only checks the immediately preceding row for the duplicate country.
    """
    if "country" not in df.columns or "original_country" not in df.columns:
        return df

    country = df["country"].where(df["country"].notna(), "").astype(str).str.strip()
    original = (
        df["original_country"].where(df["original_country"].notna(), "")
        .astype(str)
        .str.strip()
    )
    is_total = original.map(_normalize_original_country_match_value) == "total"

    country_folded = country.map(_fold_value)
    is_dup_prev = country_folded == country_folded.shift(1).fillna("")
    mask = is_total & is_dup_prev

    if not mask.any():
        return df

    df = df.copy()
    df.loc[mask, "country"] = "Total"
    return df


def _fold_text_series(series: pd.Series) -> pd.Series:
    text = series.where(series.notna(), "").astype(str)
    return (
        text.str.normalize("NFKD")
        .str.encode("ascii", "ignore")
        .str.decode("ascii")
        .str.lower()
    )


def _normalize_country_label_series(
    series: pd.Series, *, replace_zero: bool = True
) -> pd.Series:
    folded = _fold_text_series(series).str.replace(r"\s+", " ", regex=True).str.strip()
    if replace_zero:
        folded = folded.str.replace("0", "o", regex=False)
    return folded


def _matched_country_labels(
    folded_country: pd.Series, patterns: tuple[tuple[str, str], ...]
) -> pd.Series:
    if not patterns:
        return pd.Series(pd.NA, index=folded_country.index, dtype="object")

    match_columns = [
        folded_country.str.fullmatch(pattern, na=False).rename(idx)
        for idx, (pattern, _label) in enumerate(patterns)
    ]
    matches = pd.concat(match_columns, axis=1)
    has_match = matches.any(axis=1)
    if not has_match.any():
        return pd.Series(pd.NA, index=folded_country.index, dtype="object")

    labels_by_column = pd.Series(
        [label for _pattern, label in patterns], index=matches.columns
    )
    return matches.idxmax(axis=1).map(labels_by_column).where(has_match)


def _prefix_countries_by_patterns(
    df: pd.DataFrame,
    *,
    continent_name: str,
    patterns: tuple[tuple[str, str], ...],
) -> pd.DataFrame:
    if "country" not in df.columns or "continent" not in df.columns:
        return df

    continent_folded = _normalize_country_label_series(df["continent"])
    continent_mask = continent_folded == continent_name
    if not continent_mask.any():
        return df

    country_folded = _normalize_country_label_series(df["country"])
    labels = _matched_country_labels(country_folded, patterns)
    mask = continent_mask & labels.notna()
    if not mask.any():
        return df

    df = df.copy()
    df.loc[mask, "country"] = labels.loc[mask]
    return df


def reconstruct_rows_from_previous_country(
    df: pd.DataFrame, inputs: RowReconstructionInputs | None = None
) -> pd.DataFrame:
    """Prepend the exact previous country to rows matching configured inputs.

    The ``row_reconstruction`` sheet supplies the input strings to match. When
    the current ``country`` value matches one of those enabled inputs, the
    current value is replaced with ``<previous country> <current country>``.
    Matching is accent-insensitive, case-insensitive, whitespace-normalized, and
    treats OCR ``0`` as ``o`` in the same way as country-label matching.
    """
    if "country" not in df.columns or df.empty:
        return df

    configured_inputs = (
        inputs if inputs is not None else load_row_reconstruction_inputs()
    )
    if not configured_inputs:
        return df

    normalized_inputs = {
        _normalize_country_label_series(pd.Series([value])).iat[0]
        for value in configured_inputs
    }
    normalized_inputs.discard("")
    if not normalized_inputs:
        return df

    country_text = (
        df["country"].where(df["country"].notna(), "").astype(str).str.strip()
    )
    country_folded = _normalize_country_label_series(country_text)
    previous_country = country_text.shift(1).fillna("").str.strip()
    mask = country_folded.isin(normalized_inputs) & (previous_country != "")
    if not mask.any():
        return df

    df = df.copy()
    df.loc[mask, "country"] = (
        previous_country.loc[mask] + " " + country_text.loc[mask]
    )
    return df


def apply_country_label_patterns(
    df: pd.DataFrame, patterns_by_continent: CountryLabelPatterns | None = None
) -> pd.DataFrame:
    """Apply generated country label patterns grouped by continent."""
    patterns_by_continent = patterns_by_continent or load_country_label_patterns()
    for continent_name, patterns in patterns_by_continent.items():
        df = _prefix_countries_by_patterns(
            df, continent_name=continent_name, patterns=patterns
        )
    return df


def prefix_china_countries(df: pd.DataFrame) -> pd.DataFrame:
    """Prefix China to specific country labels in the ``country`` column."""
    patterns = load_country_label_patterns().get("asia", ())
    return _prefix_countries_by_patterns(
        df, continent_name="asia", patterns=patterns
    )


def prefix_germany_countries_in_europe(df: pd.DataFrame) -> pd.DataFrame:
    """Prefix Germany to specific country labels when continent is Europe."""
    patterns = tuple(
        (pattern, label)
        for pattern, label in load_country_label_patterns().get("europe", ())
        if label.startswith("Germany ")
    )
    return _prefix_countries_by_patterns(
        df, continent_name="europe", patterns=patterns
    )


def prefix_france_countries_in_europe(df: pd.DataFrame) -> pd.DataFrame:
    """Prefix France to specific country labels when continent is Europe."""
    patterns = tuple(
        (pattern, label)
        for pattern, label in load_country_label_patterns().get("europe", ())
        if label.startswith("France ")
    )
    return _prefix_countries_by_patterns(
        df, continent_name="europe", patterns=patterns
    )


def lowercase_original_country(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize values in the ``original_country`` column."""
    if "original_country" not in df.columns:
        return df

    df = df.copy()
    series = df["original_country"]
    mask = series.notna()
    if not mask.any():
        return df
    df.loc[mask, "original_country"] = (
        series.loc[mask].astype(str).map(_normalize_original_country_match_value)
    )
    return df


def lowercase_text_values(df: pd.DataFrame) -> pd.DataFrame:
    """Lowercase text values except in protected geography columns."""
    text_column_positions = [
        idx
        for idx, col in enumerate(df.columns)
        if col not in LOWERCASE_EXCLUDED_COLS
        and (
            pd.api.types.is_object_dtype(df.iloc[:, idx])
            or pd.api.types.is_string_dtype(df.iloc[:, idx])
        )
    ]
    if not text_column_positions:
        return df

    df = df.copy()
    for idx in text_column_positions:
        series = df.iloc[:, idx]
        df.iloc[:, idx] = series.map(
            lambda value: value.lower() if isinstance(value, str) else value
        )
    return df


def replace_ampersands_in_columns(
    df: pd.DataFrame, columns: tuple[str, ...]
) -> pd.DataFrame:
    """Replace '&' with 'and' in specific string columns."""
    missing = [col for col in columns if col not in df.columns]
    if len(missing) == len(columns):
        return df

    df = df.copy()
    for col in columns:
        if col not in df.columns:
            continue
        series = df[col]
        mask = series.notna()
        if not mask.any():
            continue
        df.loc[mask, col] = (
            series.loc[mask].astype(str).str.replace("&", "and", regex=False)
        )
    return df


def normalize_region_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize totals for specific region labels in original_country."""
    required = {"original_country", "country", "continent", "footnotes"}
    if not required.issubset(df.columns):
        return df

    original = (
        df["original_country"].where(df["original_country"].notna(), "")
        .astype(str)
        .str.strip()
    )
    folded = original.map(_normalize_original_country_match_value)
    target_set = {
        "north america",
        "latin america",
        "lacin america",
        "near east",
        "far east",
    }
    mask = folded.isin(target_set)
    if not mask.any():
        return df

    continent_map = {
        "north america": "AMERICA",
        "latin america": "AMERICA",
        "lacin america": "AMERICA",
        "near east": "ASIA",
        "far east": "ASIA",
    }

    df = df.copy()
    df.loc[mask, "footnotes"] = original[mask]
    df.loc[mask, "country"] = "Total"
    df.loc[mask, "continent"] = folded[mask].map(continent_map)
    return df


def remove_original_country_column(df: pd.DataFrame) -> pd.DataFrame:
    """Remove the ``original_country`` column when present."""
    if "original_country" not in df.columns:
        return df
    return df.drop(columns=["original_country"])


def process_workbook(
    input_path: Path,
    output_path: Path,
    country_label_patterns_path: Path | None = None,
) -> None:
    """Process a single workbook end-to-end and write the result to *output_path*.

    Steps:
        1. Read every sheet in a single pass with the fastest engine.
        2. Rename sheets by stripping the ``a-r_`` prefix.
        3. Clean each sheet (remove columns, filter rows, normalize country).
        4. Horizontally concatenate all sheets with spacer columns.
          5. Apply OCR corrections (``|iIl/\\`` → ``1``, ``Ss`` → ``5``,
              ``o`` → ``0``, ``b`` → ``8``, ``g`` → ``6``, ``z`` → ``2``).
        6. Deduplicate repeated metadata columns.
        7. Reinsert placeholder columns.
        8. Remove the ``original_country`` helper column.
        9. Write a single sheet named ``data``.
    """
    sheets = _read_all_sheets(input_path)

    cleaned_dfs: list[pd.DataFrame] = []
    for sheet_name in sheets:
        df = clean_dataframe(sheets[sheet_name])
        cleaned_dfs.append(df)

    combined = concatenate_sheets_horizontally(cleaned_dfs)
    combined = apply_ocr_corrections(combined)
    combined = deduplicate_metadata_columns(combined)
    combined = insert_placeholder_columns(combined)
    combined = uniquify_column_names(combined)
    combined = lowercase_original_country(combined)
    combined = replace_duplicate_country_totals(combined)
    row_reconstruction_inputs = load_row_reconstruction_inputs(
        str(country_label_patterns_path) if country_label_patterns_path else None
    )
    combined = reconstruct_rows_from_previous_country(
        combined, row_reconstruction_inputs
    )
    country_label_patterns = load_country_label_patterns(
        str(country_label_patterns_path) if country_label_patterns_path else None
    )
    combined = apply_country_label_patterns(combined, country_label_patterns)
    combined = normalize_region_totals(combined)
    combined = replace_ampersands_in_columns(combined, ("continent", "country"))
    combined = lowercase_text_values(combined)
    combined = insert_blank_rows_before_continents(combined)
    combined = remove_original_country_column(combined)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    combined.to_excel(output_path, sheet_name=OUTPUT_SHEET_NAME, index=False)
