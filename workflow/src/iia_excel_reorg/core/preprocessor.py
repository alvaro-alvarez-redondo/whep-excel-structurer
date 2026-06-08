"""Core preprocessing logic for Excel workbooks.

Implements sheet renaming, column removal, row_flag value clearing, country
normalization, horizontal concatenation, metadata deduplication, and placeholder
reinsertion.  Optimized for large-scale batch execution.
"""

from __future__ import annotations

import importlib.util
import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
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


CountryLabelPattern = tuple[str, str, str]
CountryLabelPatterns = dict[str, tuple[CountryLabelPattern, ...]]
RowReconstructionPattern = tuple[str, str, str, str, str]
RowReconstructionPatterns = dict[str, tuple[RowReconstructionPattern, ...]]


@dataclass(frozen=True)
class CountryLabelPatternConfig:
    """Country-label rules loaded from the Excel pattern workbook.

    ``country_patterns`` are direct current-row rules.
    ``row_reconstruction_patterns`` are explicit previous-row + current-row rules.

    ``get`` and ``__getitem__`` intentionally mirror the previous dict-like API for
    callers that only need direct country patterns.
    """

    country_patterns: CountryLabelPatterns
    row_reconstruction_patterns: RowReconstructionPatterns

    def get(
        self, continent_name: str, default: tuple[CountryLabelPattern, ...] = ()
    ) -> tuple[CountryLabelPattern, ...]:
        return self.country_patterns.get(continent_name, default)

    def __getitem__(self, continent_name: str) -> tuple[CountryLabelPattern, ...]:
        return self.country_patterns[continent_name]


_COUNTRY_LABEL_MIN_FUZZY_RATIO = 0.88


def _country_pattern_parts(pattern: tuple[str, ...]) -> CountryLabelPattern:
    """Return ``(regex, label, match_key)`` for legacy/new pattern entries."""
    regex = pattern[0]
    label = pattern[1]
    match_key = pattern[2] if len(pattern) > 2 else label
    return regex, label, match_key


def _parse_variant_tokens(value: object) -> tuple[str, ...]:
    if pd.isna(value):
        return ()
    text = _fold_value(str(value)).strip()
    if not text:
        return ()
    if re.search(r"[,;\s]", text):
        tokens = re.split(r"[,;\s]+", text)
    else:
        tokens = list(text)
    return tuple(token for token in tokens if token)


def _normalize_letter_substitutions(
    rows: tuple[tuple[object, object], ...],
) -> dict[str, tuple[str, ...]]:
    substitutions: dict[str, tuple[str, ...]] = {}
    for canonical, variants in rows:
        canonical_text = _fold_value(str(canonical)).strip()
        if not canonical_text:
            continue
        canonical_char = canonical_text[0]
        tokens = (canonical_char, *_parse_variant_tokens(variants))
        substitutions[canonical_char] = tuple(dict.fromkeys(tokens))
    return substitutions


def _pattern_from_canonical_label(
    canonical_label: str, letter_substitutions: dict[str, tuple[str, ...]]
) -> str:
    """Build a permissive regex from a configured country-label variant.

    The regex intentionally handles generic OCR/formatting noise only.  The
    country-specific value being recognized still comes entirely from the
    ``country_label_patterns`` configuration.
    """
    folded = _fold_value(canonical_label)
    folded = re.sub(r"\s+", " ", folded).strip()
    pieces: list[str] = []
    previous_was_word_char = False

    for char in folded:
        if char.isspace() or not char.isalnum():
            if pieces and pieces[-1] != r"[\W_]*":
                pieces.append(r"[\W_]*")
            previous_was_word_char = False
            continue

        if previous_was_word_char:
            pieces.append(r"[\W_]*")

        variants = letter_substitutions.get(char)
        if variants:
            escaped_variants = "|".join(re.escape(variant) for variant in variants)
            pieces.append(f"(?:{escaped_variants})")
        else:
            pieces.append(re.escape(char))
        previous_was_word_char = True

    return "".join(pieces)


def _country_label_match_key(value: object) -> str:
    """Normalize country-label text for generic fuzzy matching.

    This removes formatting noise and folds common OCR substitutions without
    encoding country-specific assumptions in code.
    """
    if pd.isna(value):
        return ""
    folded = _fold_value(str(value))
    folded = folded.replace("0", "o")
    folded = folded.replace("1", "i")
    folded = folded.replace("|", "i")
    folded = re.sub(r"rn", "m", folded)
    folded = re.sub(r"[^a-z0-9]+", "", folded)
    return folded


def _build_country_label_patterns(
    letter_rows: tuple[tuple[object, object], ...],
    country_rows: tuple[tuple[object, object, object], ...],
) -> CountryLabelPatterns:
    letter_substitutions = _normalize_letter_substitutions(letter_rows)
    grouped: dict[str, list[CountryLabelPattern]] = {}

    for continent, canonical_input, correct_output in country_rows:
        if pd.isna(continent) or pd.isna(canonical_input) or pd.isna(correct_output):
            continue
        continent_key = _normalize_country_label_series(pd.Series([continent])).iat[0]
        if not continent_key:
            continue
        pattern = _pattern_from_canonical_label(
            str(canonical_input), letter_substitutions
        )
        grouped.setdefault(continent_key, []).append(
            (
                pattern,
                str(correct_output),
                _country_label_match_key(str(canonical_input)),
            )
        )

    return {key: tuple(value) for key, value in grouped.items()}


def _build_row_reconstruction_patterns(
    letter_rows: tuple[tuple[object, object], ...],
    row_reconstruction_rows: tuple[tuple[object, object, object, object], ...],
) -> RowReconstructionPatterns:
    letter_substitutions = _normalize_letter_substitutions(letter_rows)
    grouped: dict[str, list[RowReconstructionPattern]] = {}

    for (
        continent,
        previous_input,
        current_input,
        correct_output,
    ) in row_reconstruction_rows:
        if (
            pd.isna(continent)
            or pd.isna(previous_input)
            or pd.isna(current_input)
            or pd.isna(correct_output)
        ):
            continue
        continent_key = _normalize_country_label_series(pd.Series([continent])).iat[0]
        if not continent_key:
            continue
        previous_text = str(previous_input)
        current_text = str(current_input)
        grouped.setdefault(continent_key, []).append(
            (
                _pattern_from_canonical_label(previous_text, letter_substitutions),
                _pattern_from_canonical_label(current_text, letter_substitutions),
                str(correct_output),
                _country_label_match_key(previous_text),
                _country_label_match_key(current_text),
            )
        )

    return {key: tuple(value) for key, value in grouped.items()}


def _default_country_label_patterns() -> CountryLabelPatternConfig:
    return CountryLabelPatternConfig(
        country_patterns=_build_country_label_patterns(
            DEFAULT_OCR_LETTER_SUBSTITUTIONS,
            DEFAULT_COUNTRY_LABEL_RULES,
        ),
        row_reconstruction_patterns=_build_row_reconstruction_patterns(
            DEFAULT_OCR_LETTER_SUBSTITUTIONS,
            DEFAULT_ROW_RECONSTRUCTION_RULES,
        ),
    )


def _enabled_rows(df: pd.DataFrame) -> pd.DataFrame:
    if "enabled" not in df.columns:
        return df
    enabled = df["enabled"].fillna(True).astype(bool)
    return df.loc[enabled]


def _enabled_pattern_rows(
    df: pd.DataFrame, input_column: str
) -> tuple[tuple[object, object, object], ...]:
    df = _enabled_rows(df)
    if df.empty:
        return ()

    return tuple(
        df[["continent", input_column, "correct_output"]].itertuples(
            index=False, name=None
        )
    )


def _enabled_row_reconstruction_rows(
    df: pd.DataFrame,
) -> tuple[tuple[object, object, object, object], ...]:
    df = _enabled_rows(df)
    if df.empty:
        return ()

    return tuple(
        df[
            ["continent", "previous_input", "current_input", "correct_output"]
        ].itertuples(index=False, name=None)
    )


def _country_label_patterns_from_excel(path: Path) -> CountryLabelPatternConfig:
    workbook = pd.ExcelFile(path)
    letter_df = pd.read_excel(workbook, sheet_name="letter_dictionary")
    country_df = pd.read_excel(workbook, sheet_name="country_patterns")
    if "row_reconstruction" in workbook.sheet_names:
        row_reconstruction_df = pd.read_excel(
            workbook, sheet_name="row_reconstruction"
        )
    else:
        row_reconstruction_df = pd.DataFrame(
            columns=[
                "continent",
                "previous_input",
                "current_input",
                "correct_output",
                "enabled",
            ]
        )

    letter_rows = tuple(
        letter_df[["canonical_char", "variants"]].itertuples(index=False, name=None)
    )

    country_rows = _enabled_pattern_rows(country_df, "canonical_input")
    row_reconstruction_rows = _enabled_row_reconstruction_rows(row_reconstruction_df)
    return CountryLabelPatternConfig(
        country_patterns=_build_country_label_patterns(letter_rows, country_rows),
        row_reconstruction_patterns=_build_row_reconstruction_patterns(
            letter_rows, row_reconstruction_rows
        ),
    )


def _country_label_patterns_description_rows() -> tuple[dict[str, str], ...]:
    return (
        {
            "sheet": "letter_dictionary",
            "column": "canonical_char",
            "description": (
                "Single intended character used when building country-label "
                "patterns from configured canonical_input values."
            ),
            "example": "m",
        },
        {
            "sheet": "letter_dictionary",
            "column": "variants",
            "description": (
                "OCR alternatives accepted for canonical_char. Use commas, "
                "semicolons, or spaces for multi-character tokens; otherwise "
                "each character is treated as a separate variant."
            ),
            "example": "m,rn",
        },
        {
            "sheet": "country_patterns",
            "column": "continent",
            "description": (
                "Continent/region value for which this country rule is active. "
                "Matching is case-insensitive and whitespace-normalized."
            ),
            "example": "EUROPE",
        },
        {
            "sheet": "country_patterns",
            "column": "canonical_input",
            "description": (
                "Current-row country label, alias, misspelling, or OCR artifact "
                "to recognize."
            ),
            "example": "united kingdom",
        },
        {
            "sheet": "country_patterns",
            "column": "correct_output",
            "description": "Canonical country label written to the country column when matched.",
            "example": "United Kingdom",
        },
        {
            "sheet": "country_patterns",
            "column": "enabled",
            "description": "Set to FALSE to keep a rule in the preset but ignore it at runtime.",
            "example": "TRUE",
        },
        {
            "sheet": "row_reconstruction",
            "column": "continent",
            "description": (
                "Continent/region value for which this previous-row plus "
                "current-row reconstruction rule is active."
            ),
            "example": "EUROPE",
        },
        {
            "sheet": "row_reconstruction",
            "column": "previous_input",
            "description": (
                "Country text expected in the immediately previous row before "
                "a reconstruction is allowed."
            ),
            "example": "united kingdom",
        },
        {
            "sheet": "row_reconstruction",
            "column": "current_input",
            "description": (
                "Country text expected in the current row before this row is "
                "rewritten to correct_output."
            ),
            "example": "dependent territories",
        },
        {
            "sheet": "row_reconstruction",
            "column": "correct_output",
            "description": "Canonical country label written to the current row when matched.",
            "example": "United Kingdom",
        },
        {
            "sheet": "row_reconstruction",
            "column": "enabled",
            "description": "Set to FALSE to keep a reconstruction rule in the preset but ignore it.",
            "example": "TRUE",
        },
    )


def write_country_label_patterns_preset(path: Path) -> None:
    """Create a country-label pattern workbook preset at *path*."""
    path.parent.mkdir(parents=True, exist_ok=True)

    letter_df = pd.DataFrame(
        DEFAULT_OCR_LETTER_SUBSTITUTIONS, columns=["canonical_char", "variants"]
    )
    country_df = pd.DataFrame(
        DEFAULT_COUNTRY_LABEL_RULES,
        columns=["continent", "canonical_input", "correct_output"],
    )
    country_df["enabled"] = True
    row_reconstruction_df = pd.DataFrame(
        DEFAULT_ROW_RECONSTRUCTION_RULES,
        columns=["continent", "previous_input", "current_input", "correct_output"],
    )
    row_reconstruction_df["enabled"] = True
    description_df = pd.DataFrame(_country_label_patterns_description_rows())

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        letter_df.to_excel(writer, sheet_name="letter_dictionary", index=False)
        country_df.to_excel(writer, sheet_name="country_patterns", index=False)
        row_reconstruction_df.to_excel(
            writer, sheet_name="row_reconstruction", index=False
        )
        description_df.to_excel(writer, sheet_name="description", index=False)


@lru_cache(maxsize=8)
def load_country_label_patterns(path: str | None = None) -> CountryLabelPatternConfig:
    """Load country label regexes generated from canonical Excel rules."""
    pattern_path = (
        Path(path) if path is not None else PREPROCESS_COUNTRY_LABEL_PATTERNS_PATH
    )
    if not pattern_path.exists():
        write_country_label_patterns_preset(pattern_path)
    return _country_label_patterns_from_excel(pattern_path)


DEFAULT_OCR_LETTER_SUBSTITUTIONS = (
    ("a", "aeou"),
    ("e", "aeou"),
    ("i", "i,l,1,|"),
    ("l", "l,i,1,|"),
    ("m", "m,rn"),
    ("o", "aeou0"),
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
DEFAULT_ROW_RECONSTRUCTION_RULES: tuple[tuple[str, str, str, str], ...] = ()


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
        df["original_country"]
        .where(df["original_country"].notna(), "")
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


def _is_fuzzy_country_label_match(candidate_key: str, pattern_key: str) -> bool:
    if not candidate_key or not pattern_key:
        return False
    if candidate_key == pattern_key:
        return True
    max_len = max(len(candidate_key), len(pattern_key))
    if max_len < 5:
        return False
    length_delta = abs(len(candidate_key) - len(pattern_key))
    allowed_delta = 1 if max_len < 10 else 2
    if length_delta > allowed_delta:
        return False
    return SequenceMatcher(None, candidate_key, pattern_key).ratio() >= (
        _COUNTRY_LABEL_MIN_FUZZY_RATIO
    )


def _matched_country_labels(
    country_text: pd.Series, patterns: tuple[tuple[str, ...], ...]
) -> pd.Series:
    """Return configured labels for country text matching configured patterns."""
    if not patterns:
        return pd.Series(pd.NA, index=country_text.index, dtype="object")

    folded_country = _normalize_country_label_series(country_text)
    pattern_parts = tuple(_country_pattern_parts(pattern) for pattern in patterns)
    match_columns = [
        folded_country.str.fullmatch(regex, na=False).rename(idx)
        for idx, (regex, _label, _match_key) in enumerate(pattern_parts)
    ]
    matches = pd.concat(match_columns, axis=1)
    has_match = matches.any(axis=1)

    labels_by_column = pd.Series(
        [label for _regex, label, _match_key in pattern_parts], index=matches.columns
    )
    labels = matches.idxmax(axis=1).map(labels_by_column).where(has_match)

    missing_mask = labels.isna()
    if missing_mask.any():
        match_keys = tuple(match_key for _regex, _label, match_key in pattern_parts)
        labels_values = tuple(label for _regex, label, _match_key in pattern_parts)
        candidate_keys = country_text.loc[missing_mask].map(_country_label_match_key)
        fuzzy_matches = candidate_keys.map(
            lambda candidate_key: next(
                (
                    labels_values[idx]
                    for idx, match_key in enumerate(match_keys)
                    if _is_fuzzy_country_label_match(candidate_key, match_key)
                ),
                pd.NA,
            )
        )
        labels.loc[missing_mask] = fuzzy_matches

    return labels


def _country_text_matches_pattern(
    country_text: pd.Series, regex: str, match_key: str
) -> pd.Series:
    regex_matches = _normalize_country_label_series(country_text).str.fullmatch(
        regex, na=False
    )
    missing_mask = ~regex_matches
    if missing_mask.any():
        candidate_keys = country_text.loc[missing_mask].map(_country_label_match_key)
        fuzzy_matches = candidate_keys.map(
            lambda candidate_key: _is_fuzzy_country_label_match(
                candidate_key, match_key
            )
        )
        regex_matches.loc[missing_mask] = fuzzy_matches
    return regex_matches


def _matched_row_reconstruction_labels(
    previous_country_text: pd.Series,
    current_country_text: pd.Series,
    patterns: tuple[RowReconstructionPattern, ...],
) -> pd.Series:
    """Return labels when both previous-row and current-row patterns match."""
    labels = pd.Series(pd.NA, index=current_country_text.index, dtype="object")
    if not patterns:
        return labels

    for (
        previous_regex,
        current_regex,
        label,
        previous_match_key,
        current_match_key,
    ) in patterns:
        missing_mask = labels.isna()
        if not missing_mask.any():
            break
        previous_matches = _country_text_matches_pattern(
            previous_country_text.loc[missing_mask],
            previous_regex,
            previous_match_key,
        )
        current_matches = _country_text_matches_pattern(
            current_country_text.loc[missing_mask],
            current_regex,
            current_match_key,
        )
        matched_index = previous_matches.index[previous_matches & current_matches]
        labels.loc[matched_index] = label

    return labels


def _previous_and_current_country_text(
    country: pd.Series,
) -> tuple[pd.Series, pd.Series]:
    current = country.where(country.notna(), "").astype(str).str.strip()
    previous = current.shift(1).fillna("").astype(str).str.strip()
    return previous, current


def _prefix_countries_by_patterns(
    df: pd.DataFrame,
    *,
    continent_name: str,
    patterns: tuple[tuple[str, ...], ...],
    row_reconstruction_patterns: tuple[RowReconstructionPattern, ...] = (),
) -> pd.DataFrame:
    if "country" not in df.columns or "continent" not in df.columns:
        return df

    continent_folded = _normalize_country_label_series(df["continent"])
    continent_mask = continent_folded == continent_name
    if not continent_mask.any():
        return df

    labels = _matched_country_labels(df["country"], patterns)
    if row_reconstruction_patterns:
        previous_country, current_country = _previous_and_current_country_text(
            df["country"]
        )
        combined_labels = _matched_row_reconstruction_labels(
            previous_country, current_country, row_reconstruction_patterns
        )
        labels = labels.where(labels.notna(), combined_labels)
    mask = continent_mask & labels.notna()
    if not mask.any():
        return df

    df = df.copy()
    df.loc[mask, "country"] = labels.loc[mask]
    return df


def apply_country_label_patterns(
    df: pd.DataFrame,
    patterns_by_continent: CountryLabelPatternConfig | CountryLabelPatterns | None = None,
) -> pd.DataFrame:
    """Apply direct and explicit row-reconstruction country label patterns."""
    pattern_config = patterns_by_continent or load_country_label_patterns()
    if isinstance(pattern_config, CountryLabelPatternConfig):
        direct_patterns = pattern_config.country_patterns
        row_reconstruction_patterns = pattern_config.row_reconstruction_patterns
    else:
        direct_patterns = pattern_config
        row_reconstruction_patterns = {}

    continent_names = tuple(
        dict.fromkeys((*direct_patterns.keys(), *row_reconstruction_patterns.keys()))
    )
    for continent_name in continent_names:
        df = _prefix_countries_by_patterns(
            df,
            continent_name=continent_name,
            patterns=direct_patterns.get(continent_name, ()),
            row_reconstruction_patterns=row_reconstruction_patterns.get(
                continent_name, ()
            ),
        )
    return df


def prefix_china_countries(df: pd.DataFrame) -> pd.DataFrame:
    """Prefix China to specific country labels in the ``country`` column."""
    patterns = load_country_label_patterns().get("asia", ())
    return _prefix_countries_by_patterns(df, continent_name="asia", patterns=patterns)


def prefix_germany_countries_in_europe(df: pd.DataFrame) -> pd.DataFrame:
    """Prefix Germany to specific country labels when continent is Europe."""
    patterns = tuple(
        pattern
        for pattern in load_country_label_patterns().get("europe", ())
        if _country_pattern_parts(pattern)[1].startswith("Germany ")
    )
    return _prefix_countries_by_patterns(df, continent_name="europe", patterns=patterns)


def prefix_france_countries_in_europe(df: pd.DataFrame) -> pd.DataFrame:
    """Prefix France to specific country labels when continent is Europe."""
    patterns = tuple(
        pattern
        for pattern in load_country_label_patterns().get("europe", ())
        if _country_pattern_parts(pattern)[1].startswith("France ")
    )
    return _prefix_countries_by_patterns(df, continent_name="europe", patterns=patterns)


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
        df["original_country"]
        .where(df["original_country"].notna(), "")
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
