"""Text normalization utilities shared across the package.

This module contains pure string helpers that have no dependency on any other
internal module.  They form the foundation of the dependency graph.
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path

# Matches all Unicode combining (diacritic) characters after NFKD normalization.
# Covers the standard combining diacritics blocks; faster than a per-character
# unicodedata.combining() call inside a generator expression.
_COMBINING_RE = re.compile(
    r"[\u0300-\u036f"   # Combining Diacritical Marks
    r"\u1dc0-\u1dff"    # Combining Diacritical Marks Supplement
    r"\u20d0-\u20ff"    # Combining Diacritical Marks for Symbols
    r"\ufe20-\ufe2f]+"  # Combining Half Marks
)


def normalize_text(value: str) -> str:
    """Return *value* lowercased, accent-stripped, and whitespace-normalized.

    Uses a compiled regex to strip Unicode combining characters after NFKD
    decomposition — significantly faster than the equivalent per-character
    ``unicodedata.combining()`` generator expression for long strings.
    """
    decomposed = unicodedata.normalize("NFKD", value)
    without_accents = _COMBINING_RE.sub("", decomposed)
    return " ".join(without_accents.replace("_", " ").strip().lower().split())


def derive_product_from_document(document_name: str) -> str:
    """Infer the product name from an Excel document filename stem.

    Tokens after the first 4-digit year token (skipping any following numeric
    tokens) are joined and normalized.  Falls back to the full stem when no
    year token is found.

    Vectorized: the ``while`` loop that skips post-year numeric tokens is
    replaced by a ``next()`` call over a range generator expression.
    """
    stem = Path(document_name).stem
    tokens = [token for token in stem.split("_") if token]
    if not tokens:
        return ""

    year_idx = next(
        (idx for idx, token in enumerate(tokens) if len(token) == 4 and token.isdigit()),
        None,
    )
    if year_idx is None:
        return normalize_text(stem)

    product_start = next(
        (i for i in range(year_idx + 1, len(tokens)) if not tokens[i].isdigit()),
        len(tokens),
    )

    product_tokens = tokens[product_start:] or tokens[-1:]
    return normalize_text(" ".join(product_tokens))


def format_elapsed(seconds: float) -> str:
    """Format *seconds* into a human-readable elapsed-time string.

    * ``< 60 s``      → ``45.50s``
    * ``>= 60 s``     → ``01:27.12`` (mm:ss.ms)
    * ``>= 60 min``   → ``01:02:30.15`` (hh:mm:ss.ms)
    """
    if seconds < 60:
        return f"{seconds:.2f}s"

    minutes, sec = divmod(seconds, 60)
    if minutes < 60:
        return f"{int(minutes):02d}:{sec:05.2f}"

    hours, minutes = divmod(int(minutes), 60)
    return f"{hours:02d}:{minutes:02d}:{sec:05.2f}"
