from __future__ import annotations

from pathlib import Path

import pandas as pd

from iia_excel_reorg.core.preprocessor import (
    apply_country_label_patterns,
    lowercase_text_values,
    lowercase_original_country,
    load_country_label_patterns,
    normalize_region_totals,
    prefix_china_countries,
    prefix_france_countries_in_europe,
    prefix_germany_countries_in_europe,
    remove_original_country_column,
    replace_duplicate_country_totals,
)


def test_replace_duplicate_country_totals_ignores_non_letters_in_original_country() -> (
    None
):
    df = pd.DataFrame(
        {
            "country": ["France", "France"],
            "original_country": ["France", "Total 42 *"],
        }
    )

    result = replace_duplicate_country_totals(df)

    assert result.loc[1, "country"] == "Total"


def test_lowercase_original_country_removes_non_letters_and_normalizes_spaces() -> None:
    df = pd.DataFrame({"original_country": ["  Total 42  * icon "]})

    result = lowercase_original_country(df)

    assert result.loc[0, "original_country"] == "total icon"


def test_lowercase_text_values_skips_geography_and_preserves_numbers() -> None:
    df = pd.DataFrame(
        {
            "continent": ["EUROPE"],
            "country": ["France"],
            "description": ["Mixed CASE"],
            "mixed": ["VALUE", 12],
            "amount": [10, 20],
        }
    )

    result = lowercase_text_values(df)

    assert result["continent"].to_list() == ["EUROPE"]
    assert result["country"].to_list() == ["France"]
    assert result["description"].to_list() == ["mixed case"]
    assert result["mixed"].to_list() == ["value", 12]
    assert result["amount"].to_list() == [10, 20]


def test_saar_is_prefixed_as_france_not_germany() -> None:
    df = pd.DataFrame({"continent": ["EUROPE"], "country": ["Saar"]})

    result = prefix_france_countries_in_europe(df)
    result = prefix_germany_countries_in_europe(result)

    assert result.loc[0, "country"] == "France Saar"


def test_country_prefixes_are_applied_across_matching_rows() -> None:
    df = pd.DataFrame(
        {
            "continent": ["ASIA", "ASIA", "EUROPE", "EUROPE", "EUROPE", "AMERICA"],
            "country": [
                "22 provinces",
                "Manchuria",
                "Bizone",
                "Soviet zone",
                "Berlin",
                "Bizone",
            ],
        }
    )

    result = prefix_china_countries(df)
    result = prefix_france_countries_in_europe(result)
    result = prefix_germany_countries_in_europe(result)

    assert result["country"].to_list() == [
        "China 22 provinces",
        "China Manchuria",
        "Germany Bizone",
        "Germany Soviet Zone",
        "Germany Berlin",
        "Bizone",
    ]


def test_country_label_patterns_can_be_loaded_from_excel(tmp_path: Path) -> None:
    path = tmp_path / "country_label_patterns.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(
            {
                "canonical_char": ["a", "e", "o", "u"],
                "variants": ["aeou", "aeou", "aeou", "aeou"],
            }
        ).to_excel(writer, sheet_name="letter_dictionary", index=False)
        pd.DataFrame(
            {
                "continent": ["EUROPE"],
                "canonical_input": ["customland"],
                "correct_output": ["Test Customland"],
                "enabled": [True],
            }
        ).to_excel(writer, sheet_name="country_patterns", index=False)

    patterns = load_country_label_patterns(str(path))
    df = pd.DataFrame({"continent": ["EUROPE"], "country": ["Castamland"]})

    result = apply_country_label_patterns(df, patterns)

    assert result.loc[0, "country"] == "Test Customland"


def test_missing_country_label_patterns_excel_generates_preset(tmp_path: Path) -> None:
    path = tmp_path / "config" / "country_label_patterns.xlsx"

    patterns = load_country_label_patterns(str(path))

    assert path.exists()
    workbook = pd.ExcelFile(path)
    assert workbook.sheet_names == [
        "letter_dictionary",
        "country_patterns",
        "description",
    ]

    letter_df = pd.read_excel(path, sheet_name="letter_dictionary")
    country_df = pd.read_excel(path, sheet_name="country_patterns")
    description_df = pd.read_excel(path, sheet_name="description")

    assert list(letter_df.columns) == ["canonical_char", "variants"]
    assert list(country_df.columns) == [
        "continent",
        "canonical_input",
        "correct_output",
        "enabled",
    ]
    assert list(description_df.columns) == [
        "sheet",
        "column",
        "description",
        "example",
    ]
    assert {"letter_dictionary", "country_patterns", "matching_behavior"}.issubset(
        set(description_df["sheet"])
    )
    assert patterns["asia"]


def test_country_label_patterns_reconstruct_split_previous_row_fragment(
    tmp_path: Path,
) -> None:
    path = tmp_path / "country_label_patterns.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(
            {
                "canonical_char": ["i", "m", "o"],
                "variants": ["i,l,1,|", "m,rn", "o,0"],
            }
        ).to_excel(writer, sheet_name="letter_dictionary", index=False)
        pd.DataFrame(
            {
                "continent": ["EUROPE"],
                "canonical_input": ["united kingdom"],
                "correct_output": ["United Kingdom"],
                "enabled": [True],
            }
        ).to_excel(writer, sheet_name="country_patterns", index=False)

    patterns = load_country_label_patterns(str(path))
    df = pd.DataFrame(
        {
            "continent": ["EUROPE", "EUROPE"],
            "country": ["United", "Kingdorn"],
        }
    )

    result = apply_country_label_patterns(df, patterns)

    assert result.loc[0, "country"] == "United"
    assert result.loc[1, "country"] == "United Kingdom"


def test_country_label_patterns_handle_formatting_and_ocr_variations_from_excel(
    tmp_path: Path,
) -> None:
    path = tmp_path / "country_label_patterns.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(
            {
                "canonical_char": ["a", "e", "i", "l", "o", "u"],
                "variants": ["aeou", "aeou", "i1l|", "l1i|", "aeou0", "aeou"],
            }
        ).to_excel(writer, sheet_name="letter_dictionary", index=False)
        pd.DataFrame(
            {
                "continent": ["AFRICA"],
                "canonical_input": ["cote divoire"],
                "correct_output": ["Cote d'Ivoire"],
                "enabled": [True],
            }
        ).to_excel(writer, sheet_name="country_patterns", index=False)

    patterns = load_country_label_patterns(str(path))
    df = pd.DataFrame({"continent": ["AFRICA"], "country": [" C0te--d lv0ire "]})

    result = apply_country_label_patterns(df, patterns)

    assert result.loc[0, "country"] == "Cote d'Ivoire"


def test_normalize_region_totals_ignores_non_letters_in_original_country() -> None:
    df = pd.DataFrame(
        {
            "original_country": ["North America 7 *", "Lacin America"],
            "country": ["North America", "Lacin America"],
            "continent": ["", ""],
            "footnotes": ["", ""],
        }
    )

    result = normalize_region_totals(df)

    assert result.loc[0, "country"] == "Total"
    assert result.loc[0, "continent"] == "AMERICA"
    assert result.loc[0, "footnotes"] == "North America 7 *"
    assert result.loc[1, "country"] == "Total"
    assert result.loc[1, "continent"] == "AMERICA"
    assert result.loc[1, "footnotes"] == "Lacin America"


def test_remove_original_country_column_drops_helper_column() -> None:
    df = pd.DataFrame(
        {
            "country": ["Total"],
            "original_country": ["total"],
            "unit": ["tonnes"],
        }
    )

    result = remove_original_country_column(df)

    assert list(result.columns) == ["country", "unit"]
