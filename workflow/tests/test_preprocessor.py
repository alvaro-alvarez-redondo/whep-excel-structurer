from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from iia_excel_reorg import preprocess_pipeline
from iia_excel_reorg.preprocess_pipeline import main as preprocess_main
from iia_excel_reorg.core.preprocessor import (
    apply_country_label_patterns,
    create_country_label_patterns_preset,
    lowercase_text_values,
    lowercase_original_country,
    load_country_label_patterns,
    load_row_reconstruction_patterns,
    normalize_region_totals,
    prefix_china_countries,
    prefix_france_countries_in_europe,
    prefix_germany_countries_in_europe,
    remove_original_country_column,
    replace_duplicate_country_totals,
    reconstruct_rows_from_previous_country,
)


def test_replace_duplicate_country_totals_ignores_non_letters_in_original_country() -> None:
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
            "continent": ["EUROPE", "EUROPE"],
            "country": ["France", "Germany"],
            "description": ["Mixed CASE", "Already lower"],
            "mixed": ["VALUE", 12],
            "amount": [10, 20],
        }
    )

    result = lowercase_text_values(df)

    assert result["continent"].to_list() == ["EUROPE", "EUROPE"]
    assert result["country"].to_list() == ["France", "Germany"]
    assert result["description"].to_list() == ["mixed case", "already lower"]
    assert result["mixed"].to_list() == ["value", 12]
    assert result["amount"].to_list() == [10, 20]


def test_lowercase_text_values_preserves_missing_string_columns() -> None:
    df = pd.DataFrame(
        {
            "row_flag": pd.Series([pd.NA, pd.NA], dtype="string"),
            "description": pd.Series(["Mixed CASE", pd.NA], dtype="string"),
        }
    )

    result = lowercase_text_values(df)

    assert result["row_flag"].isna().to_list() == [True, True]
    assert result["description"].to_list() == ["mixed case", pd.NA]


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


def test_country_label_patterns_preset_includes_row_reconstruction_examples(
    tmp_path: Path,
) -> None:
    path = tmp_path / "country_label_patterns.xlsx"

    created_path = create_country_label_patterns_preset(path)

    assert created_path == path
    with pd.ExcelFile(path) as workbook:
        assert workbook.sheet_names == [
            "letter_dictionary",
            "country_patterns",
            "row_reconstruction",
        ]
        row_reconstruction = pd.read_excel(workbook, sheet_name="row_reconstruction")

    assert list(row_reconstruction.columns) == ["input", "enabled"]
    assert row_reconstruction["input"].to_list() == [
        "dependent territories",
        "overseas territories",
        "protectorates",
    ]


def test_country_label_patterns_preset_does_not_overwrite_existing_file(
    tmp_path: Path,
) -> None:
    path = tmp_path / "country_label_patterns.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame({"input": ["custom"], "enabled": [True]}).to_excel(
            writer, sheet_name="row_reconstruction", index=False
        )

    created_path = create_country_label_patterns_preset(path)
    row_reconstruction = pd.read_excel(created_path, sheet_name="row_reconstruction")

    assert row_reconstruction["input"].to_list() == ["custom"]


def test_preprocess_pipeline_creates_country_label_patterns_preset(
    tmp_path: Path,
    monkeypatch,
) -> None:
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", ["iia-prepare"])

    preprocess_main()

    preset_path = tmp_path / "data" / "country_label_patterns.xlsx"
    assert preset_path.exists()
    row_reconstruction = pd.read_excel(preset_path, sheet_name="row_reconstruction")
    assert list(row_reconstruction.columns) == ["input", "enabled"]
    assert "dependent territories" in row_reconstruction["input"].to_list()


def test_process_entry_hides_traceback_by_default(
    tmp_path: Path,
    monkeypatch,
    capsys,
) -> None:
    def fail_process_workbook(*args, **kwargs) -> None:
        raise TypeError("example failure")

    source_path = tmp_path / "input.xlsx"
    target_path = tmp_path / "output.xlsx"
    monkeypatch.setattr(preprocess_pipeline, "process_workbook", fail_process_workbook)

    preprocess_pipeline._process_entry(
        (source_path, target_path),
        tmp_path / "country_label_patterns.xlsx",
    )

    output = capsys.readouterr().out
    assert "TypeError: example failure" in output
    assert "Traceback" not in output


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


def test_row_reconstruction_patterns_use_letter_dictionary(tmp_path: Path) -> None:
    path = tmp_path / "country_label_patterns.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(
            {
                "canonical_char": ["e"],
                "variants": ["ea"],
            }
        ).to_excel(writer, sheet_name="letter_dictionary", index=False)
        pd.DataFrame(
            {
                "input": [
                    "dependent territories",
                    "disabled label",
                    "DEPENDENT TERRITORIES",
                ],
                "enabled": [True, False, True],
            }
        ).to_excel(writer, sheet_name="row_reconstruction", index=False)

    patterns = load_row_reconstruction_patterns(str(path))
    df = pd.DataFrame({"country": ["United Kingdom", "Dapandant Tarritorias"]})

    result = reconstruct_rows_from_previous_country(df, patterns)

    assert result["country"].to_list() == [
        "United Kingdom",
        "United Kingdom Dapandant Tarritorias",
    ]


def test_reconstruct_rows_from_previous_country_concatenates_exact_previous_row() -> None:
    df = pd.DataFrame(
        {
            "country": ["United Kingdom", "Dependent   Territories", "France"],
            "value": [1, 2, 3],
        }
    )

    result = reconstruct_rows_from_previous_country(df, (r"dependent\s+territories",))

    assert result["country"].to_list() == [
        "United Kingdom",
        "United Kingdom Dependent   Territories",
        "France",
    ]
    assert result["value"].to_list() == [1, 2, 3]


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
