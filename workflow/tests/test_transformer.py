from __future__ import annotations

import importlib.util
from pathlib import Path
from typing import Any

from iia_excel_reorg.cli import (
    DuplicateOriginalDocumentIndex,
    _compute_output_subdir,
    _ensure_workspace,
    _extract_sheet_names,
    _iter_workbooks_structured,
)
from iia_excel_reorg.config import WorkbookConfig, load_config
from iia_excel_reorg.naming import (
    canonical_document_name,
    extract_source_product,
    infer_yearbook_metadata,
)
from iia_excel_reorg.transformer import (
    DocumentIndex,
    FootnoteIndex,
    GeographyIndex,
    MissingUnitCountryDocumentIndex,
    NonYearHeaderDocumentIndex,
    ProductIndex,
    UnitFootnoteDocumentIndex,
    _extract_footnotes,
    _is_continent_row,
    _is_hemisphere_row,
    transform_workbook,
)
from iia_excel_reorg.unit_rules import assign_unit
from iia_excel_reorg.xlsx_io import (
    SheetData,
    WorkbookData,
    read_workbook,
    write_workbook,
)

GREEN = "FF00FF00"
YELLOW = "FFFFFF00"
ORANGE = "FFFFA500"


def _write_config(path: Path, lines: list[str]) -> Path:
    """Write a small test YAML configuration file and return its path."""
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def _build_source_workbook(path: Path, *, include_imports: bool = False) -> None:
    """Build a representative source workbook used by transformation tests."""
    area = SheetData(name="AREA")
    area.set_cell(1, 2, "1909-1913")
    area.set_cell(1, 3, "1922")
    area.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL")
    area.set_cell(3, 1, "EUROPE.")
    area.set_cell(
        4,
        1,
        "Belgique-Luxembourg (reexports) (special case)",
        fill_rgb=GREEN,
    )
    area.set_cell(4, 2, 17268, fill_rgb=YELLOW)
    area.set_cell(4, 3, 11887, fill_rgb=ORANGE)
    area.set_cell(5, 1, "Germany", fill_rgb=GREEN)
    area.set_cell(5, 2, 284000, fill_rgb=GREEN)

    sheets = [area]

    production = SheetData(name="PRODUCTION")
    production.set_cell(1, 2, "1909-10/1913")
    production.set_cell(1, 3, "1922-23")
    production.set_cell(2, 1, "Hemisphère méridional")
    production.set_cell(3, 1, "Amérique")
    production.set_cell(4, 1, "Canada", fill_rgb=GREEN)
    production.set_cell(4, 2, 194876, fill_rgb=GREEN)
    production.set_cell(4, 3, 315569, fill_rgb=YELLOW)
    sheets.append(production)

    if include_imports:
        imports = SheetData(name="IMPORTS")
        imports.set_cell(1, 2, "1934-1938")
        imports.set_cell(1, 3, "1946")
        imports.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL")
        imports.set_cell(3, 1, "EUROPE")
        imports.set_cell(4, 1, "Austria (unit note q)", fill_rgb=GREEN)
        imports.set_cell(4, 2, 7.5, fill_rgb=GREEN)
        imports.set_cell(4, 3, 0.2, fill_rgb=YELLOW)
        sheets.append(imports)

    write_workbook(path, WorkbookData(sheets=sheets))


def _build_numeric_year_workbook(path: Path) -> None:
    """Build a workbook whose year header is stored as a numeric Excel value."""
    area = SheetData(name="AREA")
    area.set_cell(1, 2, 1900.0)
    area.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL", fill_rgb=YELLOW)
    area.set_cell(3, 1, "EUROPE.", fill_rgb=ORANGE)
    area.set_cell(4, 1, "Austria (r)", fill_rgb=GREEN)
    area.set_cell(4, 2, 12, fill_rgb=GREEN)
    write_workbook(path, WorkbookData(sheets=[area]))


def _build_ocr_year_values_workbook(path: Path) -> None:
    """Build a workbook with OCR-like i/o substitutions in year columns."""
    area = SheetData(name="AREA")
    area.set_cell(1, 2, "1900")
    area.set_cell(1, 3, "1901")
    area.set_cell(1, 4, "1902")
    area.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL")
    area.set_cell(3, 1, "EUROPE")
    area.set_cell(4, 1, "Austria")
    area.set_cell(4, 2, "IoiO")
    area.set_cell(4, 3, "bio")
    area.set_cell(4, 4, "1.2.3a")
    write_workbook(path, WorkbookData(sheets=[area]))


def _build_non_year_header_workbook(path: Path) -> None:
    """Build a workbook with one non-year header (contains no digits)."""
    area = SheetData(name="AREA")
    area.set_cell(1, 2, "Total")
    area.set_cell(1, 3, "1922")
    area.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL")
    area.set_cell(3, 1, "EUROPE")
    area.set_cell(4, 1, "Austria")
    area.set_cell(4, 2, 99)
    area.set_cell(4, 3, 12)
    write_workbook(path, WorkbookData(sheets=[area]))


def _build_multi_continent_workbook(path: Path) -> None:
    """Build a workbook spanning multiple continents to test spacer-row logic."""
    area = SheetData(name="AREA")
    area.set_cell(1, 2, "1900")
    area.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL")
    area.set_cell(3, 1, "EUROPE.")
    area.set_cell(4, 1, "Austria", fill_rgb=GREEN)
    area.set_cell(4, 2, 12, fill_rgb=GREEN)
    area.set_cell(5, 1, "ASIE.")
    area.set_cell(6, 1, "Japan", fill_rgb=YELLOW)
    area.set_cell(6, 2, 8, fill_rgb=YELLOW)
    write_workbook(path, WorkbookData(sheets=[area]))


def _build_world_total_country_workbook(path: Path) -> None:
    """Build a workbook where a totals-like country row should map to WORLD."""
    area = SheetData(name="AREA")
    area.set_cell(1, 2, "1900")
    area.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL")
    area.set_cell(3, 1, "EUROPE")
    area.set_cell(4, 1, "totaux non compris l'u r s s generaux")
    area.set_cell(4, 2, 25)
    write_workbook(path, WorkbookData(sheets=[area]))


def _build_document_variable_mapping_workbook(path: Path) -> None:
    """Build the document-variable unit mapping workbook used by config tests."""
    mapping = SheetData(name="mapping")
    mapping.set_cell(1, 1, "document")
    mapping.set_cell(1, 2, "variable")
    mapping.set_cell(1, 3, "unit")
    mapping.set_cell(2, 1, "reviewed_10_20wheat")
    mapping.set_cell(2, 2, "AREA")
    mapping.set_cell(2, 3, "hectares")
    mapping.set_cell(3, 1, "reviewed_10_20wheat")
    mapping.set_cell(3, 2, "PRODUCTION")
    mapping.set_cell(3, 3, "tonnes")
    write_workbook(path, WorkbookData(sheets=[mapping]))


def _standard_config_lines(
    document_name: str, *, unit_mode: str = "standard"
) -> list[str]:
    """Return common test config lines for a single workbook category."""
    return [
        f"unit_mode: {unit_mode}",
        "document_categories:",
        f"  {document_name}: 1",
    ]


def test_geography_detection_handles_known_ocr_and_accent_variants() -> None:
    assert _is_continent_row("AMÉR DU NORD ET AMÉR CENTRALE")
    assert _is_continent_row("Amérique méridionale")
    assert _is_continent_row("OCRANIE.")
    assert _is_continent_row("AUSTRALIE")
    assert _is_hemisphere_row("HÉMISHPÈRE SEPTENTRIONAL")
    assert _is_hemisphere_row("HÊMISPHÊRE SEPTENTRIONAL")
    assert _is_hemisphere_row("HÉMISPHÈRE SUDAFRIQUE.")
    assert not _is_continent_row("Canada")
    assert not _is_hemisphere_row("Canada")


def test_transform_workbook_assigns_units_from_rules_and_preserves_notes(
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "r_iia_trade_1950_3_5_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_source_workbook(source_path)

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_3_5_wheat"),
    )

    transform_workbook(source_path, output_path, config=load_config(config_path))

    result = read_workbook(output_path)
    assert [sheet.name for sheet in result.sheets] == ["area", "production"]

    area = result.sheets[0]
    assert [area.get_cell(1, idx).value for idx in range(1, 8)] == [
        "hemisphere",
        "continent",
        "country",
        "unit",
        "footnotes",
        "1909-1913",
        "1922",
    ]
    assert area.get_cell(2, 1).value == "HÉMISPHÈRE SEPTENTRIONAL"
    assert area.get_cell(2, 2).value == "EUROPE"
    assert area.get_cell(2, 3).value == "Belgique-Luxembourg"
    assert area.get_cell(2, 4).value == ""
    assert area.get_cell(2, 5).value == "reexports; special case"
    assert area.get_cell(2, 6).value == 17268
    assert area.get_cell(2, 7).value == 11887
    assert area.get_cell(2, 1).fill_rgb is None
    assert area.get_cell(2, 2).fill_rgb is None
    assert area.get_cell(2, 3).fill_rgb == GREEN
    assert area.get_cell(2, 4).fill_rgb is None
    assert area.get_cell(2, 5).fill_rgb is None
    assert area.get_cell(2, 6).fill_rgb == YELLOW
    assert area.get_cell(2, 7).fill_rgb == ORANGE

    production = result.sheets[1]
    assert production.get_cell(2, 1).value == "Hemisphère méridional"
    assert production.get_cell(2, 2).value == "Amérique"
    assert production.get_cell(2, 3).value == "Canada"
    assert production.get_cell(2, 4).value == ""
    assert production.get_cell(2, 6).value == 194876
    assert production.get_cell(2, 7).value == 315569


def test_transform_workbook_preserves_group_colors_and_normalizes_numeric_year_headers(
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "r_iia_trade_1950_1_1_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_numeric_year_workbook(source_path)

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_1_1_wheat"),
    )

    transform_workbook(source_path, output_path, config=load_config(config_path))

    result = read_workbook(output_path)
    area = result.sheets[0]
    assert area.get_cell(1, 6).value == "1900"
    assert area.get_cell(2, 1).fill_rgb == YELLOW
    assert area.get_cell(2, 2).fill_rgb == ORANGE
    assert area.get_cell(2, 3).fill_rgb == GREEN
    assert area.get_cell(2, 4).fill_rgb is None
    assert area.get_cell(2, 5).fill_rgb is None
    assert area.get_cell(2, 5).value == "reexports"


def test_transform_workbook_normalizes_ocr_like_year_values(tmp_path: Path) -> None:
    source_path = tmp_path / "r_iia_trade_1950_1_1_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_ocr_year_values_workbook(source_path)

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_1_1_wheat"),
    )

    transform_workbook(source_path, output_path, config=load_config(config_path))

    result = read_workbook(output_path)
    area = result.sheets[0]
    assert area.get_cell(2, 6).value == "1010"
    assert area.get_cell(2, 7).value == "10"
    assert area.get_cell(2, 8).value == "12.3"


def test_transform_workbook_tracks_docs_with_non_year_headers(tmp_path: Path) -> None:
    source_path = tmp_path / "r_iia_trade_1950_3_5_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_non_year_header_workbook(source_path)
    non_year_header_document_index = NonYearHeaderDocumentIndex()

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_3_5_wheat"),
    )

    transform_workbook(
        source_path,
        output_path,
        config=load_config(config_path),
        non_year_header_document_index=non_year_header_document_index,
    )

    result = read_workbook(output_path)
    area = result.sheets[0]
    assert [area.get_cell(1, idx).value for idx in range(1, 7)] == [
        "hemisphere",
        "continent",
        "country",
        "unit",
        "footnotes",
        "1922",
    ]
    assert non_year_header_document_index.documents == {output_path.name}


def test_transform_workbook_inserts_blank_row_before_each_new_continent(
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "r_iia_trade_1950_1_1_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_multi_continent_workbook(source_path)

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_1_1_wheat"),
    )

    transform_workbook(source_path, output_path, config=load_config(config_path))

    result = read_workbook(output_path)
    area = result.sheets[0]
    assert area.get_cell(2, 1).value is None
    assert area.get_cell(2, 2).value is None
    assert area.get_cell(2, 3).value is None
    assert area.get_cell(3, 2).value == "EUROPE"
    assert area.get_cell(3, 3).value == "Austria"
    assert area.get_cell(4, 1).value is None
    assert area.get_cell(4, 2).value is None
    assert area.get_cell(4, 3).value is None
    assert area.get_cell(5, 2).value == "ASIE"
    assert area.get_cell(5, 3).value == "Japan"
    assert area.get_cell(5, 3).fill_rgb == YELLOW


def test_transform_workbook_maps_world_total_country_to_world_continent(
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "r_iia_trade_1950_1_1_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_world_total_country_workbook(source_path)

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_1_1_wheat"),
    )

    transform_workbook(source_path, output_path, config=load_config(config_path))

    result = read_workbook(output_path)
    area = result.sheets[0]
    assert area.get_cell(2, 2).value == "WORLD"
    assert area.get_cell(2, 3).value == "totaux non compris l'u r s s generaux"


def test_transform_workbook_collects_unique_geography_labels(tmp_path: Path) -> None:
    source_path = tmp_path / "r_iia_trade_1950_3_5_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_source_workbook(source_path, include_imports=True)
    geography_index = GeographyIndex()

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_3_5_wheat"),
    )

    transform_workbook(
        source_path,
        output_path,
        config=load_config(config_path),
        geography_index=geography_index,
    )
    assert geography_index.hemispheres == {
        "HÉMISPHÈRE SEPTENTRIONAL",
        "Hemisphère méridional",
    }
    assert geography_index.continents == {"EUROPE", "Amérique"}
    assert geography_index.countries == {
        "Austria",
        "Belgique-Luxembourg",
        "Canada",
        "Germany",
    }

    geography_index.write_split_txts(tmp_path)
    assert (tmp_path / "unique_hemisphere_values.txt").read_text(
        encoding="utf-8"
    ) == "\n".join(
        ["[hemispheres]", "Hemisphère méridional", "HÉMISPHÈRE SEPTENTRIONAL", ""]
    )
    assert (tmp_path / "unique_continent_values.txt").read_text(
        encoding="utf-8"
    ) == "\n".join(["[continents]", "Amérique", "EUROPE", ""])
    assert (tmp_path / "unique_country_values.txt").read_text(
        encoding="utf-8"
    ) == "\n".join(
        ["[countries]", "Austria", "Belgique-Luxembourg", "Canada", "Germany", ""]
    )


def test_transform_workbook_assigns_units_from_document_variable_mapping(
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "reviewed_10_20wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    mapping_path = tmp_path / "document_variable_unit_mapping.xlsx"
    _build_source_workbook(source_path)
    _build_document_variable_mapping_workbook(mapping_path)

    config_path = _write_config(
        tmp_path / "config.yml",
        [
            "unit_mode: standard",
            f"document_variable_unit_mapping_file: {mapping_path.as_posix()}",
        ],
    )

    transform_workbook(source_path, output_path, config=load_config(config_path))

    result = read_workbook(output_path)
    assert result.sheets[0].name == "area"
    assert result.sheets[0].get_cell(2, 4).value == "hectares"
    assert result.sheets[1].name == "production"
    assert result.sheets[1].get_cell(2, 4).value == "tonnes"


def test_load_config_creates_mapping_template_when_file_is_missing(
    tmp_path: Path,
) -> None:
    mapping_path = tmp_path / "data" / "document_variable_unit_mapping.xlsx"
    config_path = _write_config(
        tmp_path / "config.yml",
        [
            "unit_mode: standard",
            f"document_variable_unit_mapping_file: {mapping_path.as_posix()}",
        ],
    )

    loaded_config = load_config(config_path)

    assert loaded_config.document_variable_units == {}
    assert mapping_path.exists()
    created_mapping = read_workbook(mapping_path)
    assert created_mapping.sheets[0].get_cell(1, 1).value == "document"
    assert created_mapping.sheets[0].get_cell(1, 2).value == "variable"
    assert created_mapping.sheets[0].get_cell(1, 3).value == "unit"


def test_extract_footnotes_normalizes_and_deduplicates_index_output(
    tmp_path: Path,
) -> None:
    footnote_index = FootnoteIndex()
    footnote_index.add_footnotes(
        _extract_footnotes("Austria ( unit note q ) (special case.)")
    )
    footnote_index.add_footnotes(_extract_footnotes("Belgium (special case) (r)"))

    index_path = tmp_path / "unique_footnotes.txt"
    footnote_index.write_txt(index_path)

    assert index_path.read_text(encoding="utf-8") == "\n".join(
        ["[footnotes]", "reexports", "special case", "unit note q", ""]
    )


def test_document_index_writes_sorted_unique_documents(tmp_path: Path) -> None:
    document_index = DocumentIndex()
    document_index.add_document("b.xlsx")
    document_index.add_document("a.xlsx")
    document_index.add_document("b.xlsx")

    index_path = tmp_path / "final_docs.txt"
    document_index.write_txt(index_path)

    assert index_path.read_text(encoding="utf-8") == "\n".join(
        ["[documents]", "a.xlsx", "b.xlsx", ""]
    )


def test_transform_workbook_tracks_documents_with_unit_related_footnotes(
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "r_iia_trade_1950_3_5_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_source_workbook(source_path, include_imports=True)
    unit_footnote_document_index = UnitFootnoteDocumentIndex()

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_3_5_wheat"),
    )

    transform_workbook(
        source_path,
        output_path,
        config=load_config(config_path),
        unit_footnote_document_index=unit_footnote_document_index,
    )

    assert unit_footnote_document_index.documents == {"standardized.xlsx"}
    txt_path = unit_footnote_document_index.write_txt(tmp_path / "unit_footnotes.txt")
    assert txt_path.read_text(encoding="utf-8") == "[documents]\nstandardized.xlsx\n"


def test_transform_workbook_tracks_documents_with_countries_missing_units(
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "r_iia_trade_1950_3_5_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_source_workbook(source_path, include_imports=True)
    missing_unit_country_document_index = MissingUnitCountryDocumentIndex()

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_3_5_wheat"),
    )

    transform_workbook(
        source_path,
        output_path,
        config=load_config(config_path),
        missing_unit_country_document_index=missing_unit_country_document_index,
    )

    assert missing_unit_country_document_index.documents == {"standardized.xlsx"}
    txt_path = missing_unit_country_document_index.write_txt(
        tmp_path / "missing_country_units.txt"
    )
    assert txt_path.read_text(encoding="utf-8") == (
        "[documents]\n"
        "Original Excel Name\tExcel Sheet Names\n"
    )


def test_transform_workbook_treats_na_like_units_as_missing(
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "reviewed_10_20wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_source_workbook(source_path)
    missing_unit_country_document_index = MissingUnitCountryDocumentIndex()

    config_path = _write_config(
        tmp_path / "config.yml",
        [
            "unit_mode: standard",
            "unit_overrides:",
            "  area: N/A",
            "  production: tonnes",
        ],
    )

    transform_workbook(
        source_path,
        output_path,
        config=load_config(config_path),
        missing_unit_country_document_index=missing_unit_country_document_index,
    )

    result = read_workbook(output_path)
    area = result.sheets[0]
    production = result.sheets[1]
    assert area.get_cell(2, 4).value == ""
    assert production.get_cell(2, 4).value == "tonnes"
    assert missing_unit_country_document_index.documents == {"standardized.xlsx"}


def test_transform_workbook_indexes_missing_units_even_with_blank_country_cell(
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "r_iia_trade_1950_3_5_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    area = SheetData(name="AREA")
    area.set_cell(1, 2, "1900")
    area.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL")
    area.set_cell(3, 1, "EUROPE")
    area.set_cell(4, 1, "(r)")
    area.set_cell(4, 2, 12)
    write_workbook(source_path, WorkbookData(sheets=[area]))
    missing_unit_country_document_index = MissingUnitCountryDocumentIndex()

    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("r_iia_trade_1950_3_5_wheat"),
    )

    transform_workbook(
        source_path,
        output_path,
        config=load_config(config_path),
        missing_unit_country_document_index=missing_unit_country_document_index,
    )

    result = read_workbook(output_path)
    transformed_area = result.sheets[0]
    assert transformed_area.get_cell(2, 3).value == ""
    assert transformed_area.get_cell(2, 4).value == ""
    assert missing_unit_country_document_index.documents == {"standardized.xlsx"}


def test_product_index_writes_sorted_unique_products(tmp_path: Path) -> None:
    product_index = ProductIndex()
    product_index.add_product("rice")
    product_index.add_product("wheat")
    product_index.add_product("rice")

    index_path = tmp_path / "unique_product_values.txt"
    product_index.write_txt(index_path)

    assert index_path.read_text(encoding="utf-8") == "\n".join(
        ["[products]", "rice", "wheat", ""]
    )


def test_transform_workbook_supports_inputs_mode_and_harmonized_output_names(
    tmp_path: Path,
) -> None:
    source_dir = tmp_path / "raw_inputs" / "trade" / "extracted_pages_1938_39"
    source_dir.mkdir(parents=True)
    source_path = source_dir / "reviewed_466_475arrozimp_exp.xlsx"
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    _build_source_workbook(source_path, include_imports=True)

    config_path = _write_config(
        tmp_path / "config.yml",
        [
            "unit_mode: inputs",
            "document_categories:",
            "  reviewed_466_475arrozimp_exp: 2",
            "product_translations:",
            "  arroz: rice",
        ],
    )

    config = load_config(config_path)
    output_path = output_dir / f"{config.canonical_name_for_document(source_path)}.xlsx"
    transform_workbook(source_path, output_path, config=config)

    assert output_path.name == "r_iia_trade_1938_466_475_rice.xlsx"
    result = read_workbook(output_path)
    imports = result.sheets[2]
    assert imports.name == "imports"
    assert imports.get_cell(2, 3).value == "Austria"
    assert imports.get_cell(2, 4).value == ""
    assert imports.get_cell(2, 5).value == "unit note q"
    assert imports.get_cell(2, 6).value == 7.5
    assert imports.get_cell(2, 7).value == 0.2


def test_canonical_document_name_auto_translates_unknown_products(
    monkeypatch: Any,
) -> None:
    from iia_excel_reorg.utils import naming

    monkeypatch.setattr(naming, "_auto_translate_product", lambda value: "cocoa beans")
    path = Path("raw_inputs/trade/extracted_pages_1938_39/reviewed_12_13cacaoimp.xlsx")

    assert canonical_document_name(path) == "r_iia_trade_1938_12_13_cocoa_beans"


def test_canonical_document_name_applies_alias_before_translation() -> None:
    path = Path("raw_inputs/trade/extracted_pages_1938_39/reviewed_12_13teaimp.xlsx")

    assert (
        canonical_document_name(
            path,
            product_translations={"te": "tea"},
            product_aliases={"tea": "te"},
        )
        == "r_iia_trade_1938_12_13_tea"
    )


def test_naming_and_unit_rules_cover_reviewed_documents() -> None:
    path = Path(
        "raw_inputs/trade/extracted_pages_1938_39/"
        "reviewed_239_239azucar_caña_brutaprod.xlsx"
    )
    assert infer_yearbook_metadata(path) == {
        "agency": "iia",
        "yearbook": "trade",
        "year": "1938",
    }
    assert extract_source_product(path) == "azucar cana bruta"
    assert (
        canonical_document_name(
            path,
            product_translations={"azucar cana bruta": "raw cane sugar"},
        )
        == "r_iia_trade_1938_239_239_raw_cane_sugar"
    )
    assert assign_unit("imports", "te", 1) == "__NA_UNIT__"
    assert assign_unit("imports", "te", 2) == "__NA_UNIT__"
    assert assign_unit("production", "vino", 1) == "__NA_UNIT__"
    assert assign_unit("production", "huevos", 2) == "__NA_UNIT__"
    assert assign_unit("livestock", "whatever", 1) == "__NA_UNIT__"
    assert assign_unit("production", "whatever", None) == "__NA_UNIT__"


def test_canonical_document_name_translates_multiword_reviewed_product_at_end(
    monkeypatch: Any,
) -> None:
    from iia_excel_reorg.utils import naming

    monkeypatch.setattr(
        naming,
        "_auto_translate_product",
        lambda value: "raw beet sugar",
    )
    path = Path(
        "raw_inputs/trade/extracted_pages_1938_39/"
        "reviewed_238_238azucar_remolacha_brutaprod.xlsx"
    )

    assert extract_source_product(path) == "azucar remolacha bruta"
    assert canonical_document_name(path) == "r_iia_trade_1938_238_238_raw_beet_sugar"


def test_load_config_parses_rule_based_yaml(tmp_path: Path) -> None:
    config_path = _write_config(
        tmp_path / "units.yml",
        [
            "unit_mode: standard",
            "document_categories:",
            "  reviewed_466_475arrozimp_exp: 1",
            "product_aliases:",
            "  tea: te",
            "product_translations:",
            "  arroz: rice",
            "unit_overrides:",
            "  imports: tonnes",
            "include_sheets:",
            "  - AREA",
            "  - PRODUCTION",
        ],
    )

    config = load_config(config_path)
    assert config.unit_mode == "standard"
    assert config.document_categories["reviewed_466_475arrozimp_exp"] == 1
    assert config.product_aliases["tea"] == "te"
    assert config.product_translations["arroz"] == "rice"
    assert config.unit_overrides["imports"] == "tonnes"
    assert config.include_sheets == ["AREA", "PRODUCTION"]


def test_workbook_config_canonical_name_uses_product_aliases() -> None:
    config = WorkbookConfig(
        product_aliases={"tea": "te"},
        product_translations={"te": "tea"},
    )
    path = Path("raw_inputs/trade/extracted_pages_1938_39/reviewed_12_13teaimp.xlsx")

    assert config.canonical_name_for_document(path) == "r_iia_trade_1938_12_13_tea"


def test_run_project_bootstraps_deep_translator(monkeypatch: Any) -> None:
    spec = importlib.util.spec_from_file_location(
        "run_project",
        Path(__file__).resolve().parents[2] / "run_project.py",
    )
    assert spec is not None and spec.loader is not None
    run_project = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(run_project)

    commands: list[list[str]] = []

    monkeypatch.setattr(
        importlib.util,
        "find_spec",
        lambda name: None if name == "deep_translator" else object(),
    )
    monkeypatch.setattr(
        run_project.subprocess,
        "check_call",
        lambda command: commands.append(command) or 0,
    )

    run_project._ensure_translation_dependency()

    assert commands == [
        [
            run_project.sys.executable,
            "-m",
            "pip",
            "install",
            "deep-translator>=1.11.4",
        ]
    ]


def test_compute_output_subdir_with_extracted_pages_and_subfolder() -> None:
    path = Path(
        "inputs/reviewed_iia/extracted_pages_1929_30/crops/reviewed_1_2_wheat.xlsx"
    )
    assert _compute_output_subdir(path) == Path(
        "iia_extracted_pages_1929/iia_crops_1929"
    )


def test_compute_output_subdir_with_extracted_pages_no_subfolder() -> None:
    path = Path("raw_inputs/trade/extracted_pages_1938_39/reviewed_466_475arroz.xlsx")
    assert _compute_output_subdir(path) == Path(
        "iia_extracted_pages_1938/iia_trade_1938"
    )


def test_compute_output_subdir_with_deep_nesting() -> None:
    path = Path(
        "raw_inputs/area and production/multiple product/extracted_pages_1933_34/wb.xlsx"
    )
    assert _compute_output_subdir(path) == Path(
        "iia_extracted_pages_1933/iia_multiple_product_1933"
    )


def test_compute_output_subdir_without_extracted_pages() -> None:
    path = Path("some/other/dir/workbook.xlsx")
    assert _compute_output_subdir(path) == Path(".")


def test_duplicate_original_document_index_lists_only_duplicate_names(
    tmp_path: Path,
) -> None:
    root = tmp_path / "inputs"
    first = root / "reviewed_iia" / "extracted_pages_1929_30" / "crops" / "same.xlsx"
    second = root / "reviewed_iia" / "extracted_pages_1938_39" / "trade" / "same.xlsx"
    third = root / "reviewed_iia" / "extracted_pages_1938_39" / "trade" / "other.xlsx"
    first.parent.mkdir(parents=True)
    second.parent.mkdir(parents=True, exist_ok=True)
    first.write_text("a", encoding="utf-8")
    second.write_text("b", encoding="utf-8")
    third.write_text("c", encoding="utf-8")

    index = DuplicateOriginalDocumentIndex()
    index.add_document(first, root=root)
    index.add_document(second, root=root)
    index.add_document(third, root=root)

    output_path = tmp_path / "duplicated_original_documents.txt"
    index.write_txt(output_path)

    assert output_path.read_text(encoding="utf-8") == "\n".join(
        [
            "[documents]",
            "same.xlsx",
            "  reviewed_iia/extracted_pages_1929_30/crops/same.xlsx",
            "  reviewed_iia/extracted_pages_1938_39/trade/same.xlsx",
            "",
        ]
    )


def test_iter_workbooks_structured_builds_correct_hierarchy(tmp_path: Path) -> None:
    crops_dir = tmp_path / "reviewed_iia" / "extracted_pages_1929_30" / "crops"
    trade_dir = tmp_path / "reviewed_iia" / "extracted_pages_1938_39" / "trade"
    crops_dir.mkdir(parents=True)
    trade_dir.mkdir(parents=True)

    workbook_one = crops_dir / "reviewed_1_2_wheat.xlsx"
    workbook_two = trade_dir / "reviewed_3_4_rice.xlsx"
    _build_source_workbook(workbook_one)
    _build_source_workbook(workbook_two)

    entries = _iter_workbooks_structured(tmp_path)
    paths_and_subdirs = {entry[0].name: entry[1] for entry in entries}
    assert paths_and_subdirs["reviewed_1_2_wheat.xlsx"] == Path(
        "iia_extracted_pages_1929/iia_crops_1929"
    )
    assert paths_and_subdirs["reviewed_3_4_rice.xlsx"] == Path(
        "iia_extracted_pages_1938/iia_trade_1938"
    )


def test_cli_main_creates_structured_output(
    tmp_path: Path, monkeypatch: Any, capsys: Any
) -> None:
    """End-to-end: main() populates the structured output hierarchy."""
    import iia_excel_reorg.cli as cli

    monkeypatch.setattr(cli, "LISTS_DIR", tmp_path / "lists")
    main = cli.main

    crops_dir = (
        tmp_path / "inputs" / "reviewed_iia" / "extracted_pages_1929_30" / "crops"
    )
    crops_dir.mkdir(parents=True)
    source = crops_dir / "reviewed_1_2_wheat.xlsx"
    _build_source_workbook(source, include_imports=True)

    output_root = tmp_path / "outputs"
    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("reviewed_1_2_wheat"),
    )

    import sys

    original_argv = sys.argv
    try:
        sys.argv = [
            "iia-excel-reorg",
            str(tmp_path / "inputs"),
            str(output_root),
            "--config",
            str(config_path),
        ]
        main()
    finally:
        sys.argv = original_argv

    output_subdir = output_root / "iia_extracted_pages_1929" / "iia_crops_1929"
    assert output_subdir.is_dir()
    transformed_files = list(output_subdir.glob("*.xlsx"))
    assert len(transformed_files) == 1
    lists_dir = tmp_path / "lists"
    assert (lists_dir / "unique_hemisphere_values.txt").is_file()
    assert (lists_dir / "unique_continent_values.txt").is_file()
    assert (lists_dir / "unique_country_values.txt").is_file()
    assert (lists_dir / "unique_product_values.txt").is_file()
    assert (lists_dir / "final_docs_with_missing_country_units.txt").is_file()
    assert (lists_dir / "final_docs_with_extra_non_year_columns.txt").is_file()
    final_docs_path = lists_dir / "final_docs.txt"
    assert final_docs_path.read_text(encoding="utf-8") == (
        "[documents]\n" f"{transformed_files[0].name}\n"
    )
    unit_docs_path = lists_dir / "final_docs_with_unit_footnotes.txt"
    assert unit_docs_path.read_text(encoding="utf-8") == (
        "[documents]\n" f"{transformed_files[0].name}\n"
    )
    missing_units_docs_path = lists_dir / "final_docs_with_missing_country_units.txt"
    assert missing_units_docs_path.read_text(encoding="utf-8") == (
        "[documents]\n"
        "Original Excel Name\tExcel Sheet Names\n"
        f"{source.name}\tArea; Production; Imports\n"
    )
    non_year_header_docs_path = lists_dir / "final_docs_with_extra_non_year_columns.txt"
    assert non_year_header_docs_path.read_text(encoding="utf-8") == "[documents]\n"
    duplicate_original_docs_path = lists_dir / "duplicated_original_documents.txt"
    assert duplicate_original_docs_path.read_text(encoding="utf-8") == "[documents]\n"
    captured = capsys.readouterr()
    assert "Generating txt lists" in captured.out


def test_extract_sheet_names_handles_unreadable_workbook(tmp_path: Path) -> None:
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not-a-workbook")
    assert _extract_sheet_names(broken) == "Could not read sheets"


def test_ensure_workspace_creates_missing_input_and_output_dirs(tmp_path: Path) -> None:
    input_dir = tmp_path / "raw_inputs"
    output_dir = tmp_path / "transformed"
    _ensure_workspace(input_dir, output_dir)
    assert input_dir.is_dir()
    assert output_dir.is_dir()


def test_cli_main_lists_duplicate_original_documents(
    tmp_path: Path, monkeypatch: Any
) -> None:
    import iia_excel_reorg.cli as cli

    monkeypatch.setattr(cli, "LISTS_DIR", tmp_path / "lists")
    main = cli.main

    crops_dir = (
        tmp_path / "inputs" / "reviewed_iia" / "extracted_pages_1929_30" / "crops"
    )
    trade_dir = (
        tmp_path / "inputs" / "reviewed_iia" / "extracted_pages_1938_39" / "trade"
    )
    crops_dir.mkdir(parents=True)
    trade_dir.mkdir(parents=True)
    shared_name = "reviewed_1_2_wheat.xlsx"
    _build_source_workbook(crops_dir / shared_name)
    _build_source_workbook(trade_dir / shared_name)

    output_root = tmp_path / "outputs"
    config_path = _write_config(
        tmp_path / "config.yml",
        _standard_config_lines("reviewed_1_2_wheat"),
    )

    import sys

    original_argv = sys.argv
    try:
        sys.argv = [
            "iia-excel-reorg",
            str(tmp_path / "inputs"),
            str(output_root),
            "--config",
            str(config_path),
        ]
        main()
    finally:
        sys.argv = original_argv

    duplicate_original_docs_path = tmp_path / "lists" / "duplicated_original_documents.txt"
    assert duplicate_original_docs_path.read_text(encoding="utf-8") == "\n".join(
        [
            "[documents]",
            shared_name,
            "  reviewed_iia/extracted_pages_1929_30/crops/reviewed_1_2_wheat.xlsx",
            "  reviewed_iia/extracted_pages_1938_39/trade/reviewed_1_2_wheat.xlsx",
            "",
        ]
    )
