"""Utilities for reorganizing historical Excel workbooks."""

from .config import WorkbookConfig, load_config
from .footnote_pipeline import (
    PIPELINE_NAME as FOOTNOTE_PIPELINE_NAME,
    apply_mapping_in_place,
    collect_unique_footnotes,
    generate_mapping_template,
)
from .preprocess_pipeline import main as preprocess_main
from .utils.naming import canonical_document_name, extract_source_product, infer_yearbook_metadata, sanitize_name
from .core.preprocessor import process_workbook
from .core.transformer import transform_workbook
from .services.units import assign_unit

__all__ = [
    "WorkbookConfig",
    "assign_unit",
    "apply_mapping_in_place",
    "canonical_document_name",
    "collect_unique_footnotes",
    "extract_source_product",
    "FOOTNOTE_PIPELINE_NAME",
    "generate_mapping_template",
    "infer_yearbook_metadata",
    "load_config",
    "preprocess_main",
    "process_workbook",
    "sanitize_name",
    "transform_workbook",
]
