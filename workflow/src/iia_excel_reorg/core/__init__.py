"""Core transformation business logic for the workbook reorganization pipeline.

Sub-modules
-----------
transformer
    The main :func:`~core.transformer.transform_workbook` function plus the
    :class:`~core.transformer.GeographyIndex` and
    :class:`~core.transformer.ProductIndex` accumulators.
preprocessor
    The pre-processing :func:`~core.preprocessor.process_workbook` function
    used by the ``iia-prepare`` pipeline.
"""

from .preprocessor import process_workbook
from .transformer import (
    GeographyIndex,
    ProductIndex,
    TransformationError,
    transform_workbook,
)

__all__ = [
    "GeographyIndex",
    "ProductIndex",
    "TransformationError",
    "process_workbook",
    "transform_workbook",
]
