"""Setuptools configuration for the ``iia-excel-reorg`` package."""

from setuptools import find_packages, setup

setup(
    name="iia-excel-reorg",
    version="0.1.0",
    description=(
        "Workflow to reorganize historical Excel workbooks into a standardized "
        "structure."
    ),
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    python_requires=">=3.11",
    install_requires=[
        "deep-translator>=1.11.4",
        "numpy>=1.24",
        "openpyxl>=3.1.0",
        "pandas>=2.0",
    ],
    extras_require={
        "dev": ["pytest>=8.0.0"],
        "fast": ["python-calamine>=0.2.0"],
    },
    entry_points={
        "console_scripts": [
            "iia-excel-reorg=iia_excel_reorg.cli:main",
            "iia-footnote-harmonizer=iia_excel_reorg.footnote_pipeline:main",
            "iia-prepare=iia_excel_reorg.preprocess_pipeline:main",
        ]
    },
)
