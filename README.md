# IIA Excel Reorganization Workflow

This repository contains a complete Python workflow for reorganizing historical Excel workbooks into a standardized workbook structure.

## Workflow summary

The diagram below shows the complete pipeline from raw Excel scans to
structured, analysis-ready workbooks.

```
┌─────────────────────────────────────────────────────────────────────┐
│  1. INPUT  –  place historical .xlsx scans in data/transform/00_input/  │
│                                                                     │
│  data/transform/00_input/                                           │
│  ├── trade/extracted_pages_1938_39/reviewed_466_475arrozimp.xlsx    │
│  ├── livestock/extracted_pages_1933_34/reviewed_12_15cattle.xlsx    │
│  └── area and production/                                           │
│      └── multiple product/extracted_pages_1939_45/reviewed_…xlsx   │
└────────────────────────┬────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────┐
│  2. CONFIGURE  –  create (or reuse) a YAML config file              │
│                                                                     │
│  workflow/config/example.units.yml defines:                         │
│  • unit_mode       standard | inputs                                │
│  • document_categories   stem → category number                     │
│  • product_aliases       source product → canonical product         │
│  • product_translations  canonical product → English slug           │
│  • unit_overrides        sheet (or doc:sheet) → explicit unit       │
│  • include_sheets        optional list of sheets to process         │
└────────────────────────┬────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────┐
│  3. RUN  –  execute the CLI from the project root                   │
│                                                                     │
│  iia-excel-reorg --config workflow/config/example.units.yml         │
│  (defaults: input = "data/transform/00_input/",                    │
│             output = "data/transform/01_output/")                  │
└────────────────────────┬────────────────────────────────────────────┘
                         │
                  per workbook file
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────┐
│  4. DISCOVER  –  cli._iter_workbooks_structured()                   │
│                                                                     │
│  Recursively finds every *.xlsx / *.xlsm file and maps it to an    │
│  output subdirectory by inspecting the extracted_pages_YYYY_YY      │
│  segment in its path:                                               │
│  • subfolder below extracted_pages_* (e.g. crops/) → iia_crops_*   │
│  • otherwise: folder above extracted_pages_* (e.g. trade/)         │
│    → iia_trade_*                                                    │
└────────────────────────┬────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────┐
│  5. NAME  –  naming.canonical_document_name()                       │
│                                                                     │
│  reviewed_466_475arrozimp_exp.xlsx                                  │
│        │                                                            │
│        ├─ strip "reviewed_" → r_                                    │
│        ├─ infer agency (iia) + yearbook + year from folder path     │
│        ├─ extract product body, strip known suffixes (imp, exp, …)  │
│        ├─ apply product_aliases then product_translations           │
│        └─ assemble: r_iia_trade_1938_466_475_rice                   │
└────────────────────────┬────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────┐
│  6. TRANSFORM  –  transformer.transform_workbook()                  │
│                                                                     │
│  For every sheet in the source workbook:                            │
│  a) Read row 1 → extract year / period headers                      │
│  b) Walk remaining rows:                                            │
│     • HÉMISPHÈRE row  → update hemisphere state (not written)       │
│     • Continent row   → update continent state  (not written)       │
│     • Country row     → extract country name + footnotes from (…)   │
│  c) Assign unit via unit_rules.assign_unit(variable, product, cat)  │
│  d) Write output row:                                               │
│     hemisphere | continent | country | unit | footnotes | yr cols   │
│  e) Preserve source cell fill colours on every copied cell          │
│  f) Write sheet with name lowercased                                │
└────────────────────────┬────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────────┐
│  7. OUTPUT  –  results land in data/transform/01_output/           │
│                                                                     │
│  data/transform/01_output/                                         │
│  └── iia_extracted_pages_1938/                                      │
│      └── iia_trade_1938/                                            │
│          └── r_iia_trade_1938_466_475_rice.xlsx                     │
│              ┌──────────────────────────────────────────────┐       │
│              │ hemisphere│continent│country│unit│footnotes│…│       │
│              │ N         │ EUROPE  │France │1000q│        │…│       │
│              └──────────────────────────────────────────────┘       │
└─────────────────────────────────────────────────────────────────────┘
```

**In short:** drop your scanned Excel files into `data/transform/00_input/`,
point the CLI at your config, and the tool produces one clean, consistently
structured workbook per source file inside `data/transform/01_output/`,
organised into `iia_extracted_pages_YYYY/iia_{topic}_YYYY/` subdirectories.

---

## What the workflow does

The transformer currently supports the rules you specified:

- keeps every source sheet but writes it with the **same name in lowercase**
- converts hierarchical labels into explicit metadata columns:
  - `hemisphere`
  - `continent`
  - `country`
  - `unit`
  - `footnotes`
- preserves the year or period headers exactly as they appear in the source workbook
- extracts footnotes from country labels by taking every `(...)` segment, removing parentheses, and joining notes with `; `
- preserves source cell colors on copied data rows
- keeps repeated country/entity rows exactly as they appear in the source workbook
- assigns `unit` automatically from the document category, sheet variable, and source product using the rules you provided
- harmonizes reviewed document names into the canonical `r_iia_<yearbook>_<year>_<page_start>_<page_end>_<english_product>` format
- derives missing yearbook metadata from the folder path, for example `data/transform/00_input/trade/extracted_pages_1938_39/...` becomes `trade` and `1938`
- strips source suffixes such as `sup`, `prod`, `rend`, `imp`, `exp`, and `num` before translating the product portion of the document name
- supports both the standard iia unit rules and the special `inputs` unit rules
- includes automated tests and a GitHub Actions CI workflow
- writes text indexes under `data/transform/lists/`, including unique geography values and unique renamed product values

## Project structure

```text
.
├── .github/workflows/ci.yml
├── workflow/                    ← scripts + config
│   ├── pyproject.toml
│   ├── pytest.ini
│   ├── setup.py
│   ├── config/example.units.yml
│   ├── src/iia_excel_reorg/
│   │   ├── __init__.py
│   │   ├── cli.py
│   │   ├── config.py
│   │   ├── naming.py
│   │   ├── transformer.py
│   │   ├── unit_rules.py
│   │   └── xlsx_io.py
│   └── tests/test_transformer.py
└── data/                        ← input + output
    ├── preprocess/              ← pre-processing pipeline
    │   ├── 00_input/            ← raw excels needing human prep
    │   └── 01_output/           ← human-ready output
    ├── transform/               ← main transformation pipeline
    │   ├── 00_input/            ← place source Excel files here
    │   ├── 01_output/           ← transformed output written here
    │   └── lists/               ← generated text indexes
    └── footnote_mapping.xlsx    ← shared footnote mapping template
```

## Installation

> **Important:** always install the package through `pip`, not by running
> `workflow/setup.py` directly.  Running `python workflow/setup.py install`
> (or `%run workflow/setup.py` in IPython) will fail with `ModuleNotFoundError: No module named 'setuptools'`
> unless setuptools happens to be pre-installed in your environment.  `pip`
> handles the build-system bootstrapping for you automatically.

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e ./workflow[dev]
```

## One-click VS Code runner

If you want a single file that you can open in VS Code and launch with the
**Run** button, use the repository-root script:

```bash
python run_project.py
```

That wrapper adds `workflow/src` to `sys.path` and, when no arguments are
provided, automatically uses `workflow/config/example.units.yml`.  You can still
override the defaults by passing the same arguments supported by the main CLI:

```bash
python run_project.py "data/transform/00_input" "data/transform/01_output" --config workflow/config/example.units.yml
```

> **Folder name:** the input directory is called **`data/transform/00_input/`**.
> Make sure you place your Excel files under `data/transform/00_input/` at the
> project root.

## Configuration

Create a YAML file that supplies the metadata needed by the unit and naming rules.

Example:

```yaml
unit_mode: standard

document_categories:
  reviewed_239_239azucar_caña_brutaprod: 1
  reviewed_466_475arrozimp_exp: 2

product_aliases:
  tea: te

product_translations:
  azucar cana bruta: raw cane sugar
  arroz: rice

unit_overrides:
  imports: tonnes
```

### Config fields

- `unit_mode`: `standard` or `inputs`
- `document_categories`: maps each original or canonical document stem to its category number, which is required by the unit assignment logic
- `product_aliases`: optional mapping from extracted source products to the canonical product names used in the unit rules
- `product_translations`: optional mapping from extracted source products to the English product slug used in the harmonized output filename
- `unit_overrides`: optional explicit unit override by sheet name or by `document_stem:sheet_name`
- `include_sheets`: optional list of sheet names to process

You can use the example file at `workflow/config/example.units.yml` as a starting point.

> Note: the parser intentionally supports the simple YAML structure used for this project configuration.

## Folder structure

### Where to place the input files

Drop your Excel workbooks into the **`data/transform/00_input/`** folder at the root of this
project.  That directory is already created and pre-structured to mirror the
exact source layout described below.  Git ignores any `*.xlsx` / `*.xlsm` files
inside it, so your source files won't be accidentally committed.

```text
data/transform/00_input/           ← place your source Excel files here
├── area and production/
│   ├── multiple product/
│   │   ├── extracted_pages_1933_34/   ← drop .xlsx files in the appropriate year folder
│   │   ├── extracted_pages_1938_39/
│   │   └── extracted_pages_1939_45/
│   └── single product/
│       ├── extracted_pages_1909_21/
│       ├── extracted_pages_1925_26/
│       ├── extracted_pages_1929_30/
│       ├── extracted_pages_1933_34/
│       ├── extracted_pages_1938_39/
│       └── extracted_pages_1939_45/
├── country area/
│   ├── extracted_pages_1909_21/
│   ├── extracted_pages_1925_26/
│   ├── extracted_pages_1929_30/
│   ├── extracted_pages_1932_33/
│   ├── extracted_pages_1933_34/
│   └── extracted_pages_1938_39/
├── inputs/
│   ├── extracted_pages_1909_21/
│   ├── extracted_pages_1925_26/
│   ├── extracted_pages_1929_30/
│   ├── extracted_pages_1933_34/
│   ├── extracted_pages_1938_39/
│   └── extracted_pages_1939_45/
├── land use/
│   ├── extracted_pages_1909_21/
│   ├── extracted_pages_1925_26/
│   ├── extracted_pages_1929_30/
│   ├── extracted_pages_1932_33/
│   ├── extracted_pages_1933_34/
│   └── extracted_pages_1938_39/
├── livestock/
│   ├── extracted_pages_1909_21/
│   ├── extracted_pages_1925_26/
│   ├── extracted_pages_1929_30/
│   ├── extracted_pages_1933_34/
│   ├── extracted_pages_1938_39/
│   └── extracted_pages_1939_45/
└── trade/
    ├── extracted_pages_1909_21/
    ├── extracted_pages_1925_26/
    ├── extracted_pages_1929_30/
    ├── extracted_pages_1933_34/
    └── extracted_pages_1938_39/
```

Transformed workbooks are written to **`data/transform/01_output/`** inside
the `data/` folder.  That directory is pre-created and its generated `*.xlsx`
files are gitignored.

### Expected input structure (technical detail)

The tool recognises a convention built around directories named
`extracted_pages_YYYY_YY` (where `YYYY` is the full four-digit start year and
`YY` is the two-digit end year).

```text
<input_root>/
└── <topic>/                           ← e.g. "trade", "livestock"
    └── extracted_pages_1938_39/       ← year boundary; YYYY = 1938
        ├── reviewed_466_475arrozimp_exp.xlsx   ← directly inside
        └── crops/                              ← optional sub-topic folder
            └── reviewed_239_239azucar_caña_brutaprod.xlsx
```

A two-level topic hierarchy is also supported:

```text
<input_root>/
└── <main_topic>/                      ← e.g. "area and production"
    └── <sub_topic>/                   ← e.g. "multiple product"
        └── extracted_pages_1933_34/
            └── reviewed_...xlsx
```

Files that do **not** sit inside any `extracted_pages_YYYY_YY` directory are
still processed but land directly in the output root.

### Generated output structure

```text
data/transform/01_output/
└── iia_extracted_pages_YYYY/
    └── iia_{topic}_YYYY/
        └── r_iia_<topic>_<year>_<start>_<end>_<product>.xlsx
```

The subfolder name is derived from:

1. A directory sitting **between** `extracted_pages_*` and the workbook file
   (e.g. `extracted_pages_*/crops/file.xlsx` → `iia_crops_YYYY`).
2. Or, when the file is directly inside `extracted_pages_*`, the directory
   **directly above** it (e.g. `trade/extracted_pages_*/file.xlsx`
   → `iia_trade_YYYY`; `multiple product/extracted_pages_*/file.xlsx`
   → `iia_multiple_product_YYYY`).

#### Concrete example

Input:

```text
data/transform/00_input/
└── trade/
    └── extracted_pages_1938_39/
        ├── reviewed_466_475arrozimp_exp.xlsx
        └── crops/
            └── reviewed_239_239azucar_caña_brutaprod.xlsx
```

Generated output:

```text
data/transform/01_output/
└── iia_extracted_pages_1938/
    ├── iia_trade_1938/
    │   └── r_iia_trade_1938_466_475_rice.xlsx
    └── iia_crops_1938/
        └── r_iia_trade_1938_239_239_raw_cane_sugar.xlsx
```

#### Folder and file name rules

- Spaces in any folder or file name segment are replaced with `_`.
- Consecutive underscores are collapsed to a single `_`.
- Leading and trailing underscores are stripped.
- The yearbook name is taken from the folder that sits immediately above
  `extracted_pages_YYYY_YY` and normalised with the rules above.
- The product segment is stripped of trailing suffixes (`sup`, `prod`, `rend`,
  `imp`, `exp`, `num`) and then translated to English using
  `product_translations` from the config.

## Usage

Once your Excel files are in place under `data/transform/00_input/`, run the tool
from the project root:

```bash
# Use the conventional default paths (data/transform/00_input/ → data/transform/01_output/)
iia-excel-reorg --config workflow/config/example.units.yml
```

You can also specify paths explicitly:

```bash
# Single workbook
iia-excel-reorg path/to/source.xlsx path/to/output/ --config workflow/config/example.units.yml

# Entire directory tree
iia-excel-reorg path/to/input_dir path/to/output/ --config workflow/config/example.units.yml
```

The `input` argument defaults to `data/transform/00_input/` and the `output_dir`
argument defaults to `data/transform/01_output/` — both relative to the current
working directory.

### Independent pipeline: Footnote Harmonization Pipeline

Use the independent `iia-footnote-harmonizer` command to clean and remap
footnotes directly inside `data/transform/01_output/` while preserving the
exact folder/file layout.

```bash
# 1) Scan all files and generate a mapping template
iia-footnote-harmonizer generate-template data/transform/01_output data/footnote_mapping.xlsx

# 2) Fill "Cleaned Footnote" manually in the generated template

# 3) Apply mapping in place to every workbook
iia-footnote-harmonizer apply-mapping data/transform/01_output data/footnote_mapping.xlsx
```

The template contains two columns: `Original Footnote` (auto-populated with all
unique footnotes found across files/sheets/rows) and `Cleaned Footnote` (manual
mapping target). Multiple original footnotes may map to the same cleaned value.

### Independent pipeline: Pre-processing for Human Review

Use the `iia-prepare` command (or `run_preprocessing.py` for VS Code) to copy
raw Excel files into a human-ready workspace while preserving the exact folder
structure.

```bash
# Copy all Excel files from source to prepared workspace
iia-prepare

# Or specify custom paths
iia-prepare path/to/source path/to/prepared
```

**Default paths:**
- Input: `data/preprocess/00_input/`
- Output: `data/preprocess/01_output/`

The pipeline:
1. Discovers all `.xlsx` / `.xlsm` files recursively
2. Mirrors the exact folder/subfolder structure into the output
3. Shows progress bars matching the main pipeline aesthetics
4. Is designed to be extended with custom preparation edits

**VS Code runner:**

```bash
python run_preprocessing.py
```

## Transformation rules implemented

### Sheet names

Every processed worksheet is copied into the output workbook with the original sheet name converted to lowercase.

### Headers

The output header row is always:

```text
hemisphere | continent | country | unit | footnotes | <year/period columns...>
```

Year and period labels are preserved exactly as they appear in row 1 of the source sheet.

### Hierarchy extraction

The parser treats rows containing variants of `HÉMISPHÈRE` / `HEMISPHERE` as hemisphere labels.

The parser treats continent rows such as:

- `EUROPE`
- `AMÉRIQUE`
- `ASIE`
- `AFRIQUE`
- `OCÉANIE`

as continent labels.

These structural rows are not written as data rows. Instead, they populate the `hemisphere` and `continent` columns for subsequent country rows.

### Footnote extraction

Country labels such as:

```text
Belgique-Luxembourg (reexports) (special case)
```

become:

- `country = Belgique-Luxembourg`
- `footnotes = reexports; special case`

If notes contain references to units, those references stay in `footnotes`; they do not alter the `unit` assignment.

### Harmonized document names

Reviewed source names are converted with these rules:

- `reviewed_` becomes `r_`
- missing agency defaults to `iia`
- yearbook metadata comes from the folder path, for example `transform/00_input/trade/extracted_pages_1938_39` becomes `trade` and `1938`
- the product segment is stripped of trailing suffixes like `sup`, `prod`, `rend`, `imp`, `exp`, and `num`
- the remaining product is translated to English for the output filename

Examples:

```text
reviewed_239_239azucar_caña_brutaprod -> r_iia_trade_1938_239_239_raw_cane_sugar
reviewed_466_475arrozimp_exp -> r_iia_trade_1938_466_475_rice
```

### Unit assignment

The workflow now derives `unit` from:

- the sheet variable (`area`, `production`, `imports`, `exports`, etc.)
- the extracted source product name
- the document category supplied in config
- the configured rule mode (`standard` or `inputs`)

This replaces the provisional sheet-to-unit mapping with the rule-based logic you specified.

### Color preservation

The workflow preserves source fills for:

- row label cells copied into metadata columns
- numeric year/period value cells copied into the output sheet

## Development

Run tests locally:

```bash
pytest
```

Run the CLI module directly:

```bash
python -m iia_excel_reorg.cli path/to/source.xlsx output/ --config workflow/config/example.units.yml
```

## GitHub Actions

The repository includes a CI workflow in `.github/workflows/ci.yml` that runs `pytest` on every push and pull request.
