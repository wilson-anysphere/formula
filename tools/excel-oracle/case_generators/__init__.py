"""
Modular Excel-oracle case generators.

The Excel-oracle corpus is intentionally small (~2k cases max) but must cover
every non-volatile function in `shared/functionCatalog.json`.

To reduce merge conflicts as the function catalog grows, the corpus generator is
split into per-category modules under `tools/excel-oracle/case_generators/`.

Each module exposes:

    generate(cases, *, add_case, CellInput, ...helpers) -> None

`tools/excel-oracle/generate_cases.py` is responsible for:
  - defining shared helpers (CellInput, add_case, excel serial conversion, etc.)
  - invoking these modules in a deterministic order
  - validating the final corpus against `shared/functionCatalog.json`
"""

