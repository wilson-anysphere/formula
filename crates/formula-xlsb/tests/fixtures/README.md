XLSB fixtures for `formula-xlsb` tests
=====================================

This directory contains `.xlsb` workbooks used by the regression tests in
`crates/formula-xlsb/tests/`.

In particular, `calaminex_compare.rs` automatically discovers all `*.xlsb` files
under this directory and compares:

- Calamine's `worksheet_formula` output (formula text), and
- `formula-xlsb`'s decoded `rgce` formula text.

Adding coverage
--------------

To expand rgce decoding coverage, add more `.xlsb` files under this directory
(subdirectories are fine). The compare harness will pick them up automatically.

Important: fixtures must be readable by Calamine.
------------------------------------------------

The in-repo minimal fixture generator (`tests/fixture_builder.rs`) does **not**
emit a valid `xl/styles.bin`, so generated workbooks omit styles entirely.
Calamine is tolerant of missing styles, but it will fail on an invalid/placeholder
`styles.bin`.
