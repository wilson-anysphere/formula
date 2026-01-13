# Locale translation data (`*.tsv`)

The formula engine **persists and evaluates formulas in canonical Excel (en-US) form**:

- Canonical (English) function identifiers (e.g. `SUM`, `CUBEVALUE`)
- `,` as the argument separator / union operator
- `.` as the decimal separator
- Canonical error literals (e.g. `#VALUE!`, `#GETTING_DATA`)

UI/editor workflows can translate formulas to/from locale display forms via:

- `locale::canonicalize_formula*` (localized → canonical)
- `locale::localize_formula*` (canonical → localized)
- `Engine::set_cell_formula_localized*`

This directory contains the translation tables used by the locale translation pipeline
(`crates/formula-engine/src/locale`).

## Files and completeness requirements

### Function translations (`<locale>.tsv`)

Each supported locale has a function translation TSV (e.g. `de-DE.tsv`, `fr-FR.tsv`, `es-ES.tsv`).

**Completeness goal:** every function TSV must contain **exactly one entry per function** in the
engine’s function catalog (see `shared/functionCatalog.json`).

This is important because:

- missing entries silently fall back to the canonical (English) name in both directions, which
  breaks round-tripping and localized editing;
- keeping the TSVs complete lets us catch regressions any time a function is added/renamed.

### Error translations (`<locale>.errors.tsv`)

Error literal translations are maintained in the locale registry (`src/locale/registry.rs`), but we
also commit a TSV export per locale (e.g. `de-DE.errors.tsv`) for auditing and keeping coverage in
sync with the engine’s error set (`ErrorKind`).

Upstream localized spellings (used to (re)generate the committed TSVs) live under:

- `crates/formula-engine/src/locale/data/upstream/errors/*.tsv`

## TSV format

Each TSV file is a simple tab-separated mapping:

```
Canonical<TAB>Localized
```

### Function-translation TSVs

Function-translation TSVs live directly in this directory (e.g. `de-DE.tsv`) and use the
following conventions:

- Lines starting with `#` and empty lines are ignored.
- The canonical column must use the engine’s canonical spelling.
- The localized column should use **the exact spelling Excel displays for that locale**.

### Error-literal TSVs (`<locale>.errors.tsv`)

Error literals themselves start with `#`, so error TSVs use a stricter comment convention:

- Lines where the first non-whitespace characters are `#` followed by whitespace are comments
  (e.g. `# Canonical<TAB>Localized`).
- Empty lines are ignored.
- Data lines begin with the canonical error literal (e.g. `#VALUE!`).

## Case-folding, Unicode, and why values are stored uppercase

Excel treats function identifiers case-insensitively. Our locale translation layer matches that by
normalizing identifiers before lookup:

- Identifiers are **case-folded using Unicode-aware uppercasing** (Rust `char::to_uppercase`).
- TSV entries are stored in their **already-case-folded (uppercase) form** so the runtime can do a
  direct hash lookup against the `include_str!` data without allocating or case-folding every table
  entry at startup.

Practical takeaway: keep the TSV `Localized` values uppercase (including non-ASCII characters), and
run the generators below to enforce normalization.

## Generators and `--check`

TSVs are maintained by small generator tools so we can enforce:

- completeness against the engine catalog (`shared/functionCatalog.json`);
- normalization (case-folded uppercase);
- deterministic ordering and stable diffs.

Run these from the repo root:

```bash
# Regenerate function TSVs (writes files in-place)
cargo run -p formula-engine --bin generate_locale_function_tsv

# Verify function TSVs are up to date (CI mode)
cargo run -p formula-engine --bin generate_locale_function_tsv -- --check

# Regenerate error TSVs from committed upstream mapping sources
node scripts/generate-locale-error-tsvs.mjs

# Verify error TSVs are up to date (CI mode)
node scripts/generate-locale-error-tsvs.mjs --check
```

The error TSV generator derives the canonical error literal list from
`formula_engine::value::ErrorKind::as_code` (scraped from `crates/formula-engine/src/value/mod.rs`)
so new error kinds automatically flow through the generator.

`--check` exits non-zero if any files would change.

## Adding a new locale

1. **Create the TSV(s):**
   - Add `crates/formula-engine/src/locale/data/<locale>.tsv` for function names.
   - Add `crates/formula-engine/src/locale/data/upstream/errors/<locale>.tsv` and run the error TSV
     generator to produce `crates/formula-engine/src/locale/data/<locale>.errors.tsv`.
2. **Register the locale in code:**
   - Add a `static <LOCALE>_FUNCTIONS: FunctionTranslations = ...include_str!("data/<locale>.tsv")`
     in `crates/formula-engine/src/locale/registry.rs`.
   - Add a `pub static <LOCALE>: FormulaLocale = ...` entry with separators + boolean literals +
     error literal mappings.
   - Add the locale to `get_locale()` in `registry.rs`.
   - Re-export the new constant from `crates/formula-engine/src/locale/mod.rs` if it should be
     accessible as `locale::<LOCALE>`.
3. **Add tests:** extend `crates/formula-engine/tests/locale_parsing.rs` with basic round-trip tests
   for separators, a couple of translated functions, and at least one localized error literal.
4. **Run generators in `--check` mode** to ensure TSVs stay in sync with the engine catalog.

