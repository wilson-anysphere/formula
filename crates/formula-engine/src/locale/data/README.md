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

### Function translation sources (`sources/<locale>.json`)

The `*.tsv` files in this directory are **generated artifacts**.

Locale-specific function translations are sourced from deterministic JSON files under:

- `crates/formula-engine/src/locale/data/sources/*.json`

Missing entries are treated as identity mappings (canonical == localized).

#### Generating `sources/<locale>.json` from a real Excel install (Windows)

The most reliable way to obtain a complete translation mapping for a locale is to ask
**real Microsoft Excel** what it displays for each canonical function name.

From repo root on Windows (requires Excel desktop installed and configured for that locale):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-function-translations.ps1 `
  -LocaleId de-DE `
  -OutPath crates/formula-engine/src/locale/data/sources/de-DE.json
```

Notes / caveats:

- The extracted spellings reflect the **active Excel UI language** (Office language packs + Excel
  display language settings). The script prints the detected Excel UI locale and warns if it does
  not match `-LocaleId`.

For debugging, you can also pass:

- `-Visible` to watch Excel work
- `-MaxFunctions N` to run a smaller subset

After updating the source JSON, regenerate and verify the generated TSVs:

```bash
node scripts/generate-locale-function-tsv.js
node scripts/generate-locale-function-tsv.js --check
```

### Error translations (`<locale>.errors.tsv`)

Error literal translations are maintained in the locale registry (`src/locale/registry.rs`), but we
also commit a TSV export per locale (e.g. `de-DE.errors.tsv`) for auditing and keeping coverage in
sync with the engine’s error set (`ErrorKind`).

Upstream localized spellings (used to (re)generate the committed TSVs) live under:

- `crates/formula-engine/src/locale/data/upstream/errors/*.tsv`

To extract/verify localized spellings against a real Excel install, see:

- `tools/excel-oracle/extract-error-literals.ps1`

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

The generator outputs entries sorted by canonical error literal for deterministic diffs.

## Case-folding, Unicode, and why values are stored uppercase

Excel treats function identifiers case-insensitively. Our locale translation layer matches that by
normalizing identifiers before lookup and when loading TSVs:

- The engine uses `crate::value::casefold` (Unicode-aware uppercasing via `char::to_uppercase`) so
  case-insensitive matching behaves like Excel (e.g. `ß` → `SS`).
- When building the locale translation maps, both the canonical and localized columns are
  case-folded into hash keys so lookups are case-insensitive and duplicates are detected reliably.

Practical takeaway: keep the TSV `Localized` values uppercase (including non-ASCII characters), and
run the generators below to enforce normalization. This is primarily for **deterministic diffs** and
to mirror Excel’s UI conventions; the runtime still accepts mixed-case input.

## Generators and `--check`

TSVs are maintained by small generator tools so we can enforce:

- completeness against the engine catalog (`shared/functionCatalog.json`);
- normalization (case-folded uppercase);
- deterministic ordering and stable diffs.

Run these from the repo root:

```bash
# If the engine function catalog changed:
node scripts/generate-function-catalog.js

# Regenerate function TSVs (writes files in-place)
node scripts/generate-locale-function-tsv.js

# Verify function TSVs are up to date (CI mode)
node scripts/generate-locale-function-tsv.js --check

# Regenerate error TSVs from committed upstream mapping sources
node scripts/generate-locale-error-tsvs.mjs

# Verify error TSVs are up to date (CI mode)
node scripts/generate-locale-error-tsvs.mjs --check
```

The function TSV generator uses `SOURCE_DATE_EPOCH` (seconds since Unix epoch) to produce a stable
generation date in the TSV header. If it is not set, it defaults to `0` for reproducible output.

The error TSV generator derives the canonical error literal list from
`formula_engine::value::ErrorKind::as_code` (scraped from `crates/formula-engine/src/value/mod.rs`)
so new error kinds automatically flow through the generator.

`--check` exits non-zero if any files would change.

## External-data worksheet functions / errors

These locales include explicit coverage for Excel external-data worksheet functions:

- `RTD`
- `CUBEVALUE`
- `CUBEMEMBER`
- `CUBEMEMBERPROPERTY`
- `CUBERANKEDMEMBER`
- `CUBESET`
- `CUBESETCOUNT`
- `CUBEKPIMEMBER`

And for the external-data loading error literal:

- `#GETTING_DATA`

The newer external-data errors (`#CONNECT!`, `#FIELD!`, `#BLOCKED!`, `#UNKNOWN!`) currently
round-trip unchanged (canonical) for all supported locales.

## Structured references

Excel table structured references have a small set of reserved **item keywords** such as:

- `[#Headers]`, `[#Data]`, `[#Totals]`, `[#All]`
- `[@]` / `[#This Row]`

Unlike function names / separators / error literals, these item keywords appear to be **canonical
(English) in Excel's formula language across our supported locales** (`de-DE`, `fr-FR`, `es-ES`).

Accordingly:

- `locale::canonicalize_formula*` and `locale::localize_formula*` intentionally **do not translate**
  structured-reference item keywords.
- Separators **inside structured references** are also canonical (commas inside `Table1[[...],[...]]`
  are not locale-dependent), so translation avoids rewriting anything inside `[...]` bracket groups.

## Structured references (`Table1[...]`) and bracketed segments

Excel uses `[...]` for:

- **Structured references** (tables), e.g. `Table1[Col]`, `Table1[[#Headers],[Col]]`, `[@Col]`
- **External workbook prefixes**, e.g. `[Book.xlsx]Sheet1!A1`

Our locale translation pipeline in `src/locale/translate.rs` treats **everything inside
`[...]` as opaque** and does **not** translate it.

This matches Excel behavior for the supported locales (`de-DE`, `fr-FR`, `es-ES`):

- Structured-reference **item keywords are canonical (English)** and are not localized:
  `#All`, `#Data`, `#Headers`, `#Totals`, `#This Row`
- Structured-reference **syntax separators** inside the brackets are also canonical and are not
  locale-dependent (e.g. the comma in `Table1[[#Headers],[Col]]` remains `,` even when the locale
  uses `;` for function arguments).

Treating bracket content as opaque is also important for correctness because it may contain:

- workbook names (`[Book.xlsx]`) that must never be rewritten
- table/column identifiers that may collide with localized boolean keywords (e.g. `WAHR`)
- Excel’s `]]` escape sequence for a literal `]` inside table/column names

## Adding a new locale

1. **Create the sources:**
   - Add `crates/formula-engine/src/locale/data/sources/<locale>.json` for function name translations.
   - Add `crates/formula-engine/src/locale/data/upstream/errors/<locale>.tsv` for error literals.
2. **Run generators:**
   - Run `node scripts/generate-locale-function-tsv.js` to produce/update
     `crates/formula-engine/src/locale/data/<locale>.tsv`.
   - Run `node scripts/generate-locale-error-tsvs.mjs` to produce/update
     `crates/formula-engine/src/locale/data/<locale>.errors.tsv`.
3. **Register the locale in code:**
   - Add a `static <LOCALE>_FUNCTIONS: FunctionTranslations = ...include_str!("data/<locale>.tsv")`
     in `crates/formula-engine/src/locale/registry.rs`.
   - Add a `pub static <LOCALE>: FormulaLocale = ...` entry with separators + boolean literals +
     error literal mappings.
   - Add the locale to `get_locale()` in `registry.rs`.
   - Re-export the new constant from `crates/formula-engine/src/locale/mod.rs` if it should be
     accessible as `locale::<LOCALE>`.
4. **Add tests:** extend `crates/formula-engine/tests/locale_parsing.rs` with basic round-trip tests
   for separators, a couple of translated functions, and at least one localized error literal.
5. **Run generators in `--check` mode** to ensure TSVs stay in sync with the engine catalog.
