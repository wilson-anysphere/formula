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

Note: Excel also exposes `TRUE()` and `FALSE()` as zero-arg worksheet functions, and their spellings
are localized (e.g. `WAHR()` / `FALSCH()`, `VRAI()` / `FAUX()`, `VERDADERO()` / `FALSO()`). These
need to be present in the TSVs just like any other function, even though localized boolean literals
(`WAHR`, `VRAI`, `VERDADERO`, …) are handled separately by the parser.

### Function translation sources (`sources/<locale>.json`)

The `*.tsv` files in this directory are **generated artifacts**.

Locale-specific function translations are sourced from deterministic JSON files under:

- `crates/formula-engine/src/locale/data/sources/*.json`

Missing entries are treated as identity mappings (canonical == localized).

**Important:** `sources/<locale>.json` should be generated from a **real Excel install** via
[`tools/excel-oracle/extract-function-translations.ps1`](../../../../../tools/excel-oracle/extract-function-translations.ps1)
whenever possible. Hand-maintained or web-scraped translation tables are frequently incomplete, and
any missing entry will silently fall back to English in the generated TSVs.

For readability and cleaner diffs, the committed `sources/*.json` files omit explicit identity
entries (e.g. `ABS -> ABS`). After generating a source JSON, normalize it in-place with:

```bash
node scripts/normalize-locale-function-sources.js
# or:
pnpm normalize:locale-function-sources
```

To verify that the committed sources are normalized (CI-style check), run:

```bash
node scripts/normalize-locale-function-sources.js --check
# or:
pnpm check:locale-function-sources
```

Note: after normalization, the JSON will typically contain **fewer entries** than
`shared/functionCatalog.json`, since identity mappings are omitted.

#### `es-ES` (Spanish) source requirements

Like `de-DE` and `fr-FR`, `es-ES` must be backed by an Excel-extracted mapping generated from the
**full function catalog** (see `shared/functionCatalog.json`).

Do **not** replace `sources/es-ES.json` with partial online translation tables: those commonly omit
large parts of Excel’s function surface area, causing many Spanish spellings to degrade to canonical
(English) names via identity-mapping fallback.

If you cannot run the Excel extractor (Windows + Excel desktop required), prefer to **not** touch
`es-ES` sources rather than committing a partial mapping.

#### Generating `sources/<locale>.json` from a real Excel install (Windows)

The most reliable way to obtain a complete translation mapping for a locale is to ask
**real Microsoft Excel** what it displays for each canonical function name.

From repo root on Windows (requires Excel desktop installed and configured for that locale):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-function-translations.ps1 `
  -LocaleId de-DE `
  -OutPath crates/formula-engine/src/locale/data/sources/de-DE.json
```

Example for Spanish (`es-ES`):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-function-translations.ps1 `
  -LocaleId es-ES `
  -OutPath crates/formula-engine/src/locale/data/sources/es-ES.json
```

Notes / caveats:

- The extracted spellings reflect the **active Excel UI language** (Office language packs + Excel
  display language settings). The script prints the detected Excel UI locale and warns if it does
  not match `-LocaleId`.
- For `de-DE`/`es-ES`/`fr-FR`, the extractor also runs a quick sanity check on a few sentinel
  translations (e.g. `SUM`/`IF`) and warns if Excel appears misconfigured.
- The extractor also warns if Excel maps multiple canonical functions to the same localized
  spelling; in that case `scripts/generate-locale-function-tsv.js` will fail due to ambiguity.

For debugging, you can also pass:

- `-Visible` to watch Excel work
- `-MaxFunctions N` to run a smaller subset (debugging only; do not commit partial sources)
- PowerShell's `-Verbose` switch for per-function `Formula` / `FormulaLocal` logging

After updating the source JSON, regenerate and verify the generated TSVs:

```bash
node scripts/normalize-locale-function-sources.js
node scripts/generate-locale-function-tsv.js
node scripts/generate-locale-function-tsv.js --check
```

#### Verification checklist (function TSVs)

When updating `sources/<locale>.json` (especially `es-ES`) and regenerating `<locale>.tsv`, verify:

1. If you re-extracted from Excel, confirm the extractor ran against the full catalog:
   - It should print `Wrote <N> translations ...` where `<N>` matches the number of functions in
     `shared/functionCatalog.json` (before normalization removes identity mappings).
   - It should not report skipped functions (skipped/missing entries silently fall back to English).
2. `node scripts/generate-locale-function-tsv.js --check` passes.
   - Note: `--check` only verifies that the committed TSVs match what would be generated from the
     committed JSON sources; it does **not** prove that the sources are complete (missing entries are
     silently treated as identity mappings).
3. **Sentinel translations are present** (spot-check directly in the TSV).
   For `es-ES`, these should *not* be identity mappings:
   - `SUM` → `SUMA`
   - `IF` → `SI`
   - `NPV` → `VNA`
   - `IRR` → `TIR`
   - `PV` → `VA`
   - `FV` → `VF`
   - `PMT` → `PAGO`
   - `RATE` → `TASA`
4. **Identity mappings are not suspiciously high.**
   Missing translations are emitted as `Canonical == Localized`, so a partial source can look
   “complete” while being mostly English. As a rough heuristic, `es-ES` should have an identity
   count in the same ballpark as `de-DE` / `fr-FR` (and should not have obvious identity mappings
   for major function groups like `SUM*`, `IF*`, `COUNT*`, `AVERAGE*`, etc).

   Quick count (Node required):

   ```bash
   node --input-type=module -e '
   import fs from "node:fs";
   for (const locale of ["de-DE","fr-FR","es-ES"]) {
     const tsv = fs.readFileSync(`crates/formula-engine/src/locale/data/${locale}.tsv`, "utf8");
     let ident = 0;
     for (const line of tsv.split(/\r?\n/)) {
       if (!line || line.startsWith("#")) continue;
       const [canon, loc] = line.split("\t");
       if (canon === loc) ident++;
     }
     console.log(`${locale}: ${ident} identity mappings`);
   }'
   ```

5. **No localized-name collisions.**
   - `tools/excel-oracle/extract-function-translations.ps1` warns if Excel maps multiple canonical
     functions to the same localized spelling.
   - `scripts/generate-locale-function-tsv.js` fails if it detects any collisions (ambiguity breaks
     localized → canonical lookup).

CI/tests also provide guard rails:

- `crates/formula-engine/tests/locale_function_tsv_completeness.rs` enforces that each locale TSV
  contains exactly one entry per catalog function and has no ambiguous localized collisions.
- `crates/formula-engine/tests/locale_de_de_function_sentinels.rs` asserts a small set of core
  German spellings (e.g. `SUM` → `SUMME`, `IF` → `WENN`) to catch regressions.
- `crates/formula-engine/tests/locale_es_es_function_sentinels.rs` is an explicit regression test
  for Spanish financial function spellings (including `NPV`/`IRR`), since missing entries otherwise
  silently fall back to English.

From repo root you can run (optional but recommended when editing locale data):

```bash
bash scripts/cargo_agent.sh test -p formula-engine --test locale_function_tsv_completeness
bash scripts/cargo_agent.sh test -p formula-engine --test locale_es_es_function_sentinels
```

### Error translations (`<locale>.errors.tsv`)

Locale-specific error literal spellings are tracked in TSV files in this directory
(e.g. `de-DE.errors.tsv`).

These TSVs are **committed artifacts** that are kept in sync with the engine’s canonical error set
([`ErrorKind`]) via the generator below.

**Runtime behavior:** [`FormulaLocale`] references an [`ErrorTranslations`] table backed by the
committed `*.errors.tsv` files (wired up via `include_str!()` in `src/locale/registry.rs`). The
engine parses these TSVs lazily at runtime to translate between canonical and localized error
literals.

Upstream localized spellings (used to (re)generate the committed TSVs) live under:

- `crates/formula-engine/src/locale/data/upstream/errors/*.tsv`

To extract/verify localized spellings against a real Excel install, see:

- `tools/excel-oracle/extract-error-literals.ps1`

Note: upstream error TSVs can contain multiple localized spellings per canonical error literal
(aliases). The extractor preserves any existing aliases in the output file; for `es-ES`, it also
records both inverted- and non-inverted-punctuation variants (e.g. `#¡VALOR!` and `#VALOR!`).

For debugging, you can also pass:

- `-Visible` to watch Excel work
- `-MaxErrors N` to run a smaller subset
- PowerShell’s `-Verbose` switch for per-error `FormulaLocal` / `.Text` logging

Note: `*.errors.tsv` exports are expected to match the runtime locale tables. The test
`crates/formula-engine/tests/locale_error_tsv_sync.rs` enforces that bidirectional mapping. If you
update upstream error spellings, regenerate the committed TSVs and ensure tests continue to pass.

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
- Data lines begin with the canonical error literal (e.g. `#VALUE!`) and map to the localized error
  literal (also starts with `#`).

The generator outputs entries sorted by canonical error literal for deterministic diffs.

Some locales include multiple localized spellings for the same canonical error literal to support
Excel-compatible alias spellings. In those cases:

- The **first** entry for a canonical error is treated as the preferred display spelling for
  localization (canonical → localized).
- All entries are accepted for canonicalization (localized → canonical).

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

The newer external-data errors (`#CONNECT!`, `#FIELD!`, `#BLOCKED!`, `#UNKNOWN!`) are included in
the error TSVs so both canonicalization (localized → canonical) and localization (canonical →
localized) are stable and Excel-compatible.

## Structured references

Excel table structured references have a small set of reserved **item keywords** such as:

- `[#Headers]`, `[#Data]`, `[#Totals]`, `[#All]`
- `[@]` / `[#This Row]`

Unlike function names / separators / error literals, these item keywords appear to be **canonical
(English) in Excel's formula language across our supported locales** (`de-DE`, `fr-FR`, `es-ES`).

To verify this behavior against a real Excel install (Windows + Excel desktop required), you can run:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-structured-reference-keywords.ps1 `
  -LocaleId de-DE `
  -OutPath out/
```

The script writes a small JSON report containing `Formula` vs `FormulaLocal` for a few sentinel
structured references (including `#All` and `#This Row`), making it easy to see whether Excel is
localizing these tokens for your current UI language configuration.

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
      - Add a `static <LOCALE>_ERRORS: ErrorTranslations = ...include_str!("data/<locale>.errors.tsv")`
        in `crates/formula-engine/src/locale/registry.rs`.
      - Add a `pub static <LOCALE>: FormulaLocale = ...` entry with separators + boolean literals and
        set `errors: &<LOCALE>_ERRORS` and `functions: &<LOCALE>_FUNCTIONS`.
      - Add the locale to `get_locale()` in `registry.rs`.
      - Update `crates/formula-engine/src/locale/mod.rs` (`normalize_locale_id`) so the engine can
        actually resolve locale tags to your new locale id (especially if you add a second locale for
        an existing language, e.g. `fr-CA` vs `fr-FR`).
     - Re-export the new constant from `crates/formula-engine/src/locale/mod.rs` if it should be
       accessible as `locale::<LOCALE>`.
4. **Add tests:** extend `crates/formula-engine/tests/locale_parsing.rs` with basic round-trip tests
   for separators, a couple of translated functions, and at least one localized error literal.
5. **Run generators in `--check` mode** to ensure TSVs stay in sync with the engine catalog.
