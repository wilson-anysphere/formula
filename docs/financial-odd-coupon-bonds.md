# Odd-coupon bond functions (ODDF\*/ODDL\*) — Excel parity notes

Excel’s odd-coupon bond functions have a large compatibility surface:

- day-count basis conventions (`basis`)
- end-of-month (EOM) coupon schedules
- settlement / issue / coupon date validation rules
- yield solver robustness

This document is meant to be a **maintainer guide**: if you change or extend these functions, it
should be clear what invariants must be preserved to avoid Excel regressions.

## Functions

Odd *first* coupon period:

- `ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, [basis])`
- `ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, [basis])`

Odd *last* coupon period:

- `ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, [basis])`
- `ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency, [basis])`

## Canonical references (Microsoft)

The most stable public docs we link to are the VBA `WorksheetFunction` pages:

- <https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.oddfprice>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.oddfyield>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.oddlprice>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.oddlyield>

## Where the implementation lives

The core implementation lives at:

- `crates/formula-engine/src/functions/financial/odd_coupon.rs`

Implementation notes and the intended math model are documented inline at:

- `crates/formula-engine/src/functions/financial/bonds_odd.rs`

The thin “Excel built-in” wrappers (argument coercion, error mapping, registration in the
function registry) typically live under:

- `crates/formula-engine/src/functions/financial/`

## Tests and oracle cases

### Unit tests

Primary targeted tests (and some boundary validations) live at:

- `crates/formula-engine/tests/functions/financial_odd_coupon.rs`

Date-ordering equality edge cases and other boundary validations are locked via:

- `crates/formula-engine/tests/odd_coupon_date_boundaries.rs`
- `crates/formula-engine/tests/functions/financial_oddcoupons.rs`

These tests currently focus on invariants that should *always* hold (e.g. independence from the
workbook date system). As additional edge cases are discovered (EOM schedules, long first coupons,
30/360 boundaries, solver convergence cliffs), add tests here to pin behavior.

### Excel oracle dataset (cross-check against real Excel)

To validate parity against real Excel results, use the Excel oracle harness:

- `tools/excel-oracle/README.md`
- case corpus: `tests/compatibility/excel-oracle/cases.json`
- pinned dataset: `tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json`

Note: the pinned dataset in CI is currently a **synthetic baseline** generated from the engine
(see the dataset `source.note`). Treat it as a regression test for current engine behavior.
To validate against *real* Excel, regenerate the dataset with
`tools/excel-oracle/run-excel-oracle.ps1` on Windows + Excel (Task 393).

When adding odd-coupon coverage to the oracle corpus, prefer:

- fixed calendar dates (avoid `TODAY()`, `NOW()`, volatile functions)
- explicit `basis` / `frequency`
- cases that cover both Excel 1900 and 1904 date systems (the engine should match in both)

The generator includes a small set of boundary-date equality cases (e.g. `issue == settlement`,
`settlement == first_coupon`, `first_coupon == maturity`, `issue == first_coupon`,
`settlement == last_interest`, `last_interest == maturity`).

It also includes a small set of **negative yield / negative coupon rate** validation cases
(tagged `odd_coupon_validation`) to confirm Excel’s input-domain behavior for:

- negative yields (including `yld < -1` but still `yld > -frequency`)
- the boundary `yld == -frequency` (`#DIV/0!`) vs `yld < -frequency` (`#NUM!`)
- negative coupon rates (whether Excel returns `#NUM!`)

For quick Windows + Excel runs, use the derived subset corpus:

- `tools/excel-oracle/odd_coupon_validation_cases.json`

Current engine behavior:

- **ODDF\*** enforces strict date ordering:
  - `issue < settlement < first_coupon <= maturity`
  - Equality boundaries like `issue == settlement`, `settlement == first_coupon`, and
    `issue == first_coupon` are rejected with `#NUM!` (see the oracle boundary cases + unit tests).
  - `first_coupon == maturity` is allowed (single odd stub period paid at maturity).
- **ODDL\*** requires `settlement < maturity` and `last_interest < maturity`, but allows settlement
  dates **on or before** `last_interest` (as well as inside the odd-last stub).
  - `settlement == last_interest` is allowed (it implies zero accrued interest).
  - `settlement == maturity` and `last_interest == maturity` are rejected with `#NUM!`.

See `crates/formula-engine/tests/odd_coupon_date_boundaries.rs` and
`crates/formula-engine/tests/functions/financial_oddcoupons.rs`.

These cases are tagged as `boundary` + `odd_coupon` and are included in the Excel-oracle CI smoke
gate so regressions in date validation are caught early. Since the pinned dataset in CI is
currently a synthetic baseline, treat these as regression coverage until we pin a real Excel
oracle dataset; at that point they serve as parity checks for Excel’s boundary behavior.

## High-risk compatibility areas

If you modify these functions, re-check these areas carefully:

1. **`basis` day-count mapping**: must match Excel (`0..=4` with specific 30/360 rules).
    - Note: for `basis=4` (European 30/360), day counts use `DAYS360(..., TRUE)`, but the modeled
      coupon-period length `E` is fixed: `E = 360/frequency` (see `coupon_schedule::coupon_period_e`).
      This intentionally diverges from `DAYS360(PCD, NCD, TRUE)` for some end-of-month schedules
      involving February (e.g. Feb 28 → Aug 31 yields 182 under European `DAYS360`, not 180).
2. **Coupon schedule generation (anchor date + EOM stepping)**: Excel’s end-of-month behavior is
   subtle; changing schedule stepping can change `E`, accrued interest, and discounting exponents.
    - Current implementation detail (ODDF\*): the regular coupon schedule is generated by stepping
      backward from `maturity` in whole coupon periods (`12 / frequency` months).
   - The schedule is treated as EOM iff `maturity` is itself end-of-month, and month stepping uses
     `date_time::eomonth` (Excel’s `EOMONTH`) in that case (otherwise `date_time::edate`).
   - Current implementation detail (ODDL\*): to compute the regular period length `E`, we step one
     coupon period backward from `last_interest`; EOM detection for that step is based on
     `last_interest` being end-of-month.
   - This logic lives in `crates/formula-engine/src/functions/financial/odd_coupon.rs` (see
     `coupon_schedule_from_maturity` and `coupon_date_with_eom`).
3. **Error behavior**: Excel’s choice of `#NUM!` vs `#VALUE!` varies by argument and coercion path.
4. **Yield domain + solvers**:
   - Domain: the per-period discount base must stay positive (`1 + yld/frequency > 0`, i.e.
      `yld > -frequency`). The boundary `yld == -frequency` produces `#DIV/0!` in Excel; below that is
      `#NUM!`.
   - Solver behavior: Newton-Raphson failures must not silently return incorrect values; when Excel
      converges, we should converge as well (usually via a fallback/bracketing strategy).

## Long odd periods (DFC/E > 1, DLM/E > 1)

Excel supports both **short** and **long** odd coupon periods.

The “long” cases are important because they stress:

- coupon scaling (`coupon_1 = C * DFC/E`, `coupon_last = C * DLM/E`) when the odd period spans **multiple**
  regular coupon intervals
- discount exponent logic when `DSC/E > 1` (ODDF\*) or `DSM/E > 1` (ODDL\*)
- `basis=1` actual/actual `E` computation across leap years

Our implementation follows the standard Excel-style model:

- Regular coupon payment per period: `C = redemption * rate / frequency`
- ODDF\*:
  - `A = days(issue, settlement)`
  - `DFC = days(issue, first_coupon)`
  - `DSC = days(settlement, first_coupon)`
  - `E = regular coupon period length (days)`
  - First coupon amount: `C1 = C * (DFC/E)` (so `DFC/E > 1` produces a long first coupon)
  - Accrued interest: `AI = C * (A/E)`
  - Discount exponent for cashflow `i` on the coupon schedule: `t_i = (DSC/E) + (i-1)`
- ODDL\*:
  - If `settlement >= last_interest` (settlement inside the odd-last stub):
    - `A = days(last_interest, settlement)`
    - `DLM = days(last_interest, maturity)`
    - `DSM = days(settlement, maturity)`
    - `E = regular coupon period length (days)`
    - Final coupon amount: `Clast = C * (DLM/E)` (so `DLM/E > 1` produces a long last coupon)
    - Accrued interest: `AI = C * (A/E)`
    - Discount exponent: `t = DSM/E`
  - If `settlement < last_interest`, pricing must include the remaining regular coupon payments through
    `last_interest` (inclusive) plus the final odd-stub payment at maturity. Accrued interest is
    computed from the regular coupon period containing settlement (see
    `crates/formula-engine/src/functions/financial/odd_coupon.rs::oddl_equation`).

### Where the long-stub cases live

- Engine unit tests:
  - `crates/formula-engine/tests/functions/financial_odd_coupon.rs` (search for `round_trip_long_stub`)
- Excel oracle subsets (for quick Windows + Excel runs):
  - `tools/excel-oracle/odd_coupon_long_stub_cases.json` (uses canonical `caseId`s from `cases.json`)
- Canonical oracle corpus:
  - `tests/compatibility/excel-oracle/cases.json` (tagged `odd_coupon` + `long_stub`)

## Date validation rules (Excel docs + engine behavior)

The official Microsoft VBA `WorksheetFunction` docs specify strict date ordering and note that
date-like inputs are truncated to integers before validation:

- ODDF\* (docs): `maturity > first_coupon > settlement > issue`
  - <https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.oddfprice>
  - <https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.oddfyield>
- ODDL\* (docs): `maturity > settlement > last_interest`
  - <https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.oddlprice>
  - <https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.oddlyield>

In practice, parity testing against the curated excel-oracle corpus shows Excel rejects the ODDF\*
equality boundaries `issue == settlement` and `settlement == first_coupon` with `#NUM!`, matching
the strict ordering documented in Microsoft’s `WorksheetFunction` docs.

The current engine implementation enforces:

- ODDF\*: `issue < settlement < first_coupon <= maturity` (allows `first_coupon == maturity`)
- ODDL\*: `settlement < maturity` and `last_interest < maturity` (settlement may be before, on, or
  after `last_interest`; see `odd_coupon.rs::oddl_equation`).

These boundaries are covered by:

- Unit tests:
  - `crates/formula-engine/tests/odd_coupon_date_boundaries.rs`
- Excel oracle corpus cases tagged `odd_coupon` + `boundary`:
  - `tests/compatibility/excel-oracle/cases.json` (search for e.g. “ODDFPRICE boundary: settlement == first_coupon”)

Note: the engine currently allows `first_coupon == maturity` for ODDF\* (a single odd coupon paid
at maturity). This behavior is covered by `crates/formula-engine/tests/odd_coupon_date_boundaries.rs`
and by the excel-oracle boundary cases "ODDFPRICE boundary: first_coupon == maturity" and
"ODDFYIELD boundary: first_coupon == maturity".

### Notes on `basis = 1` (Actual/Actual)

The Microsoft docs list `basis=1` as **Actual/Actual**. For odd-coupon functions, `basis=1` implies
the regular coupon-period length `E` is computed as the **actual number of days between coupon
dates** (consistent with `COUP*`/`PRICE` behavior).

### Notes on `basis = 4` (30E/360 European)

`basis=4` is **European 30/360** (`DAYS360(..., method=TRUE)`).

For `basis=4`, day counts use `DAYS360(..., TRUE)`, but the modeled coupon-period length `E` is:

```
E = 360/frequency
```

Day counts like `A`/`DFC`/`DSC` still use `DAYS360(..., TRUE)`, so for some end-of-month schedules
involving February `DAYS360(PCD, NCD, TRUE)` can differ from the modeled `E` (e.g. Feb 28 → Aug 31
yields 182 under European `DAYS360`, not 180).

For `basis=4`, the remaining days in the coupon period (`DSC`) are computed as `DSC = E - A`, so
`DSC` is not always equal to `DAYS360(settlement, NCD, TRUE)` (see
`crates/formula-engine/tests/functions/financial_coupons.rs`).

### Excel oracle run (odd-coupon cases only)

To confirm these rules against a specific Excel version/build, run the oracle harness on Windows +
Excel and filter to the `odd_coupon` tag:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tests/compatibility/excel-oracle/cases.json `
  -OutPath /path/to/excel-odd-coupon-results.json `
  -IncludeTags odd_coupon
```

To focus specifically on the negative yield / negative coupon validation scenarios, filter to
the `odd_coupon_validation` tag (or use the subset corpus `tools/excel-oracle/odd_coupon_validation_cases.json`):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tests/compatibility/excel-oracle/cases.json `
  -OutPath /path/to/excel-odd-coupon-validation-results.json `
  -IncludeTags odd_coupon_validation
```
