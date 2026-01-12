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

Implementation notes and the intended math model are documented inline at:

- `crates/formula-engine/src/functions/financial/bonds_odd.rs`

The thin “Excel built-in” wrappers (argument coercion, error mapping, registration in the
function registry) typically live under:

- `crates/formula-engine/src/functions/financial/`

## Tests and oracle cases

### Unit tests

Primary targeted tests live at:

- `crates/formula-engine/tests/functions/financial_odd_coupon.rs`

These tests currently focus on invariants that should *always* hold (e.g. independence from the
workbook date system). As additional edge cases are discovered (EOM schedules, long first coupons,
30/360 boundaries, solver convergence cliffs), add tests here to pin behavior.

### Excel oracle dataset (cross-check against real Excel)

To validate parity against real Excel results, use the Excel oracle harness:

- `tools/excel-oracle/README.md`
- case corpus: `tests/compatibility/excel-oracle/cases.json`
- pinned dataset: `tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json`

When adding odd-coupon coverage to the oracle corpus, prefer:

- fixed calendar dates (avoid `TODAY()`, `NOW()`, volatile functions)
- explicit `basis` / `frequency`
- cases that cover both Excel 1900 and 1904 date systems (the engine should match in both)

## High-risk compatibility areas

If you modify these functions, re-check these areas carefully:

1. **`basis` day-count mapping**: must match Excel (`0..=4` with specific 30/360 rules).
2. **Coupon schedule generation**: Excel’s end-of-month behavior is subtle; changing schedule
   stepping can change `E`, accrued interest, and discounting exponents.
3. **Error behavior**: Excel’s choice of `#NUM!` vs `#VALUE!` varies by argument and coercion path.
4. **Yield solvers**: Newton-Raphson failures must not silently return incorrect values; when Excel
   converges, we should converge as well (usually via a fallback/bracketing strategy).

