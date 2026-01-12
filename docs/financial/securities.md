# Discount securities, interest-at-maturity, and T-bill functions (Excel-compatible spec)

This document is an **internal implementation spec** for the Excel-compatible security functions implemented in:

- Core math + validation: `crates/formula-engine/src/functions/financial/securities.rs`
- Engine-facing builtins (argument coercion, defaults, error mapping): `crates/formula-engine/src/functions/financial/builtins_securities.rs`

Functions covered:

- Discount securities:
  - `DISC`, `PRICEDISC`, `YIELDDISC`, `INTRATE`, `RECEIVED`
- Interest-at-maturity securities:
  - `PRICEMAT`, `YIELDMAT`
- Treasury bills:
  - `TBILLPRICE`, `TBILLYIELD`, `TBILLEQ`

The intent is to clearly define:

- **Excel-compatible input coercion** (dates, basis) and **defaulting**
- **Validation rules** and error types (`#NUM!` vs `#DIV/0!` vs `#VALUE!`)
- **Exact formulas**, including `YEARFRAC` usage and `TBILLEQ`’s 182-day branching

---

## Terminology and shared derived values

### Dates

All date arguments (`settlement`, `maturity`, `issue`) are treated as Excel date serials in the active `ExcelDateSystem`.

For the discount-security and interest-at-maturity functions we use:

- `yf = YEARFRAC(settlement, maturity, basis)`

For `PRICEMAT` / `YIELDMAT`:

- `im = YEARFRAC(issue, maturity, basis)`
- `is = YEARFRAC(issue, settlement, basis)`
- `sm = YEARFRAC(settlement, maturity, basis)`

For the T-bill functions we use:

- `DSM = maturity - settlement` (an integer serial-day difference, “days from settlement to maturity”)

### Day-count basis (`basis`) and `YEARFRAC`

These functions use `YEARFRAC` directly, with Excel-style bases `0..=4` implemented by:

- `crates/formula-engine/src/functions/date_time/mod.rs::yearfrac`

| basis | Name | `YEARFRAC(start,end,basis)` behavior |
|------:|------|--------------------------------------|
| 0 | US (NASD) 30/360 | `DAYS360(start,end,FALSE) / 360` |
| 1 | Actual/Actual | Actual/Actual “year + remainder” algorithm (Excel-compatible) |
| 2 | Actual/360 | `(end - start) / 360` |
| 3 | Actual/365 | `(end - start) / 365` |
| 4 | European 30/360 | `DAYS360(start,end,TRUE) / 360` |

> `TBILLPRICE`, `TBILLYIELD`, and `TBILLEQ` do not accept a `basis` argument; they always use `DSM` (actual serial-day difference) with 360/365 constants as defined below.

---

## Argument coercion and defaults (engine/builtins layer)

This section documents the coercions performed by the engine-facing builtins in
`crates/formula-engine/src/functions/financial/builtins_securities.rs`.

### Date arguments (`settlement`, `maturity`, `issue`)

Dates accept either text or numbers:

- If the argument is **text**, it is parsed using `DATEVALUE` semantics:
  - `date_time::datevalue(text, locale, now_utc, system)`
  - Invalid text dates return `#VALUE!`.
- Otherwise, the argument is coerced to a number. If it is **non-finite** (`NaN`, `±Inf`) → `#NUM!`.
- Numeric date serials are **floored** (time-of-day is ignored):
  - Example: `43831.9` → `43831`
- Floored serials must fit in `i32` (`[-2^31, 2^31-1]`); otherwise → `#NUM!`.

### Optional `basis`

- Missing `basis` defaults to `0`.
- An explicit blank argument (`,,` in a formula) or a blank cell also defaults to `0`.
- Otherwise the argument is coerced to a finite number and **truncated** toward zero to an integer,
  then validated.
- If the truncated value is not in `0..=4` → `#NUM!`.
- If the argument cannot be coerced to a number (e.g. `"nope"`) → `#VALUE!`.

### Numeric (non-date) arguments

Required numeric arguments (`pr`, `redemption`, `discount`, `rate`, …) are coerced to numbers using
standard engine rules. If coercion yields a **non-finite** number → `#NUM!`.

---

## Shared validation rules (core functions)

Unless otherwise noted, all validations described here are implemented in
`crates/formula-engine/src/functions/financial/securities.rs`.

### Date ordering

- For `DISC`, `PRICEDISC`, `YIELDDISC`, `INTRATE`, `RECEIVED`:
  - Require `settlement < maturity` (strict).
  - Otherwise → `#NUM!`.
- For `PRICEMAT`, `YIELDMAT`:
  - Require `issue <= settlement < maturity`.
  - Additionally require `issue < maturity` (follows from the above, but validated explicitly).
  - Otherwise → `#NUM!`.
- For `TBILLPRICE`, `TBILLYIELD`, `TBILLEQ`:
  - Require `settlement < maturity` (strict).
  - Define `DSM = maturity - settlement` and require `1 <= DSM <= 365`.
  - Otherwise → `#NUM!`.

### Basis validation

For functions that take `basis`, the allowed set is:

- `basis ∈ {0,1,2,3,4}`
- Missing/blank basis defaults to `0` (see builtins rules above)
- Anything else → `#NUM!`

### Numeric domain constraints and finiteness

All numeric inputs must be finite (`is_finite`). Non-finite inputs → `#NUM!`.

If the final computed result (or a required intermediate like a denominator) is non-finite (`NaN`, `±Inf`), the function returns `#NUM!`.

Additional Excel-style domain constraints (implemented as `#NUM!` unless otherwise stated):

- `pr`, `redemption`, `investment` must be `> 0` where used as prices/amounts.
- `discount`, `rate`, `yld` must be `> 0` where required by the function.
- Computed prices that must be positive are rejected:
  - `PRICEDISC`: requires `1 - discount * YEARFRAC(...) > 0`.
  - `TBILLPRICE`: requires computed price `> 0`.
  - `TBILLEQ`: requires `1 - discount * DSM/360 > 0` (i.e., a positive implied price).

### When `#DIV/0!` occurs

These functions return `#DIV/0!` when a required divisor is exactly zero:

- `DISC`, `YIELDDISC`, `INTRATE`: if `YEARFRAC(settlement,maturity,basis) == 0`
- `YIELDMAT`: if `sm == YEARFRAC(settlement,maturity,basis) == 0`
- `RECEIVED`: if `1 - discount * YEARFRAC(settlement,maturity,basis) == 0`
- `PRICEMAT`: if `1 + yld * sm == 0` (defensive check; with `yld > 0` and `sm > 0` this should not occur)
- `YIELDMAT`: if `pr + accr == 0` (defensive check; with `pr > 0` and `accr >= 0` this should not occur)

All other divide-by-zero-like invalid domains are reported as `#NUM!` (e.g., non-positive implied prices).

---

## Discount security functions

All discount-security functions share:

- Date validation: `settlement < maturity`
- Basis validation: `basis ∈ 0..=4` (default `0`)
- Year fraction: `yf = YEARFRAC(settlement, maturity, basis)`

### `DISC`

Signature (Excel): `DISC(settlement, maturity, pr, redemption, [basis])`

Validations:

- `pr > 0`, `redemption > 0`
- `yf == 0` → `#DIV/0!`

Formula:

```text
DISC = (redemption - pr) / redemption / yf
```

### `PRICEDISC`

Signature (Excel): `PRICEDISC(settlement, maturity, discount, redemption, [basis])`

Validations:

- `discount > 0`, `redemption > 0`
- Let `factor = 1 - discount * yf`; require `factor > 0` (non-positive prices → `#NUM!`)

Formula:

```text
PRICEDISC = redemption * (1 - discount * yf)
```

### `YIELDDISC`

Signature (Excel): `YIELDDISC(settlement, maturity, pr, redemption, [basis])`

Validations:

- `pr > 0`, `redemption > 0`
- `yf == 0` → `#DIV/0!`

Formula:

```text
YIELDDISC = (redemption - pr) / pr / yf
```

### `INTRATE`

Signature (Excel): `INTRATE(settlement, maturity, investment, redemption, [basis])`

Validations:

- `investment > 0`, `redemption > 0`
- `yf == 0` → `#DIV/0!`

Formula:

```text
INTRATE = (redemption - investment) / investment / yf
```

### `RECEIVED`

Signature (Excel): `RECEIVED(settlement, maturity, investment, discount, [basis])`

Validations:

- `investment > 0`, `discount > 0`
- Let `denom = 1 - discount * yf`
  - `denom == 0` → `#DIV/0!`
  - `denom < 0` → `#NUM!`

Formula:

```text
RECEIVED = investment / (1 - discount * yf)
```

---

## Interest-at-maturity security functions

These functions assume a $100 face value (Excel convention), with interest paid at maturity.

Shared derived values:

```text
im = YEARFRAC(issue, maturity, basis)
is = YEARFRAC(issue, settlement, basis)
sm = YEARFRAC(settlement, maturity, basis)

fv   = 100 * (1 + rate * im)
accr = 100 * rate * is
```

### `PRICEMAT`

Signature (Excel): `PRICEMAT(settlement, maturity, issue, rate, yld, [basis])`

Validations:

- Date ordering: `issue <= settlement < maturity`
- `rate > 0`, `yld > 0`
- Let `denom = 1 + yld * sm`:
  - `denom == 0` → `#DIV/0!`

Formula:

```text
PRICEMAT = fv / (1 + yld * sm) - accr
```

### `YIELDMAT`

Signature (Excel): `YIELDMAT(settlement, maturity, issue, rate, pr, [basis])`

Validations:

- Date ordering: `issue <= settlement < maturity`
- `rate > 0`, `pr > 0`
- `sm == 0` → `#DIV/0!`
- Let `denom = pr + accr`:
  - `denom == 0` → `#DIV/0!`

Formula:

```text
YIELDMAT = (fv / (pr + accr) - 1) / sm
```

---

## Treasury bill functions

These functions use an actual-day `DSM` bounded to at most one year:

```text
DSM = maturity - settlement
Require 1 <= DSM <= 365
```

All numeric inputs must be finite, and the required ones are strictly positive.

### `TBILLPRICE`

Signature (Excel): `TBILLPRICE(settlement, maturity, discount)`

Validations:

- `discount > 0`
- `DSM` must satisfy `1 <= DSM <= 365`
- Computed price must be `> 0` (otherwise → `#NUM!`)

Formula:

```text
TBILLPRICE = 100 * (1 - discount * DSM / 360)
```

### `TBILLYIELD`

Signature (Excel): `TBILLYIELD(settlement, maturity, pr)`

Validations:

- `pr > 0`
- `DSM` must satisfy `1 <= DSM <= 365`

Formula:

```text
TBILLYIELD = (100 - pr) / pr * (360 / DSM)
```

### `TBILLEQ`

Signature (Excel): `TBILLEQ(settlement, maturity, discount)`

Validations:

- `discount > 0`
- `DSM` must satisfy `1 <= DSM <= 365`
- The implied bill price must be positive:
  - `price_factor = 1 - discount * DSM / 360`
  - Require `price_factor > 0` (otherwise → `#NUM!`)

#### Branching rule (Excel-compatible)

Excel’s `TBILLEQ` switches formulas at 182 days:

- If `DSM <= 182`: use a simple bond-equivalent yield conversion:

  ```text
  TBILLEQ = 365 * discount / (360 - discount * DSM)
  ```

  Additionally require `360 - discount * DSM > 0` (otherwise → `#NUM!`).

- If `DSM > 182`: use a semiannual-compounding bond-equivalent yield:

  ```text
  TBILLEQ = 2 * ((1 / price_factor)^(365 / (2*DSM)) - 1)
  ```

  Implementation note: the engine computes this branch via `ln` + `exp_m1` for numeric stability, but
  it is mathematically equivalent to the exponent form above.
