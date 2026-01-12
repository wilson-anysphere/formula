# Bond / Coupon functions (Excel-compatible spec)

This document is an **internal implementation spec** for the Excel-compatible “bond math” functions:

- Coupon schedule helpers: `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`
- Accrued interest: `ACCRINT`, `ACCRINTM`
- Bond valuation: `PRICE`, `YIELD`, `DURATION`, `MDURATION`

The intent is to clearly define:

- **Validation behavior** (`#NUM!` vs valid)
- **Coupon date schedule derivation** (PCD/NCD) via maturity-anchored month-stepping (including Excel’s end-of-month pinning)
- **Day-count basis** handling (basis `0..4`) and its impact on `A`, `DSC`, `E`
- **Clean vs dirty** price and accrued interest
- `YIELD` root-finding expectations (safeguarded Newton with bracketing + bisection fallback)

Implementation should reuse existing helpers where possible:

- Month stepping / date conversion: `crates/formula-engine/src/functions/date_time/mod.rs` (`edate`, `eomonth`, `days360`, `yearfrac`)
- Iterative solver: `crates/formula-engine/src/functions/financial/iterative.rs` (`solve_root_newton_bisection`, `EXCEL_ITERATION_TOLERANCE`)

> Scope note: `COUP*`, `PRICE`, `YIELD`, `DURATION`, `MDURATION` assume a **regular coupon schedule**
> (no odd first/last period). Excel provides `ODDF*` / `ODDL*` for irregular schedules; those are
> out of scope for this doc. See [`docs/financial-odd-coupon-bonds.md`](../financial-odd-coupon-bonds.md).

---

## Terminology and shared derived values

All date arguments (`settlement`, `maturity`, `issue`, `first_interest`) are Excel date serials and must represent a valid calendar date in the active `ExcelDateSystem`.

For the coupon-schedule/bond functions we use:

- `f` = `frequency` (coupons per year), must be one of `{1, 2, 4}`
- `m` = months per coupon = `12 / f` (integer)
- `PCD` = previous coupon date **on or before** settlement (previous coupon date)
- `NCD` = next coupon date **strictly after** settlement (next coupon date)
- `N` = number of coupon payments remaining after settlement **up to and including maturity**
- `A` = days from `PCD` to `settlement` (accrued days within the current coupon period)
- `DSC` = days from `settlement` to `NCD` (days to next coupon)
- `E` = days in the coupon period that contains settlement (days from `PCD` to `NCD`, or a basis-derived constant)

These values show up in Excel’s documentation and are the “glue” between `COUP*`, `PRICE`, `YIELD`, `DURATION`, and `MDURATION`.

---

## Argument validation rules

### Common date validations

For `COUP*`, `PRICE`, `YIELD`, `DURATION`, `MDURATION`:

- `settlement` and `maturity` must be valid dates
- `settlement < maturity` (strict). `settlement >= maturity` → `#NUM!`

For `ACCRINT`:

- `issue`, `first_interest`, `settlement` must be valid dates
- `issue < settlement` (strict). `issue >= settlement` → `#NUM!`
- `issue < first_interest` (strict). `issue >= first_interest` → `#NUM!`

For `ACCRINTM`:

- `issue`, `settlement` must be valid dates
- `issue < settlement` (strict). `issue >= settlement` → `#NUM!`

### Allowed `frequency`

Where `frequency` exists (`COUP*`, `ACCRINT`, `PRICE`, `YIELD`, `DURATION`, `MDURATION`):

- Excel-style coercion: numeric `frequency` inputs are **truncated toward zero** before validation (`frequency = trunc(frequency)`).
- After truncation, allowed values are: `{1, 2, 4}`. Anything else → `#NUM!`.
- Examples:
  - `2.9 → 2` (valid)
  - `1.999999999 → 1` (valid)
  - `0.9 → 0 → #NUM!` (invalid after truncation)

> Implementation note: the core Rust helpers typically accept integer `frequency` / `basis` and rely on the builtin argument-coercion layer to apply Excel-style truncation; the helpers still validate membership in `{1,2,4}` and `0..=4`.

### Allowed `basis`

Where `basis` exists:

- Missing `basis` defaults to `0`.
- Excel-style coercion: numeric `basis` inputs are **truncated toward zero** before validation (`basis = trunc(basis)`).
- After truncation, allowed values are: `0..=4`. Anything else → `#NUM!`.
- Example: `3.7 → 3` (valid), `5.0 → 5 → #NUM!`

### Numeric domain checks (non-date)

The following should be finite (`is_finite`) and within Excel-like domains:

- `rate` (coupon rate): typically `>= 0` (Excel allows `0`, rejects nonsensical values as `#NUM!`).
- `par` / `redemption`: must be `> 0` (or `#NUM!`).
- `price` (for `YIELD`): must be `> 0` (or `#NUM!`).
- `yield` (for `PRICE`/`DURATION`/`MDURATION`): must keep the per-period discount factor positive:
  - Let `d = 1 + yld / f`. Require `d > 0` (otherwise PV is undefined → `#NUM!`).
  - Excel-style error mapping:
    - `d == 0` → `#DIV/0!`
    - `d < 0` → `#NUM!`

---

## Coupon schedule derivation (PCD/NCD) via month stepping

### Core rule: maturity-anchored schedule

For the `COUP*` and bond valuation functions (`PRICE`, `YIELD`, `DURATION`, `MDURATION`) we treat **maturity as a coupon date** and build the schedule by stepping in increments of `m = 12 / f` months.

Coupon dates are derived by adding/subtracting months using Excel’s `EDATE` semantics:

- Add `m` months to a date by preserving the day-of-month if possible.
- If the target month has fewer days, clamp to the last valid day-of-month.

Implementation should reuse `date_time::edate`:

- `crates/formula-engine/src/functions/date_time/mod.rs::edate(start_date, months, system)`

### Excel end-of-month (EOM) pinning rule (as implemented)

Excel has an additional coupon-schedule rule that goes beyond plain `EDATE` month stepping:

- If `maturity` is the **last day of its month** (i.e. `maturity == EOMONTH(maturity, 0)`), Excel treats the *entire coupon schedule* as an **end-of-month schedule** and pins every coupon date to month-end.

This is implemented by computing coupon dates as an offset from maturity **and then** applying `EOMONTH(..., 0)`:

```text
coupon_date(k) = EOMONTH(EDATE(maturity, -k*m), 0)     # when maturity is EOM
coupon_date(k) = EDATE(maturity, -k*m)                # otherwise
```

This matters when `maturity` is month-end but not the 31st (e.g. Feb 28/29, Nov 30): without EOM pinning, `EDATE` offsets preserve the 28th/30th and do not “restore” later month-ends (Aug 31, Dec 31, ...).

### Finding `PCD` and `NCD`

Definition:

- `PCD` is the coupon date `<= settlement` in the maturity-anchored schedule.
- `NCD` is the coupon date `> settlement` in the maturity-anchored schedule.

Implementation note (maturity-anchored, non-drifting):

- Because `EDATE` clamps to the end of month when the target month is shorter, *iteratively*
  stepping `EDATE(EDATE(...), ...)` can drift the day-of-month (e.g. 31st → 30th). Excel’s `COUP*`
  functions behave as if each coupon date is computed as an offset from `maturity`. Additionally,
  when `maturity` is month-end, Excel pins all coupon dates to month-end (see the EOM pinning rule
  above).

Reference pseudocode (scan coupon periods back from `maturity`):

```text
m = 12 / frequency
eom = (maturity == EOMONTH(maturity, 0))
coupon_date(k) = if eom then EOMONTH(EDATE(maturity, -k*m), 0) else EDATE(maturity, -k*m)
for n in 1..:
  pcd = coupon_date(n)
  ncd = if n == 1 then maturity else coupon_date(n-1)
  if pcd <= settlement < ncd:
      return (pcd, ncd, n)  # n is COUPNUM
```

This naturally supports:

- settlement exactly on a coupon date: `PCD = settlement`, `NCD = next coupon date`
- end-of-month behavior (via `EDATE` clamping, plus EOM pinning when `maturity` is month-end)

### Computing `COUPNUM` (number of remaining coupons)

In a maturity-anchored implementation, `COUPNUM` is the same `n` returned by the `(PCD, NCD)` scan
above: the number of remaining coupon payment dates from `NCD` through `maturity`, inclusive.

This definition matches the needs of `PRICE`/`YIELD`/`DURATION` where “remaining coupon payments” includes the maturity payment.

---

## Day-count basis and its effect on A / DSC / E

### Basis table (Excel-compatible)

| basis | Name | Days between two dates | Notes / mapping to current helpers |
|------:|------|------------------------|------------------------------------|
| 0 | US (NASD) 30/360 | `DAYS360(start,end,FALSE)` | Use `date_time::days360(..., method=false)` |
| 1 | Actual/Actual | `end - start` (actual days) | For year fractions use `date_time::yearfrac(..., basis=1)` |
| 2 | Actual/360 | `end - start` (actual days) | Denominator fixed at 360 for year fractions |
| 3 | Actual/365 | `end - start` (actual days) | Denominator fixed at 365 for year fractions |
| 4 | European 30/360 | `DAYS360(start,end,TRUE)` | Use `date_time::days360(..., method=true)` |

### Computing `A` and `DSC`

`A` (days accrued since the previous coupon) is always computed as a basis-specific day count:

- basis `0`: `DAYS360(PCD, settlement, FALSE)`
- basis `4`: `DAYS360(PCD, settlement, TRUE)`
- basis `1`/`2`/`3`: `settlement - PCD` (actual days)

`DSC` (days from settlement to the next coupon) follows an Excel quirk:

- basis `0` (US/NASD 30/360): **`COUPDAYSNC` is not `DAYS360(settlement, NCD, FALSE)`**. Excel
   models `E` as a fixed `360/frequency` coupon period and defines `DSC = E - A` (so `A + DSC = E` for any
   settlement date within the coupon period).
- basis `4` (European 30/360): `A` is computed via `DAYS360(..., TRUE)`. `E` is modeled as a fixed
  `360/frequency` coupon period and Excel defines `DSC = E - A` (so `A + DSC = E` for any settlement date
  within the coupon period).
  - This means `E` can intentionally diverge from `DAYS360(PCD, NCD, TRUE)` for some February/EOM schedules.
  - And `DSC` is not always equal to `DAYS360(settlement, NCD, TRUE)` if the day-count is non-additive
    when the interval is split at month-end boundaries.
- basis `1`/`2`/`3`: `DSC = NCD - settlement` (actual days).

### Computing `E` (days in coupon period)

`E` is **not always** `days_between(PCD, NCD, basis)` in Excel. Excel uses basis-specific conventions:

- basis `0`/`2`: `E = 360 / frequency` (constant)
- basis `3`: `E = 365 / frequency` (constant)
- basis `1`: `E = actual_days(PCD..NCD)` (variable; depends on the coupon period)
- basis `4`: `E = 360 / frequency` (constant)

This convention is important because for basis `2`/`3` you can have `A + DSC != E` (since `A`/`DSC`
are actual days but `E` is a fixed “model year” fraction), while for basis `0` and `4` (30/360)
Excel keeps additivity by defining `DSC = E - A`.

---

## `COUP*` function outputs

All `COUP*` helpers share the same schedule derivation and day-count definitions from the previous sections.

Given inputs `(settlement, maturity, frequency, basis)`:

- `COUPPCD(...)` returns `PCD` (date serial)
- `COUPNCD(...)` returns `NCD` (date serial)
- `COUPNUM(...)` returns `N` (count of coupons remaining, integer)
- `COUPDAYBS(...)` returns `A` (days accrued since `PCD`; integer day count expressed as a number)
- `COUPDAYSNC(...)` returns `DSC` (days until `NCD`; integer day count expressed as a number)
- `COUPDAYS(...)` returns `E` (modeled coupon period length; for `basis=3` this is `365/f` and can be fractional)

---

## `ACCRINT`: accrued interest for a security with periodic coupons

Signature (Excel):

`ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])`

### Derived values

- `f = frequency`, `m = 12/f` months per coupon
- Coupon payment per regular period: `C = par * rate / f`
- Coupon schedule is anchored at `first_interest` (not maturity). Coupon dates are computed as offsets
  from the anchor (rather than iteratively stepping `EDATE` from the previous coupon date) to avoid
  month-end drift.
  - Excel also applies an end-of-month (EOM) pinning rule for anchored schedules: if `first_interest`
    is the last day of its month, then all coupon dates are pinned to month-end.

### Finding the relevant coupon period (PCD/NCD) anchored at `first_interest`

If `settlement < first_interest`:

- `NCD = first_interest`
- `PCD = EOMONTH(EDATE(first_interest, -m), 0)` if `first_interest` is month-end, otherwise `EDATE(first_interest, -m)`

If `settlement >= first_interest`:

```text
eom = (first_interest == EOMONTH(first_interest, 0))
coupon_date(k) = if eom then EOMONTH(EDATE(first_interest, k*m), 0) else EDATE(first_interest, k*m)

k = 0
pcd = first_interest
ncd = coupon_date(k+1)
while settlement >= ncd:
  k = k + 1
  pcd = ncd
  ncd = coupon_date(k+1)
```

### Accrued-interest start date (`calc_method`)

`calc_method` is a boolean flag. Missing `calc_method` defaults to `FALSE` (`0`).

- For `settlement < first_interest`:
  - `calc_method = 0`: accrue from `issue`
  - `calc_method = 1`: accrue from `PCD` (regular period start before `first_interest`)
- For `settlement >= first_interest`: accrue from `PCD` (standard “since last coupon” behavior); `calc_method` is ignored.

### Day-count fraction for accrual

Let:

- `E` be the regular period length in days (per the basis conventions from earlier):
   - basis 0/2: `E = 360 / frequency` (constant)
   - basis 3: `E = 365 / frequency` (constant)
   - basis 1: `E = actual_days(PCD..NCD)` (variable)
   - basis 4: `E = 360 / frequency` (constant)

Let:

- `A_start = days_between(accrual_start, settlement, basis)`

Then:

- `ACCRINT = C * (A_start / E)`

### Notes / edge cases

- `ACCRINT` returns an amount in the same unit as `par` (e.g., `par=1000` yields accrued interest in “currency units”, not “per 100”).
- For negative/invalid numeric domains (`par <= 0`, `rate < 0`, invalid `basis`, invalid `frequency`) return `#NUM!`.

---

## `ACCRINTM`: accrued interest for a security paying interest at maturity

Signature (Excel): `ACCRINTM(issue, settlement, rate, par, [basis])`

For securities that pay interest only at maturity, accrued interest is computed using a year fraction between `issue` and `settlement`:

- `yf = YEARFRAC(issue, settlement, basis)`
- `ACCRINTM = par * rate * yf`

Implementation should reuse:

- `crates/formula-engine/src/functions/date_time/mod.rs::yearfrac`

---

## Clean vs dirty price and accrued interest

### Accrued interest (periodic coupon bonds)

For regular coupon bonds, Excel’s accrued interest within the current coupon period is:

- Coupon payment per period: `C = 100 * rate / f` (rate is per 100 face value)
- Accrued interest at settlement: `AI = C * (A / E)`

(`A` and `E` are basis-dependent as described above.)

Note the scaling difference:

- `PRICE`/`YIELD` are **per 100 face value**:
  - Coupon cashflows use `100 * rate / f`.
  - `redemption` is a *separate* maturity cashflow (amount repaid per 100 face value).
- `ACCRINT` scales coupon cashflows by `par` and returns an amount in the same units as `par`.

### Clean vs dirty

- **Dirty price** (a.k.a. “full price”) includes accrued interest.
- **Clean price** excludes accrued interest.

Relationship:

- `Dirty = Clean + AI`
- Excel’s `PRICE(...)` returns **Clean**.

---

## `PRICE`: price per 100 face value (periodic coupons)

Signature (Excel): `PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])`

### Derived values

- `(PCD, NCD)` from the coupon schedule
- `N = COUPNUM(...)`
- `A`, `DSC`, `E` per basis
- Per-period yield: `k = yld / f`
- Discount base: `d = 1 + k` (must be `> 0`; `d == 0` → `#DIV/0!`, `d < 0` → `#NUM!`)
- Fractional periods to next coupon: `t0 = DSC / E`

> Note: `t0` is a count of **coupon periods**, not years.

### Cash-flow present value (dirty price)

Let `C = 100 * rate / f`.

Each remaining coupon occurs at period exponent:

- `t_i = t0 + (i - 1)` for `i = 1..N`

Dirty price is PV of coupons + redemption:

- `Dirty = Σ_{i=1..N} [ C / d^{t_i} ] + redemption / d^{t_N}`

Excel’s `PRICE` returns:

- `PRICE = Dirty - AI`
- `AI = C * (A / E)`

### Special case: `N = 1`

If there is only one remaining payment date (maturity is the next coupon date), the PV reduces to a single discounted cash flow:

- `Dirty = (C + redemption) / d^{t0}`
- `PRICE = Dirty - C * (A / E)`

### Implementation notes

- For `N > 1`, coupons can be computed with a loop, or with a geometric series.
- Prefer numeric stability:
  - Use `powf` / `powi` carefully; ensure `d > 0` before exponentiation.
  - Reject non-finite intermediate results as `#NUM!`.

---

## `YIELD`: solve yield from a clean price

Signature (Excel): `YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])`

`YIELD` inverts `PRICE`. Define:

```text
f(y) = PRICE(settlement, maturity, rate, y, redemption, frequency, basis) - pr
```

We solve `f(y) = 0`.

### Root-finding method

Use a safeguarded Newton-Raphson solver (Newton + bracketing / bisection) via:

- `crates/formula-engine/src/functions/financial/iterative.rs::solve_root_newton_bisection`
- Convergence tolerance: `EXCEL_ITERATION_TOLERANCE` (`1e-7`)

Recommended constants (matching existing style):

- `MAX_ITER_YIELD = 100` (same as `XIRR`)
- `guess = rate` when `rate > 0`, otherwise `0.1` (10%)
- Bracketing bounds: `lower = -f + 1e-8`, `upper = 1e10`

### Domain restrictions

Yield must keep discount base positive:

- `d(y) = 1 + y / f > 0`

If `d(y) <= 0` at any point in evaluation or derivative, treat as non-evaluable:

- `d(y) == 0` → `#DIV/0!`
- `d(y) < 0` → `#NUM!`

### Analytic derivative (recommended)

Accrued interest `AI` is independent of `y`, so:

- `d/dy PRICE(y) = d/dy Dirty(y)`

Let each discounted cash flow have period exponent `t` (in coupon periods) and amount `CF`:

- `PV = CF / d^t`
- `dPV/dy = -(t / f) * CF / d^{t+1}`

So:

```text
d/dy Dirty(y) = - Σ (t / f) * CF / d^{t+1}
```

Use the same `t_i` exponents as in the `PRICE` section (`t0 + i - 1`), with `CF = C` for coupons and `CF = redemption` at `t_N` (or combine `C+redemption` at `t0` for `N=1`).

If derivative is `0`/non-finite (preventing Newton steps) or the safeguarded solver fails to converge → `#NUM!`.

---

## `DURATION` and `MDURATION` (Macaulay and modified duration)

Signatures (Excel):

- `DURATION(settlement, maturity, coupon, yld, frequency, [basis])`
- `MDURATION(settlement, maturity, coupon, yld, frequency, [basis])`

These functions do **not** take `price` or `redemption`; Excel assumes:

- `redemption = 100`
- Coupon payment per period: `C = coupon * 100 / f`

### Core computation (shared)

Use the same schedule/`A`/`DSC`/`E` derivation as `PRICE`:

- `(PCD, NCD)`, `N`, `DSC`, `E`
- `t0 = DSC / E`
- `k = yld / f`, `d = 1 + k` (`d == 0` → `#DIV/0!`, `d < 0` → `#NUM!`)

Compute:

- Present value (dirty): `P = Σ PV(CF_i)`
- Weighted PV in **years**: `W = Σ (time_years_i * PV(CF_i))`

Where for `i = 1..N`:

- `t_periods_i = t0 + (i - 1)`
- `time_years_i = t_periods_i / f`
- `CF_i = C` for `i < N`
- `CF_N = C + 100`
- `PV(CF_i) = CF_i / d^{t_periods_i}`

Then:

- `DURATION = W / P`
- `MDURATION = DURATION / d`

### Special case: `N = 1`

If there is only one remaining cash flow date, then:

- `P = PV(CF_1)`
- `W = time_years_1 * PV(CF_1)`

So:

- `DURATION = time_years_1 = (DSC / E) / f`
- `MDURATION = DURATION / (1 + yld / f)`

This is useful both as an optimization and as a correctness check in tests.

---

## Worked examples (numeric, hand-computable)

The goal of these is to make the “shape” of the Excel algorithms concrete and to serve as future unit-test fixtures.

> Date arithmetic below is expressed in calendar dates; implementations operate on serial dates.

### Example 1 — COUP schedule derivation (PCD/NCD/NUM) + day counts

Inputs:

- `settlement = 2023-05-15`
- `maturity   = 2024-11-30`
- `frequency  = 2`  ⇒ `m = 6 months`
- `basis      = 0` (US 30/360)

Because `maturity` is month-end, the coupon schedule is treated as end-of-month (EOM) and pinned via `EOMONTH(EDATE(...), 0)`. Step backwards from maturity by 6 months:

| Step | Coupon date |
|------|-------------|
| 0 | 2024-11-30 (maturity) |
| -6 mo | 2024-05-31 |
| -12 mo | 2023-11-30 |
| -18 mo | 2023-05-31 |
| -24 mo | 2022-11-30 |

`settlement=2023-05-15` lies in `(2022-11-30, 2023-05-31]`, so:

- `COUPPCD = PCD = 2022-11-30`
- `COUPNCD = NCD = 2023-05-31`

Coupons after settlement up to maturity are:

- 2023-05-31
- 2023-11-30
- 2024-05-31
- 2024-11-30 (maturity)

So:

- `COUPNUM = 4`

Day-count values for basis 0 (30/360 US):

- `E = 360 / f = 360 / 2 = 180`
- `A   = days360(2022-11-30, 2023-05-15, FALSE) = 165`
- `DSC = E - A = 15` (Excel quirk for `COUPDAYSNC` on 30/360 bases)

So the corresponding coupon helpers should return:

- `COUPDAYS   = 180`
- `COUPDAYBS  = 165`
- `COUPDAYSNC = 15`

### Example 2 — `PRICE` with `N=2` and zero accrued interest

Choose settlement on a coupon date to make `A = 0` and avoid accrued-interest bookkeeping.

Inputs:

- `settlement  = 2024-01-01` (also a coupon date)
- `maturity    = 2025-01-01`
- `rate        = 10%`
- `yld         = 12%`
- `redemption  = 100`
- `frequency   = 2`
- `basis       = 0` (basis doesn’t affect this example because `A=0` and dates align to periods)

Derived values:

- `C = 100 * rate / f = 100 * 0.10 / 2 = 5`
- Per-period yield `k = 0.12 / 2 = 0.06`
- `d = 1 + k = 1.06`
- Coupon dates after settlement: 2024-07-01 and 2025-01-01 ⇒ `N = 2`
- With settlement at `PCD`, `A = 0`, so `AI = 0` and `PRICE = Dirty`.

Present value:

- First coupon (t=1): `5 / 1.06 = 4.716981...`
- Final coupon + redemption (t=2): `105 / 1.06^2 = 105 / 1.1236 = 93.449626...`

So:

- `PRICE ≈ 4.716981 + 93.449626 = 98.166607`

### Example 3 — `DURATION`/`MDURATION` when `N=1` (single remaining cash flow date)

Inputs:

- `settlement = 2024-04-01`
- `maturity   = 2024-07-01` (next coupon date == maturity)
- `coupon     = 10%` (does not affect Macaulay duration when `N=1`)
- `yld        = 12%`
- `frequency  = 2`
- `basis      = 0` (30/360)

With semiannual coupons, the surrounding coupon dates are:

- `PCD = 2024-01-01`
- `NCD = 2024-07-01 = maturity`
- So `N = 1`

Day counts (basis 0):

- `E   = 360/f = 180`
- `DSC = days360(2024-04-01, 2024-07-01, FALSE) = 90`
- `t0 = DSC/E = 90/180 = 0.5` coupon periods

Convert to years:

- `time_years = t0 / f = 0.5 / 2 = 0.25 years`

Therefore:

- `DURATION = 0.25`
- `MDURATION = DURATION / (1 + yld/f) = 0.25 / 1.06 = 0.235849...`

This case is a good unit test because it avoids summations and is independent of the coupon rate.

---

## Known ambiguous areas (and chosen resolution)

Excel’s bond functions are historically underspecified. The following are known sources of mismatch between implementations; we document the intended choice for this codebase.

1. **Whether `NCD` is “on or after” vs “strictly after” settlement**
   - Resolution: `NCD` is **strictly after** settlement; `PCD` is **on or before** settlement.
   - This matches typical Excel interpretations and makes `settlement` on a coupon date yield `A=0`.

2. **`COUPDAYS` (`E`) conventions by basis**
   - Resolution:
       - basis 0/2: `E = 360 / frequency` (constant)
       - basis 3: `E = 365 / frequency` (constant)
       - basis 1: `E = actual_days(PCD..NCD)` (variable)
       - basis 4: `E = 360 / frequency` (constant)
   - This implies `A + DSC` may not equal `E` for basis 2/3; that is expected.

3. **`ACCRINT` `calc_method` behavior**
   - Resolution: `calc_method` only affects the **first coupon period** (`settlement < first_interest`):
     - `calc_method = 0` (default): accrue from `issue`
     - `calc_method = 1`: accrue from the regular coupon period start (`PCD` relative to `first_interest`)
   - For `settlement >= first_interest`, accrued interest is computed from `PCD` relative to `settlement` (standard “since last coupon” behavior).

4. **Irregular coupon schedules**
   - Resolution: `COUP*`, `PRICE`, `YIELD`, `DURATION`, `MDURATION` assume regular schedules.
   - If maturity does not align to a regular coupon schedule, behavior is unspecified here; prefer implementing and routing such cases to `ODD*` functions once those exist, rather than trying to emulate irregular schedules in these functions.

5. **End-of-month (EOM) coupon schedule pinning**
   - Resolution: if `maturity` is month-end, treat the entire coupon schedule as month-end pinned:
     `coupon_date(k) = EOMONTH(EDATE(maturity, -k*m), 0)`.
   - This matches Excel and avoids day-of-month drift on month-end schedules.
