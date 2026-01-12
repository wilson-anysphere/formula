# Bond / Coupon functions (Excel-compatible spec)

This document is an **internal implementation spec** for the Excel-compatible “bond math” functions:

- Coupon schedule helpers: `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`
- Accrued interest: `ACCRINT`, `ACCRINTM`
- Bond valuation: `PRICE`, `YIELD`, `DURATION`, `MDURATION`

The intent is to clearly define:

- **Validation behavior** (`#NUM!` vs valid)
- **Coupon date schedule derivation** (PCD/NCD) via month-stepping
- **Day-count basis** handling (basis `0..4`) and its impact on `A`, `DSC`, `E`
- **Clean vs dirty** price and accrued interest
- `YIELD` root-finding expectations (Newton-Raphson)

Implementation should reuse existing helpers where possible:

- Month stepping / date conversion: `crates/formula-engine/src/functions/date_time/mod.rs` (`edate`, `days360`, `yearfrac`)
- Iterative solver: `crates/formula-engine/src/functions/financial/iterative.rs` (`newton_raphson`, `EXCEL_ITERATION_TOLERANCE`)

> Scope note: `COUP*`, `PRICE`, `YIELD`, `DURATION`, `MDURATION` assume a **regular coupon schedule** (no odd first/last period). Excel provides `ODDF*` / `ODDL*` for irregular schedules; those are out of scope for this doc.

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

- Allowed values: `1`, `2`, `4`
- Anything else → `#NUM!`
- Do not accept non-integer values (Excel treats them as invalid); the function argument coercion layer should provide an integer already, but the function should defensively reject if `frequency` is not exactly one of those values.

### Allowed `basis`

Where `basis` exists:

- Allowed values: `0..=4`
- Missing `basis` defaults to `0`
- Anything else → `#NUM!`

### Numeric domain checks (non-date)

The following should be finite (`is_finite`) and within Excel-like domains:

- `rate` (coupon rate): typically `>= 0` (Excel allows `0`, rejects nonsensical values as `#NUM!`).
- `par` / `redemption`: must be `> 0` (or `#NUM!`).
- `price` (for `YIELD`): must be `> 0` (or `#NUM!`).
- `yield` (for `PRICE`/`DURATION`/`MDURATION`): must keep the per-period discount factor positive:
  - Let `d = 1 + yld / f`. Require `d > 0` (otherwise PV is undefined → `#NUM!`).

---

## Coupon schedule derivation (PCD/NCD) via month stepping

### Core rule: maturity-anchored schedule

For the `COUP*` and bond valuation functions (`PRICE`, `YIELD`, `DURATION`, `MDURATION`) we treat **maturity as a coupon date** and build the schedule by stepping in increments of `m = 12 / f` months.

Coupon dates are derived by adding/subtracting months using Excel’s `EDATE` semantics:

- Add `m` months to a date by preserving the day-of-month if possible.
- If the target month has fewer days, clamp to the last valid day-of-month.

Implementation should reuse `date_time::edate`:

- `crates/formula-engine/src/functions/date_time/mod.rs::edate(start_date, months, system)`

### Finding `PCD` and `NCD`

Definition:

- `PCD` is the coupon date `<= settlement` in the maturity-anchored schedule.
- `NCD` is the coupon date `> settlement` in the maturity-anchored schedule.

Reference pseudocode (month-stepping backward from maturity):

```text
m = 12 / frequency
ncd = maturity
loop:
  pcd = EDATE(ncd, -m)
  if pcd <= settlement < ncd:
     return (pcd, ncd)
  ncd = pcd
```

This naturally supports:

- settlement exactly on a coupon date: `PCD = settlement`, `NCD = next coupon date`
- end-of-month clamping (via `EDATE`)

### Computing `COUPNUM` (number of remaining coupons)

Once `(PCD, NCD)` is known, `COUPNUM` is the count of coupon dates `>= NCD` and `<= maturity`, stepping forward by `m` months.

Reference pseudocode:

```text
count = 0
d = NCD
while d <= maturity:
  count += 1
  d = EDATE(d, +m)
return count
```

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

For all bases, `A` and `DSC` are computed as a “day count” between two dates:

- `A = days_between(PCD, settlement, basis)`
- `DSC = days_between(settlement, NCD, basis)`

Where:

- For basis `0` and `4` use `days360`.
- For basis `1`, `2`, `3` use actual days: `end_serial - start_serial`.

### Computing `E` (days in coupon period)

`E` is **not always** `days_between(PCD, NCD, basis)` in Excel. Excel uses basis-specific conventions:

- basis `0` and `4`: `E = 360 / f` (constant)
- basis `2`: `E = 360 / f` (constant)
- basis `3`: `E = 365 / f` (constant)
- basis `1`: `E = actual_days(NCD - PCD)` (variable, depends on the coupon period)

This convention is important because for basis `2`/`3` you can have `A + DSC != E` (since `A`/`DSC` are actual days but `E` is a fixed “model year” fraction).

---

## `COUP*` function outputs

All `COUP*` helpers share the same schedule derivation and day-count definitions from the previous sections.

Given inputs `(settlement, maturity, frequency, basis)`:

- `COUPPCD(...)` returns `PCD` (date serial)
- `COUPNCD(...)` returns `NCD` (date serial)
- `COUPNUM(...)` returns `N` (count of coupons remaining, integer)
- `COUPDAYBS(...)` returns `A` (days accrued since `PCD`, integer)
- `COUPDAYSNC(...)` returns `DSC` (days until `NCD`, integer)
- `COUPDAYS(...)` returns `E` (days in coupon period containing settlement, integer)

---

## `ACCRINT`: accrued interest for a security with periodic coupons

Signature (Excel):

`ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])`

### Derived values

- `f = frequency`, `m = 12/f` months per coupon
- Coupon payment per regular period: `C = par * rate / f`
- Coupon schedule is anchored at `first_interest` (not maturity), stepping by `±m` months using `EDATE`.

### Finding the relevant coupon period (PCD/NCD) anchored at `first_interest`

If `settlement < first_interest`:

- `NCD = first_interest`
- `PCD = EDATE(first_interest, -m)` (start of the “regular” coupon period that ends at `first_interest`)

If `settlement >= first_interest`:

```text
pcd = first_interest
ncd = EDATE(first_interest, +m)
while settlement >= ncd:
  pcd = ncd
  ncd = EDATE(ncd, +m)
```

### Accrued-interest start date (`calc_method`)

Let `calc_method` default to `0`.

- For `settlement < first_interest`:
  - `calc_method = 0`: accrue from `issue`
  - `calc_method = 1`: accrue from `PCD` (regular period start before `first_interest`)
- For `settlement >= first_interest`: accrue from `PCD` (standard “since last coupon” behavior); `calc_method` is ignored.

### Day-count fraction for accrual

Let:

- `E` be the regular period length in days (per the basis conventions from earlier):
  - basis 1: `E = actual_days(NCD - PCD)` (variable)
  - basis 0/2/4: `E = 360/f`
  - basis 3: `E = 365/f`

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

- Coupon payment per period: `C = rate * redemption / f`
- Accrued interest at settlement: `AI = C * (A / E)`

(`A` and `E` are basis-dependent as described above.)

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
- Discount base: `d = 1 + k` (must be `> 0`)
- Fractional periods to next coupon: `t0 = DSC / E`

> Note: `t0` is a count of **coupon periods**, not years.

### Cash-flow present value (dirty price)

Let `C = rate * redemption / f`.

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

Use Newton-Raphson via:

- `crates/formula-engine/src/functions/financial/iterative.rs::newton_raphson`
- Convergence tolerance: `EXCEL_ITERATION_TOLERANCE` (`1e-7`)

Recommended constants (matching existing style):

- `MAX_ITER_YIELD = 100` (same as `XIRR`)
- `guess = 0.1` (10%) if no better guess is available

### Domain restrictions

Yield must keep discount base positive:

- `d(y) = 1 + y / f > 0`

If `d(y) <= 0` at any point in evaluation or derivative, treat as non-evaluable (Newton step fails → `#NUM!`).

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

If derivative is `0`, non-finite, or Newton fails to converge → `#NUM!`.

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
- `k = yld / f`, `d = 1 + k` (must be `> 0`)

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

Step backwards from maturity by 6 months (`EDATE`):

| Step | Coupon date |
|------|-------------|
| 0 | 2024-11-30 (maturity) |
| -6 mo | 2024-05-30 |
| -12 mo | 2023-11-30 |
| -18 mo | 2023-05-30 |
| -24 mo | 2022-11-30 |

`settlement=2023-05-15` lies in `(2022-11-30, 2023-05-30]`, so:

- `COUPPCD = PCD = 2022-11-30`
- `COUPNCD = NCD = 2023-05-30`

Coupons after settlement up to maturity are:

- 2023-05-30
- 2023-11-30
- 2024-05-30
- 2024-11-30 (maturity)

So:

- `COUPNUM = 4`

Day-count values for basis 0 (30/360 US):

- `E = 360 / f = 360 / 2 = 180`
- `A   = days360(2022-11-30, 2023-05-15, FALSE) = 165`
- `DSC = days360(2023-05-15, 2023-05-30, FALSE) = 15`

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

- `C = rate * redemption / f = 0.10 * 100 / 2 = 5`
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

2. **`COUPDAYS` for basis 1 vs basis 2/3**
   - Resolution:
     - basis 1: `E = actual_days(PCD→NCD)`
     - basis 2: `E = 360/f` (constant)
     - basis 3: `E = 365/f` (constant)
   - This implies `A + DSC` may not equal `E` for basis 2/3; that is expected.

3. **`ACCRINT` `calc_method` behavior**
   - Resolution: `calc_method` only affects the **first coupon period** (`settlement < first_interest`):
     - `calc_method = 0` (default): accrue from `issue`
     - `calc_method = 1`: accrue from the regular coupon period start (`PCD` relative to `first_interest`)
   - For `settlement >= first_interest`, accrued interest is computed from `PCD` relative to `settlement` (standard “since last coupon” behavior).

4. **Irregular coupon schedules**
   - Resolution: `COUP*`, `PRICE`, `YIELD`, `DURATION`, `MDURATION` assume regular schedules.
   - If maturity does not align to a regular coupon schedule, behavior is unspecified here; prefer implementing and routing such cases to `ODD*` functions once those exist, rather than trying to emulate irregular schedules in these functions.
