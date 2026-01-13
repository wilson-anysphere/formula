//! Excel odd-coupon bond functions (`ODDF*` / `ODDL*`)
//!
//! This module is intentionally documentation-first: Excel’s odd-coupon bond pricing functions
//! have a dense set of conventions (day-count basis, coupon schedules, accrued interest, and
//! solver behavior) and regress easily when refactoring. Keep any future implementation aligned
//! with the semantics described here.
//!
//! See also developer documentation:
//! `docs/financial-odd-coupon-bonds.md`.
//!
//! ---
//!
//! ## Functions covered
//!
//! - `ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, [basis])`
//! - `ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, [basis])`
//! - `ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, [basis])`
//! - `ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency, [basis])`
//!
//! The `ODDF*` pair handles an **odd (irregular) first coupon period**, while `ODDL*` handles an
//! **odd last coupon period**.
//!
//! These functions return a price / yield **per 100 face value**, consistent with Excel’s other
//! bond functions (e.g. `PRICE`, `YIELD`, `PRICEDISC`, …).
//!
//! ---
//!
//! ## Shared terminology and conventions
//!
//! **Dates**
//!
//! - Inputs are Excel serial dates (floating point). Like Excel, implementations should
//!   interpret them as dates by applying `floor()` (discard time-of-day).
//! - All day differences are computed in “serial day” units; as a result, results should be
//!   invariant under the workbook date system (Excel 1900 vs 1904) as long as all input dates are
//!   in the same system (see `crates/formula-engine/tests/functions/financial_odd_coupon.rs`).
//!
//! **Frequency**
//!
//! `frequency` must be one of `{1, 2, 4}` corresponding to annual / semiannual / quarterly
//! coupons. Any other value is a `#NUM!` in Excel.
//!
//! **Day-count basis mapping**
//!
//! Excel uses the same `basis` encoding as `YEARFRAC`:
//!
//! | basis | Convention | Notes |
//! |------:|------------|-------|
//! | 0 | US (NASD) 30/360 | See `date_time::days360(..., method=false)` |
//! | 1 | Actual/Actual | “Anniversary” method; see `date_time::yearfrac` for similar edge cases |
//! | 2 | Actual/360 | |
//! | 3 | Actual/365 | |
//! | 4 | European 30/360 | See `date_time::days360(..., method=true)` |
//!
//! Any `basis` outside `0..=4` is `#NUM!` in Excel.
//!
//! **Coupon period length `E`**
//!
//! Many Excel bond formulas are expressed in terms of:
//!
//! - `E`: the day-count length of a *regular* coupon period.
//! - `A`: accrued days from the start of the current accrual period to settlement.
//! - `DSC`: days from settlement to the next coupon date.
//!
//! For odd-coupon functions, `E` is still defined as the length of a *regular* coupon period at
//! the given `frequency` (even though the first/last period is irregular).
//!
//! Implementation note: computing `E` requires generating the regular coupon schedule (see “EOM
//! stepping” below) and applying Excel’s basis-specific coupon-period-length conventions (which
//! match `COUPDAYS`).
//!
//! For reference, this engine matches `coupon_schedule::coupon_period_e`:
//!
//! - basis `0` / `2` / `4`: `E = 360 / frequency` (constant)
//!   - Note: for `basis=4`, day counts are still European 30E/360 (`DAYS360(..., TRUE)`), but Excel
//!     models the period length `E` (as used in `COUPDAYS` and these odd-coupon functions) as a
//!     fixed `360/frequency`. This can therefore differ from `DAYS360(PCD, NCD, TRUE)` for some
//!     end-of-month schedules involving February.
//! - basis `3`: `E = 365 / frequency` (constant)
//! - basis `1`: `E = ncd - pcd` (actual days between regular coupon dates)
//!
//! **Coupon amount**
//!
//! Let:
//!
//! - `R` = `redemption` (typically 100)
//! - `C` = regular coupon payment per 100 face value = `100 * rate / frequency`
//! - `y` = yield per period = `yld / frequency`
//!
//! In odd periods, Excel prorates coupon cashflows linearly with `E`:
//!
//! - Odd-period coupon cashflow = `C * (D / E)` where `D` is the day-count length of the odd
//!   accrual period (first or last).
//!
//! The **accrued interest** term in Excel’s clean-price functions remains `C * (A / E)`, which
//! matches prorating the odd coupon and then taking the elapsed fraction:
//!
//! `C * (D/E) * (A/D) == C * (A/E)`.
//!
//! **Yield domain (Excel semantics)**
//!
//! Excel’s bond functions model discounting using the per-period base:
//!
//! ```text
//! g = 1 + (yld / frequency)
//! ```
//!
//! This means that yields can be negative (even below -100%) as long as `g > 0`, i.e. `yld > -frequency`.
//! At the exact boundary `yld == -frequency`, the discount base becomes 0 and Excel returns `#DIV/0!`.
//!
//! ---
//!
//! ## Coupon date schedule and EOM stepping
//!
//! Excel’s coupon date functions (`COUP*`) use a specific coupon schedule rule that differs from
//! naïve `EDATE` stepping:
//!
//! - Coupon dates advance in steps of `12 / frequency` months.
//! - If the schedule is an “end-of-month” (EOM) schedule, dates are pinned to the last day of the
//!   month (e.g. 31st → 30th/28th/29th → 31st).
//!
//! The EOM determination is an Excel quirk and should be treated as part of the compatibility
//! surface.
//!
//! Implementation note (this repo):
//!
//! - For **ODDF***, the regular coupon schedule is **maturity-anchored**: we generate coupon dates
//!   by stepping **backward from `maturity`** in whole coupon periods (`12 / frequency` months).
//!   This matches Excel’s `COUP*` behavior: month stepping is anchored on `maturity`, not on
//!   `first_coupon` (which may be clamped in shorter months).
//! - **ODDF* EOM schedule detection:** the schedule is treated as EOM iff `maturity` itself is
//!   end-of-month. When `maturity` is EOM, month stepping uses `date_time::eomonth` (Excel’s
//!   `EOMONTH`) rather than `date_time::edate`, ensuring non-drifting EOM schedules.
//! - This logic lives in `crates/formula-engine/src/functions/financial/odd_coupon.rs`, notably the
//!   helpers `coupon_schedule_from_maturity` (schedule generation) and `coupon_date_with_eom` (EOM vs
//!   non-EOM stepping).
//! - For **ODDL***, we only need the prior regular coupon date before `last_interest` to compute `E`;
//!   we step backward from `last_interest` by `12 / frequency` months. If `last_interest` is EOM,
//!   the step is pinned to EOM (equivalent to `EOMONTH` stepping).
//!
//! ---
//!
//! ## ODDFPRICE / ODDFYIELD (odd first coupon)
//!
//! Inputs:
//!
//! - `settlement` (`S`): settlement date (purchase date)
//! - `maturity` (`M`): maturity date (redemption)
//! - `issue` (`I`): issue date (start of accrual)
//! - `first_coupon` (`F`): first coupon payment date
//!
//! Typical validation (Excel-style `#NUM!`):
//!
//! - Chronology (current engine behavior): `I <= S <= F <= M` with `I < F` and `S < M`.
//!   - `F == M` is allowed (single odd coupon + redemption).
//!   - Boundary equalities `I == S` (zero accrued interest) and `S == F` (settlement on first coupon)
//!     are allowed (see `crates/formula-engine/tests/odd_coupon_date_boundaries.rs`).
//!   - The equality `I == F` is rejected with `#NUM!` (see the same test file).
//!   - Note: Microsoft’s `WorksheetFunction` docs describe strict inequalities; confirm actual Excel
//!     boundary behavior via the Excel oracle harness before changing these rules.
//! - `rate >= 0`, `yld` (or `pr`) finite, `redemption > 0`
//! - `frequency ∈ {1,2,4}`, `basis ∈ 0..=4`
//!
//! ### Pricing model (clean price)
//!
//! Define the day-count quantities (under `basis`):
//!
//! - `A = days(I, S)` accrued days from issue to settlement
//! - `DFC = days(I, F)` days in the odd first accrual period (issue → first coupon)
//! - `DSC = days(S, F)` days from settlement to first coupon
//!   - For additive day-count conventions (the actual/actual bases `1/2/3`), `DSC == DFC - A`.
//!   - For 30/360 bases (`0/4`), Excel’s 30/360 day-count is **not strictly additive** across
//!     intermediate dates (because end-of-month adjustments depend on the start date). In that
//!     case, treat `DSC` as its own quantity computed directly from `(S, F)`.
//! - `E = days(prev_coupon(F), F)` days in a regular coupon period
//! - `N = count_regular_coupons(F..=M)` number of coupon dates from `F` through `M` inclusive
//!
//! Cashflows:
//!
//! - At `F`: odd first coupon = `C * (DFC / E)`
//! - At each subsequent regular coupon date: `C`
//! - At `M`: `C + R` (final coupon + redemption)
//!
//! Discounting:
//!
//! - Per-period discount base: `1 + y` where `y = yld / frequency`
//!   - Domain: require `1 + y > 0` (equivalently `yld > -frequency`).
//!   - Excel-style errors: `yld == -frequency` → `#DIV/0!`, `yld < -frequency` → `#NUM!`.
//! - Exponent for a cashflow `j` regular periods after `F` is:
//!   - `j + (DSC / E)` (so the first coupon uses `j=0`)
//!
//! Clean price per 100 face:
//!
//! ```text
//! PV = ( C*(DFC/E) ) / (1+y)^(DSC/E)
//!    + Σ_{j=1..N-2} C / (1+y)^(j + DSC/E)
//!    + (C + R) / (1+y)^((N-1) + DSC/E)
//!
//! ODDFPRICE = PV - C*(A/E)
//! ```
//!
//! Where the summation is empty if there are no intermediate coupons (i.e. `N` small).
//!
//! ### Yield solver
//!
//! `ODDFYIELD` solves for `yld` such that `ODDFPRICE(yld) == pr`.
//!
//! Excel is tolerant of difficult regions (low price / high yield) where Newton steps can jump
//! out of the valid domain (`1+y <= 0`). A robust strategy is:
//!
//! 1. Newton-Raphson on the price equation using an analytic derivative (fast path).
//! 2. If Newton fails to converge or exits the domain, fall back to a bracketed method (bisection
//!    / secant) on a conservative yield interval.
//!
//! This repo already contains a Newton helper at
//! `functions::financial::iterative::newton_raphson`; document any fallback behavior alongside the
//! implementation because it affects parity.
//!
//! ---
//!
//! ## ODDLPRICE / ODDLYIELD (odd last coupon)
//!
//! Inputs:
//!
//! - `settlement` (`S`): settlement date
//! - `maturity` (`M`): maturity date
//! - `last_interest` (`L`): last coupon payment date (start of the odd last accrual period)
//!
//! Typical validation (Excel-style `#NUM!`):
//!
//! - Chronology:
//!   - `S < M` and `L < M` (both strict; settlement may be **before, on, or after** `L`).
//!   - Boundary equalities `S == M` and `L == M` are rejected with `#NUM!`. The equality `S == L`
//!     is allowed (see `crates/formula-engine/tests/functions/financial_odd_coupon.rs`).
//! - `rate >= 0`, `yld` (or `pr`) finite, `redemption > 0`
//! - `frequency ∈ {1,2,4}`, `basis ∈ 0..=4`
//!
//! ### Pricing model (clean price)
//!
//! Note: the formulas below describe the common case `S >= L`. For `S < L`, see
//! `crates/formula-engine/src/functions/financial/odd_coupon.rs::oddl_equation`.
//!
//! Day-count quantities:
//!
//! - `DLM = days(L, M)` days in the odd last accrual period (last_interest → maturity)
//! - `E = days(prev_coupon(L), L)` days in a regular coupon period
//!
//! If `S >= L` (settlement inside the odd last accrual period), also define:
//!
//! - `A = days(L, S)` accrued days in the last period
//! - `DSM = days(S, M)` days from settlement to maturity
//!   - For additive day-count conventions (the actual/actual bases `1/2/3`), `DSM == DLM - A`.
//!   - For 30/360 bases (`0/4`), Excel’s 30/360 day-count is **not strictly additive** across
//!     intermediate dates; compute `DSM` directly from `(S, M)`.
//!
//! If `S < L`, pricing must include the remaining regular coupon payments through `L` (inclusive),
//! plus the final odd-stub payment at `M`. Accrued interest is computed from the regular coupon
//! period containing settlement.
//!
//! Cashflows:
//!
//! - At `M`: odd last coupon + redemption = `R + C*(DLM/E)` (always).
//! - If `S < L`: additional regular coupon payments of `C` at each coupon date from the next
//!   coupon date after settlement through `L` (inclusive).
//!
//! Discounting:
//!
//! - Per-period discount base: `1 + y` where `y = yld / frequency`:
//!   - Domain: require `1 + y > 0` (equivalently `yld > -frequency`).
//!   - Excel-style errors: `yld == -frequency` → `#DIV/0!`, `yld < -frequency` → `#NUM!`.
//! - If `S >= L`, the exponent is `DSM / E`.
//! - If `S < L`, discount regular coupon payments using `PRICE`-style exponents `(DSC/E) + j`, and
//!   discount the final maturity cashflow by the time to `L` plus the odd stub length `DLM/E`.
//!
//! Clean price per 100 face:
//!
//! For `S >= L`:
//!
//! ```text
//! PV = ( R + C*(DLM/E) ) / (1+y)^(DSM/E)
//! ODDLPRICE = PV - C*(A/E)
//! ```
//!
//! ### Yield solver
//!
//! `ODDLYIELD` solves `ODDLPRICE(yld) == pr` with the same solver strategy notes as `ODDFYIELD`.
//!
//! ---
//!
//! ## Known Excel quirks worth preserving
//!
//! - **Date coercion:** Excel floors serial dates; do not round.
//! - **Date system invariance:** computations should rely on day differences, not absolute serial
//!   values.
//! - **EOM schedule:** coupon schedule generation must match Excel’s EOM behavior.
//! - **Error taxonomy:** invalid inputs return `#NUM!` vs `#VALUE!` in Excel in subtle ways; tests
//!   should pin the behavior (see `financial_odd_coupon.rs`).
//! - **Solver robustness:** yield solvers must behave reasonably for extreme prices/yields and not
//!   return spurious `#NUM!` where Excel converges.
