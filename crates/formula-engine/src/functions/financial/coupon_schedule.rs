use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::functions::date_time;

const MAX_COUPON_STEPS: usize = 50_000;

pub(crate) fn validate_frequency(frequency: i32) -> ExcelResult<i32> {
    match frequency {
        1 | 2 | 4 => Ok(frequency),
        _ => Err(ExcelError::Num),
    }
}

pub(crate) fn validate_basis(basis: i32) -> ExcelResult<i32> {
    if (0..=4).contains(&basis) {
        Ok(basis)
    } else {
        Err(ExcelError::Num)
    }
}

pub(crate) fn validate_serial(serial: i32, system: ExcelDateSystem) -> ExcelResult<()> {
    let _ = crate::date::serial_to_ymd(serial, system)?;
    Ok(())
}

fn is_eom(date_serial: i32, system: ExcelDateSystem) -> ExcelResult<bool> {
    Ok(date_time::eomonth(date_serial, 0, system)? == date_serial)
}

fn shift_coupon_months_eom(
    anchor: i32,
    months: i32,
    eom: bool,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    let shifted = date_time::edate(anchor, months, system)?;
    if eom {
        date_time::eomonth(shifted, 0, system)
    } else {
        Ok(shifted)
    }
}

fn coupon_period_months(frequency: i32) -> ExcelResult<i32> {
    match frequency {
        1 => Ok(12),
        2 => Ok(6),
        4 => Ok(3),
        _ => Err(ExcelError::Num),
    }
}

/// Previous coupon date (PCD), next coupon date (NCD), and number of coupons remaining (COUPNUM).
///
/// Schedule is anchored at `maturity` and stepped backwards by `12/frequency` months using
/// `EDATE` semantics, with Excel's end-of-month pinning rule:
///
/// If `maturity` is the last day of its month, Excel treats the schedule as end-of-month (EOM)
/// and pins all coupon dates to month-end. This matters when `maturity` is month-end but not the
/// 31st (e.g. Feb 28), because repeated `EDATE` offsets preserve the 28th/30th rather than
/// restoring later month-ends (Aug 31, Nov 30, ...).
pub(crate) fn coupon_pcd_ncd_num(
    settlement: i32,
    maturity: i32,
    frequency: i32,
    system: ExcelDateSystem,
) -> ExcelResult<(i32, i32, i32)> {
    let months = coupon_period_months(frequency)?;
    let eom = is_eom(maturity, system)?;
    // IMPORTANT: Coupon schedules are anchored to `maturity` (not stepped iteratively).
    //
    // `EDATE` month-stepping is not additive due to end-of-month clamping; stepping backwards one
    // period at a time can "drift" away from month-end coupon schedules (e.g. Aug 31 -> Feb 28 ->
    // Aug 28). Excel's COUP* functions behave as if each coupon date is computed directly as an
    // offset from maturity, so we do the same here.
    for n in 1..=MAX_COUPON_STEPS {
        let n_i32 = n as i32;
        let months_back = n_i32.checked_mul(months).ok_or(ExcelError::Num)?;
        let pcd = shift_coupon_months_eom(maturity, -months_back, eom, system)?;

        let ncd = if n == 1 {
            maturity
        } else {
            let prev_n = n_i32.checked_sub(1).ok_or(ExcelError::Num)?;
            let prev_months_back = prev_n.checked_mul(months).ok_or(ExcelError::Num)?;
            shift_coupon_months_eom(maturity, -prev_months_back, eom, system)?
        };

        if pcd <= settlement && settlement < ncd {
            return Ok((pcd, ncd, n_i32));
        }
    }

    Err(ExcelError::Num)
}

/// Day-count between two dates for coupon schedule computations.
///
/// - basis 0/4: 30/360 via DAYS360
/// - basis 1/2/3: actual days via serial difference
pub(crate) fn days_between(
    start_date: i32,
    end_date: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<i64> {
    match basis {
        0 => date_time::days360(start_date, end_date, false, system),
        4 => date_time::days360(start_date, end_date, true, system),
        1 | 2 | 3 => Ok(i64::from(end_date) - i64::from(start_date)),
        _ => Err(ExcelError::Num),
    }
}

/// Coupon-period length `E` (days) following Excel-compatible conventions:
/// - basis 0/2/4: 360/frequency (constant)
///   - For basis 4, day counts like `COUPDAYBS` use European 30E/360 (`DAYS360(..., TRUE)`), but
///     Excel still models the coupon period length `E` returned by `COUPDAYS` as `360/frequency`.
///     This can therefore differ from `DAYS360(PCD, NCD, TRUE)` for some end-of-month schedules
///     involving February (see tests in `financial_coupons.rs`).
/// - basis 3: 365/frequency (constant)
/// - basis 1: actual days between PCD and NCD (variable)
pub(crate) fn coupon_period_e(
    pcd: i32,
    ncd: i32,
    frequency: i32,
    basis: i32,
    _system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let freq = f64::from(frequency);
    if !freq.is_finite() || freq <= 0.0 {
        return Err(ExcelError::Num);
    }

    let e = match basis {
        1 => (i64::from(ncd) - i64::from(pcd)) as f64,
        0 | 2 | 4 => 360.0 / freq,
        3 => 365.0 / freq,
        _ => return Err(ExcelError::Num),
    };

    if !e.is_finite() || e <= 0.0 {
        Err(ExcelError::Num)
    } else {
        Ok(e)
    }
}

fn validate_coupon_args(
    settlement: i32,
    maturity: i32,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<()> {
    if settlement >= maturity {
        return Err(ExcelError::Num);
    }
    validate_frequency(frequency)?;
    validate_basis(basis)?;
    validate_serial(settlement, system)?;
    validate_serial(maturity, system)?;
    Ok(())
}

/// COUPPCD(settlement, maturity, frequency, [basis])
pub fn couppcd(
    settlement: i32,
    maturity: i32,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    validate_coupon_args(settlement, maturity, frequency, basis, system)?;
    let (pcd, _ncd, _n) = coupon_pcd_ncd_num(settlement, maturity, frequency, system)?;
    Ok(pcd)
}

/// COUPNCD(settlement, maturity, frequency, [basis])
pub fn coupncd(
    settlement: i32,
    maturity: i32,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    validate_coupon_args(settlement, maturity, frequency, basis, system)?;
    let (_pcd, ncd, _n) = coupon_pcd_ncd_num(settlement, maturity, frequency, system)?;
    Ok(ncd)
}

/// COUPNUM(settlement, maturity, frequency, [basis])
pub fn coupnum(
    settlement: i32,
    maturity: i32,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_coupon_args(settlement, maturity, frequency, basis, system)?;
    let (_pcd, _ncd, n) = coupon_pcd_ncd_num(settlement, maturity, frequency, system)?;
    Ok(n as f64)
}

/// COUPDAYBS(settlement, maturity, frequency, [basis])
pub fn coupdaybs(
    settlement: i32,
    maturity: i32,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_coupon_args(settlement, maturity, frequency, basis, system)?;
    let (pcd, _ncd, _n) = coupon_pcd_ncd_num(settlement, maturity, frequency, system)?;
    let a = days_between(pcd, settlement, basis, system)? as f64;
    if !a.is_finite() || a < 0.0 {
        return Err(ExcelError::Num);
    }
    Ok(a)
}

/// COUPDAYSNC(settlement, maturity, frequency, [basis])
pub fn coupdaysnc(
    settlement: i32,
    maturity: i32,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_coupon_args(settlement, maturity, frequency, basis, system)?;
    let (pcd, ncd, _n) = coupon_pcd_ncd_num(settlement, maturity, frequency, system)?;
    // Excel's COUPDAYSNC is not always computed as `days_between(settlement, NCD, basis)`.
    //
    // For 30/360 bases (0=US 30/360, 4=European 30E/360), Excel computes DSC as the remaining
    // portion of the modeled coupon period:
    //   DSC = E - A
    //
    // This preserves the additivity invariant `A + DSC == E`. For basis=0, this is required
    // because `DAYS360(..., FALSE)` is not additive for some month-end schedules.
    //
    // For basis=4 (European 30E/360), keep `E` consistent with `COUPDAYS` by using
    // `coupon_period_e`, which models the coupon period length as `360/frequency` (not as
    // `DAYS360(PCD, NCD, TRUE)`).
    let dsc = match basis {
        0 => {
            let e = coupon_period_e(pcd, ncd, frequency, basis, system)?;
            let a = days_between(pcd, settlement, basis, system)? as f64;
            e - a
        }
        4 => {
            let e = coupon_period_e(pcd, ncd, frequency, basis, system)?;
            let a = days_between(pcd, settlement, basis, system)? as f64;
            e - a
        }
        _ => days_between(settlement, ncd, basis, system)? as f64,
    };
    if !dsc.is_finite() || dsc < 0.0 {
        return Err(ExcelError::Num);
    }
    Ok(dsc)
}

/// COUPDAYS(settlement, maturity, frequency, [basis])
pub fn coupdays(
    settlement: i32,
    maturity: i32,
    frequency: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_coupon_args(settlement, maturity, frequency, basis, system)?;
    let (pcd, ncd, _n) = coupon_pcd_ncd_num(settlement, maturity, frequency, system)?;
    coupon_period_e(pcd, ncd, frequency, basis, system)
}
