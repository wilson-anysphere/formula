use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::functions::date_time;

fn validate_finite_positive(n: f64) -> ExcelResult<()> {
    if !n.is_finite() || n <= 0.0 {
        return Err(ExcelError::Num);
    }
    Ok(())
}

fn validate_finite(n: f64) -> ExcelResult<()> {
    if !n.is_finite() {
        return Err(ExcelError::Num);
    }
    Ok(())
}

fn validate_basis(basis: i32) -> ExcelResult<()> {
    if !(0..=4).contains(&basis) {
        return Err(ExcelError::Num);
    }
    Ok(())
}

fn validate_settlement_maturity(settlement: i32, maturity: i32) -> ExcelResult<()> {
    if settlement >= maturity {
        return Err(ExcelError::Num);
    }
    Ok(())
}

/// DISC(settlement, maturity, pr, redemption, [basis])
pub fn disc(
    settlement: i32,
    maturity: i32,
    pr: f64,
    redemption: f64,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_settlement_maturity(settlement, maturity)?;
    validate_finite_positive(pr)?;
    validate_finite_positive(redemption)?;
    validate_basis(basis)?;

    let f = date_time::yearfrac(settlement, maturity, basis, system)?;
    if f == 0.0 {
        return Err(ExcelError::Div0);
    }

    let result = (redemption - pr) / redemption / f;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

/// PRICEDISC(settlement, maturity, discount, redemption, [basis])
pub fn pricedisc(
    settlement: i32,
    maturity: i32,
    discount: f64,
    redemption: f64,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_settlement_maturity(settlement, maturity)?;
    validate_finite_positive(discount)?;
    validate_finite_positive(redemption)?;
    validate_basis(basis)?;

    let f = date_time::yearfrac(settlement, maturity, basis, system)?;
    let factor = 1.0 - discount * f;
    if !factor.is_finite() {
        return Err(ExcelError::Num);
    }
    // Discount instruments cannot have non-positive prices.
    if factor <= 0.0 {
        return Err(ExcelError::Num);
    }

    let result = redemption * factor;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

/// YIELDDISC(settlement, maturity, pr, redemption, [basis])
pub fn yielddisc(
    settlement: i32,
    maturity: i32,
    pr: f64,
    redemption: f64,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_settlement_maturity(settlement, maturity)?;
    validate_finite_positive(pr)?;
    validate_finite_positive(redemption)?;
    validate_basis(basis)?;

    let f = date_time::yearfrac(settlement, maturity, basis, system)?;
    if f == 0.0 {
        return Err(ExcelError::Div0);
    }

    let result = (redemption - pr) / pr / f;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

/// INTRATE(settlement, maturity, investment, redemption, [basis])
pub fn intrate(
    settlement: i32,
    maturity: i32,
    investment: f64,
    redemption: f64,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_settlement_maturity(settlement, maturity)?;
    validate_finite_positive(investment)?;
    validate_finite_positive(redemption)?;
    validate_basis(basis)?;

    let f = date_time::yearfrac(settlement, maturity, basis, system)?;
    if f == 0.0 {
        return Err(ExcelError::Div0);
    }

    let result = (redemption - investment) / investment / f;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

/// RECEIVED(settlement, maturity, investment, discount, [basis])
pub fn received(
    settlement: i32,
    maturity: i32,
    investment: f64,
    discount: f64,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_settlement_maturity(settlement, maturity)?;
    validate_finite_positive(investment)?;
    validate_finite_positive(discount)?;
    validate_basis(basis)?;

    let f = date_time::yearfrac(settlement, maturity, basis, system)?;
    let denom = 1.0 - discount * f;
    if !denom.is_finite() {
        return Err(ExcelError::Num);
    }
    if denom == 0.0 {
        return Err(ExcelError::Div0);
    }
    if denom < 0.0 {
        return Err(ExcelError::Num);
    }

    let result = investment / denom;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

fn validate_issue_dates(issue: i32, settlement: i32, maturity: i32) -> ExcelResult<()> {
    if issue > settlement {
        return Err(ExcelError::Num);
    }
    if issue >= maturity {
        return Err(ExcelError::Num);
    }
    Ok(())
}

/// PRICEMAT(settlement, maturity, issue, rate, yld, [basis])
pub fn pricemat(
    settlement: i32,
    maturity: i32,
    issue: i32,
    rate: f64,
    yld: f64,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_settlement_maturity(settlement, maturity)?;
    validate_issue_dates(issue, settlement, maturity)?;
    validate_finite_positive(rate)?;
    validate_finite_positive(yld)?;
    validate_basis(basis)?;

    let im = date_time::yearfrac(issue, maturity, basis, system)?;
    let is_ = date_time::yearfrac(issue, settlement, basis, system)?;
    let sm = date_time::yearfrac(settlement, maturity, basis, system)?;

    validate_finite(im)?;
    validate_finite(is_)?;
    validate_finite(sm)?;

    let fv = 100.0 * (1.0 + rate * im);
    let accr = 100.0 * rate * is_;
    let denom = 1.0 + yld * sm;
    if denom == 0.0 {
        return Err(ExcelError::Div0);
    }
    if !fv.is_finite() || !accr.is_finite() || !denom.is_finite() {
        return Err(ExcelError::Num);
    }

    let result = fv / denom - accr;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

/// YIELDMAT(settlement, maturity, issue, rate, pr, [basis])
pub fn yieldmat(
    settlement: i32,
    maturity: i32,
    issue: i32,
    rate: f64,
    pr: f64,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    validate_settlement_maturity(settlement, maturity)?;
    validate_issue_dates(issue, settlement, maturity)?;
    validate_finite_positive(rate)?;
    validate_finite_positive(pr)?;
    validate_basis(basis)?;

    let im = date_time::yearfrac(issue, maturity, basis, system)?;
    let is_ = date_time::yearfrac(issue, settlement, basis, system)?;
    let sm = date_time::yearfrac(settlement, maturity, basis, system)?;

    validate_finite(im)?;
    validate_finite(is_)?;
    validate_finite(sm)?;
    if sm == 0.0 {
        return Err(ExcelError::Div0);
    }

    let fv = 100.0 * (1.0 + rate * im);
    let accr = 100.0 * rate * is_;
    if !fv.is_finite() || !accr.is_finite() {
        return Err(ExcelError::Num);
    }

    let denom = pr + accr;
    if denom == 0.0 {
        return Err(ExcelError::Div0);
    }
    if !denom.is_finite() {
        return Err(ExcelError::Num);
    }

    let result = (fv / denom - 1.0) / sm;
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

fn validate_tbill_dsm(settlement: i32, maturity: i32) -> ExcelResult<i32> {
    validate_settlement_maturity(settlement, maturity)?;
    let dsm = maturity - settlement;
    if !(1..=365).contains(&dsm) {
        return Err(ExcelError::Num);
    }
    Ok(dsm)
}

/// TBILLPRICE(settlement, maturity, discount)
pub fn tbillprice(settlement: i32, maturity: i32, discount: f64) -> ExcelResult<f64> {
    validate_finite_positive(discount)?;
    let dsm = validate_tbill_dsm(settlement, maturity)? as f64;

    let price = 100.0 * (1.0 - discount * dsm / 360.0);
    if !price.is_finite() {
        return Err(ExcelError::Num);
    }
    if price <= 0.0 {
        return Err(ExcelError::Num);
    }
    Ok(price)
}

/// TBILLYIELD(settlement, maturity, pr)
pub fn tbillyield(settlement: i32, maturity: i32, pr: f64) -> ExcelResult<f64> {
    validate_finite_positive(pr)?;
    let dsm = validate_tbill_dsm(settlement, maturity)? as f64;

    let result = (100.0 - pr) / pr * (360.0 / dsm);
    if result.is_finite() {
        Ok(result)
    } else {
        Err(ExcelError::Num)
    }
}

/// TBILLEQ(settlement, maturity, discount)
pub fn tbilleq(settlement: i32, maturity: i32, discount: f64) -> ExcelResult<f64> {
    validate_finite_positive(discount)?;
    let dsm_i32 = validate_tbill_dsm(settlement, maturity)?;
    let dsm = dsm_i32 as f64;

    // Price per 1 face value (100 cancels out), computed using the bill discount rate.
    let price_factor = 1.0 - discount * dsm / 360.0;
    if !price_factor.is_finite() || price_factor <= 0.0 {
        // Non-positive prices are invalid for T-bills.
        return Err(ExcelError::Num);
    }

    if dsm_i32 <= 182 {
        let denom = 360.0 - discount * dsm;
        if denom <= 0.0 || !denom.is_finite() {
            return Err(ExcelError::Num);
        }
        let result = 365.0 * discount / denom;
        if result.is_finite() {
            Ok(result)
        } else {
            Err(ExcelError::Num)
        }
    } else {
        // Bond-equivalent yield (nominal annual rate, semiannual compounding).
        //
        // We compute:
        //   y = 2 * ((1 / price_factor)^(365 / (2*dsm)) - 1)
        //
        // This matches the convention used by Excel's TBILLEQ for bills longer than 182 days.
        let exponent = 365.0 / (2.0 * dsm);
        let ln_ratio = -price_factor.ln(); // ln(1 / price_factor)
        let scaled = exponent * ln_ratio;
        if !scaled.is_finite() {
            return Err(ExcelError::Num);
        }
        let factor_minus_1 = scaled.exp_m1();
        let result = 2.0 * factor_minus_1;
        if result.is_finite() {
            Ok(result)
        } else {
            Err(ExcelError::Num)
        }
    }
}

