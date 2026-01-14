mod accrued_interest;
mod amortization;
mod bonds;
mod bonds_odd;
mod builtins;
mod builtins_accrint;
mod builtins_amortization;
mod builtins_bonds;
mod builtins_depreciation_ext;
mod builtins_french_depreciation;
mod builtins_helpers;
mod builtins_misc;
mod builtins_odd_coupon;
mod builtins_pduration;
mod builtins_schedules;
mod builtins_securities;
mod cashflows;
mod coupon_schedule;
mod depreciation;
mod duration;
mod french_depreciation;
mod iterative;
mod misc;
mod odd_coupon;
mod schedules;
mod securities;
mod time_value;

pub use accrued_interest::{accrint, accrintm};
pub use amortization::{cumipmt, cumprinc};
pub use bonds::{duration, mduration, price, yield_rate};
pub use cashflows::{irr, mirr, npv, xirr, xnpv};
pub use coupon_schedule::{coupdaybs, coupdays, coupdaysnc, coupncd, coupnum, couppcd};
pub use depreciation::{db, ddb, sln, syd, vdb};
pub use duration::pduration;
pub use french_depreciation::amordegrec as amordegrc;
pub use french_depreciation::{amordegrec, amorlinc};
pub use misc::{dollarde, dollarfr, ispmt};
pub use odd_coupon::{oddfprice, oddfyield, oddlprice, oddlyield};
pub use schedules::fvschedule;
pub use securities::{
    disc, intrate, pricedisc, pricemat, received, tbilleq, tbillprice, tbillyield, yielddisc,
    yieldmat,
};
pub use time_value::{effect, fv, ipmt, nominal, nper, pmt, ppmt, pv, rate, rri};

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. This `#[used]` anchor forces the linker to keep the
// `financial::builtins` module so its `inventory::submit!` entries are retained.
#[cfg(target_arch = "wasm32")]
#[used]
static FORCE_LINK_FINANCIAL_BUILTINS: fn() = builtins::__force_link;

// Referenced from `functions/mod.rs` to ensure the `financial` module itself is linked on wasm.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {
    builtins::__force_link();
    builtins_accrint::__force_link();
    builtins_amortization::__force_link();
    builtins_bonds::__force_link();
    builtins_depreciation_ext::__force_link();
    builtins_french_depreciation::__force_link();
    builtins_misc::__force_link();
    builtins_odd_coupon::__force_link();
    builtins_pduration::__force_link();
    builtins_schedules::__force_link();
    builtins_securities::__force_link();
}
