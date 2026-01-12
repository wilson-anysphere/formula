mod builtins;
mod builtins_depreciation_ext;
mod builtins_french_depreciation;
mod builtins_securities;
mod bonds_odd;
mod cashflows;
mod depreciation;
mod french_depreciation;
mod iterative;
mod securities;
mod time_value;

pub use cashflows::{irr, mirr, npv, xirr, xnpv};
pub use depreciation::{db, ddb, sln, syd, vdb};
pub use french_depreciation::{amordegrec, amorlinc};
pub use securities::{
    disc, intrate, pricedisc, pricemat, received, tbilleq, tbillprice, tbillyield, yielddisc,
    yieldmat,
};
pub use time_value::{effect, fv, ipmt, nominal, nper, pmt, ppmt, pv, rate, rri};
