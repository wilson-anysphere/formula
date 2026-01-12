pub(crate) mod complex;
pub(crate) mod special;

mod convert;

pub(crate) use convert::convert;
mod base;
mod bit;

pub(crate) use base::{
    base_from_decimal, decimal_from_text, fixed_base_to_decimal, fixed_base_to_fixed_base,
    fixed_decimal_to_fixed_base, FixedBase,
};
pub(crate) use bit::{bitand, bitlshift, bitor, bitrshift, bitxor, BIT_MAX};
