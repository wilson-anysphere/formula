//! SIMD range kernels.
//!
//! Implementations use the `wide` crate for portable SIMD where available and
//! fall back to scalar logic for edge cases (NaN filtering, etc).

mod kernels;

pub use kernels::{
    add_f64, count_if_blank_as_zero_f64, count_if_f64, count_ignore_nan_f64, div_f64, max_if_f64,
    max_ignore_nan_f64, min_if_f64, min_ignore_nan_f64, mul_f64, sub_f64, sum_count_if_f64,
    sum_count_ignore_nan_f64, sum_if_f64, sum_ignore_nan_f64, sumproduct_ignore_nan_f64, CmpOp,
    NumericCriteria,
};
