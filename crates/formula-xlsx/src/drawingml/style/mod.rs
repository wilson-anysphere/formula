//! Minimal DrawingML style parsing for charts/shapes.
//!
//! Excel charts embed DrawingML shape properties (`c:spPr`, `c:txPr`, `a:*`) for
//! fills/lines/markers/text. The renderer needs these to respect series and
//! per-point overrides as well as theme-based colors (`a:schemeClr`).

mod color;
mod line;
mod marker;
mod shape;
mod text;

pub use color::parse_color;
pub use line::parse_ln;
pub use marker::parse_marker;
pub use shape::{parse_solid_fill, parse_sppr};
pub use text::parse_txpr;

#[cfg(test)]
mod tests;
