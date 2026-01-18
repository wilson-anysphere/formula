mod a1;
mod filter;
mod parse;
mod sort;
mod types;
mod visibility;
mod worksheet;

pub use a1::{parse_a1_range, to_a1_range, A1ParseError};
pub use filter::{
    apply_autofilter, apply_autofilter_with_value_locale, AutoFilter, ColumnFilter, DateComparison,
    FilterCriterion, FilterError, FilterJoin, FilterResult, FilterValue, FilterViewId, FilterViews,
    ModelAutoFilterError,
    NumberComparison, TextMatch, TextMatchKind,
};
pub use sort::{
    sort_range, sort_range_with_value_locale, RowPermutation, SortError, SortKey, SortOrder,
    SortSpec, SortValueType,
};
pub use types::{CellValue, HeaderOption, RangeData, RangeDataError, RangeRef};
pub use visibility::{HiddenRows, RowVisibility};
pub use worksheet::{apply_autofilter_to_outline, apply_autofilter_to_outline_with_value_locale};
pub use worksheet::{sort_worksheet_range, sort_worksheet_range_with_value_locale};
