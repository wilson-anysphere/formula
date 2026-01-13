use crate::date::ExcelDateSystem;
use formula_format::{DateSystem, FormatOptions, Locale, Value as FmtValue};

pub(crate) fn format_number_general_with_options(
    n: f64,
    locale: Locale,
    date_system: ExcelDateSystem,
) -> String {
    // The core engine may contain non-finite values when users construct values programmatically.
    // Excel itself does not have NaN/Infinity numbers, but we avoid surfacing a formatting error
    // string during implicit string coercion.
    if !n.is_finite() {
        return n.to_string();
    }

    let options = FormatOptions {
        locale,
        date_system: match date_system {
            // `formula-format` always uses the Lotus 1-2-3 leap-year bug behavior for the 1900
            // date system (Excel compatibility).
            ExcelDateSystem::Excel1900 { .. } => DateSystem::Excel1900,
            ExcelDateSystem::Excel1904 => DateSystem::Excel1904,
        },
    };

    formula_format::format_value(FmtValue::Number(n), None, &options).text
}
