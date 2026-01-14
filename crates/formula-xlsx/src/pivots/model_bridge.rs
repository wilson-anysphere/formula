//! Conversions from parsed XLSX pivot parts into `formula_model::pivots` types.
//!
//! These types are intended for UI/persistence ("model") use and can include
//! formatting hints (like `ValueField.number_format`) that are not required by
//! the pivot calculation engine.

use crate::styles::StylesPart;

use super::{PivotCacheDefinition, PivotTableDefinition};

use formula_model::pivots::{AggregationType, PivotFieldRef, ValueField};

/// Convert the `<dataFields>` section of a parsed pivot table definition into
/// model `ValueField` entries.
///
/// This is a best-effort conversion that focuses on the metadata needed for the
/// Formula pivot UX and round-tripping:
/// - source field name resolution via the pivot cache definition
/// - aggregation type (`subtotal`)
/// - number format hints (`numFmtId`)
pub fn pivot_table_to_model_value_fields(
    table: &PivotTableDefinition,
    cache_def: &PivotCacheDefinition,
    styles: &StylesPart,
) -> Vec<ValueField> {
    table
        .data_fields
        .iter()
        .filter_map(|df| {
            let field_idx = df.fld? as usize;
            let source_field_name = cache_def.cache_fields.get(field_idx)?.name.clone();
            let aggregation = map_subtotal(df.subtotal.as_deref());
            let default_name = format!(
                "{} of {}",
                aggregation_display_name(aggregation),
                source_field_name
            );
            let name = df
                .name
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or(default_name);

            let number_format = df
                .num_fmt_id
                .and_then(|id| resolve_pivot_num_fmt_id(id, styles));

            let source_field: PivotFieldRef = source_field_name.into();
            Some(ValueField {
                source_field: source_field.into(),
                name,
                aggregation,
                number_format,
                show_as: None,
                base_field: None,
                base_item: None,
            })
        })
        .collect()
}

fn resolve_pivot_num_fmt_id(num_fmt_id: u32, styles: &StylesPart) -> Option<String> {
    if num_fmt_id == 0 {
        return None;
    }

    if num_fmt_id <= u16::MAX as u32 {
        let id_u16 = num_fmt_id as u16;
        if let Some(code) = styles.num_fmt_code_for_id(id_u16) {
            return Some(code.to_string());
        }
        if let Some(code) = formula_format::builtin_format_code(id_u16) {
            return Some(code.to_string());
        }
    }

    Some(format!(
        "{}{}",
        formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX,
        num_fmt_id
    ))
}

fn map_subtotal(subtotal: Option<&str>) -> AggregationType {
    let Some(subtotal) = subtotal else {
        return AggregationType::Sum;
    };

    match subtotal.to_ascii_lowercase().as_str() {
        "sum" => AggregationType::Sum,
        "count" => AggregationType::Count,
        "average" | "avg" => AggregationType::Average,
        "min" => AggregationType::Min,
        "max" => AggregationType::Max,
        "product" => AggregationType::Product,
        "countnums" => AggregationType::CountNumbers,
        "stddev" => AggregationType::StdDev,
        "stddevp" => AggregationType::StdDevP,
        "var" => AggregationType::Var,
        "varp" => AggregationType::VarP,
        _ => AggregationType::Sum,
    }
}

fn aggregation_display_name(agg: AggregationType) -> &'static str {
    match agg {
        AggregationType::Sum => "Sum",
        AggregationType::Count => "Count",
        AggregationType::Average => "Average",
        AggregationType::Min => "Min",
        AggregationType::Max => "Max",
        AggregationType::Product => "Product",
        AggregationType::CountNumbers => "CountNums",
        AggregationType::StdDev => "StdDev",
        AggregationType::StdDevP => "StdDevP",
        AggregationType::Var => "Var",
        AggregationType::VarP => "VarP",
    }
}
