//! Conversions from parsed XLSX pivot parts into `formula_engine::pivot` types.
//!
//! The `formula-xlsx` crate's core responsibility is high-fidelity import/export.
//! For in-app pivot computation we also need to turn pivot cache/table metadata
//! into the engine's self-contained pivot types.

use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotValue, SubtotalPosition,
    ValueField,
};

use super::{PivotCacheDefinition, PivotCacheValue, PivotTableDefinition};

/// Convert a parsed pivot cache (definition + record iterator) into a pivot-engine
/// source range.
///
/// The first row of the returned range is a header row constructed from
/// `def.cache_fields[*].name`.
pub fn pivot_cache_to_engine_source(
    def: &PivotCacheDefinition,
    records: impl Iterator<Item = Vec<PivotCacheValue>>,
) -> Vec<Vec<PivotValue>> {
    let mut out = Vec::new();

    out.push(
        def.cache_fields
            .iter()
            .map(|f| PivotValue::Text(f.name.clone()))
            .collect(),
    );

    for record in records {
        let mut row = Vec::with_capacity(def.cache_fields.len());
        for field_idx in 0..def.cache_fields.len() {
            let value = record
                .get(field_idx)
                .cloned()
                .unwrap_or(PivotCacheValue::Missing);
            row.push(pivot_cache_value_to_engine(value));
        }
        out.push(row);
    }

    out
}

fn pivot_cache_value_to_engine(value: PivotCacheValue) -> PivotValue {
    match value {
        PivotCacheValue::String(s) => PivotValue::Text(s),
        PivotCacheValue::Number(n) => PivotValue::Number(n),
        PivotCacheValue::Bool(b) => PivotValue::Bool(b),
        PivotCacheValue::Missing => PivotValue::Blank,
        PivotCacheValue::Error(_) => PivotValue::Blank,
        PivotCacheValue::DateTime(s) => {
            if s.is_empty() {
                PivotValue::Blank
            } else {
                PivotValue::Text(s)
            }
        }
        PivotCacheValue::Index(_) => PivotValue::Blank,
    }
}

/// Convert a parsed pivot table definition into a pivot-engine config.
///
/// This is a best-effort conversion; unsupported layout / display options are
/// ignored.
pub fn pivot_table_to_engine_config(
    table: &PivotTableDefinition,
    cache_def: &PivotCacheDefinition,
) -> PivotConfig {
    let row_fields = table
        .row_fields
        .iter()
        .filter_map(|idx| cache_def.cache_fields.get(*idx as usize))
        .map(|f| PivotField::new(f.name.clone()))
        .collect::<Vec<_>>();

    let column_fields = table
        .col_fields
        .iter()
        .filter_map(|idx| cache_def.cache_fields.get(*idx as usize))
        .map(|f| PivotField::new(f.name.clone()))
        .collect::<Vec<_>>();

    let value_fields = table
        .data_fields
        .iter()
        .filter_map(|df| {
            let field_idx = df.fld? as usize;
            let source_field = cache_def.cache_fields.get(field_idx)?.name.clone();
            let aggregation = map_subtotal(df.subtotal.as_deref());
            let name = df
                .name
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| {
                    format!(
                        "{} of {}",
                        aggregation_display_name(aggregation),
                        source_field
                    )
                });
            Some(ValueField {
                source_field,
                name,
                aggregation,
                show_as: None,
                base_field: None,
                base_item: None,
            })
        })
        .collect::<Vec<_>>();

    // Excel does not render a "Grand Total" column unless there is at least one
    // column field.
    let grand_totals = GrandTotals {
        rows: table.row_grand_totals,
        columns: table.col_grand_totals && !column_fields.is_empty(),
    };

    PivotConfig {
        row_fields,
        column_fields,
        value_fields,
        filter_fields: Vec::new(),
        calculated_fields: Vec::new(),
        calculated_items: Vec::new(),
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals,
    }
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
