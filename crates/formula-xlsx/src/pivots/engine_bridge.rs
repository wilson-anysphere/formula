//! Conversions from parsed XLSX pivot parts into `formula_engine::pivot` types.
//!
//! The `formula-xlsx` crate's core responsibility is high-fidelity import/export.
//! For in-app pivot computation we also need to turn pivot cache/table metadata
//! into the engine's self-contained pivot types.

use formula_engine::pivot::{
    AggregationType, FilterField, GrandTotals, Layout, PivotConfig, PivotField, PivotValue,
    ShowAsType, SubtotalPosition, ValueField,
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

            let show_as = map_show_data_as(df.show_data_as.as_deref());

            // `dataField@baseField` is an index into the cache fields list.
            let base_field = df.base_field.and_then(|base_field_idx| {
                cache_def
                    .cache_fields
                    .get(base_field_idx as usize)
                    .map(|f| f.name.clone())
            });

            // `dataField@baseItem` refers to an item within `baseField`'s shared-items table.
            //
            // We currently do not parse shared items from `pivotCacheDefinition`, so we can't map
            // this index to a stable display string yet.
            let base_item = None;
            Some(ValueField {
                source_field,
                name,
                aggregation,
                show_as,
                base_field,
                base_item,
            })
        })
        .collect::<Vec<_>>();

    let filter_fields = table
        .page_fields
        .iter()
        .filter_map(|idx| {
            let field_idx = *idx as usize;
            let source_field = cache_def.cache_fields.get(field_idx)?.name.clone();
            Some(FilterField {
                source_field,
                allowed: None,
            })
        })
        .collect::<Vec<_>>();

    // Excel does not render a "Grand Total" column unless there is at least one
    // column field.
    let grand_totals = GrandTotals {
        rows: table.row_grand_totals,
        columns: table.col_grand_totals && !column_fields.is_empty(),
    };

    let layout = if table.compact == Some(true) {
        Layout::Compact
    } else {
        Layout::Tabular
    };

    let subtotals = match table.subtotal_location.as_deref() {
        Some(v) if v.eq_ignore_ascii_case("AtTop") => SubtotalPosition::Top,
        Some(v) if v.eq_ignore_ascii_case("AtBottom") => SubtotalPosition::Bottom,
        _ => SubtotalPosition::None,
    };

    PivotConfig {
        row_fields,
        column_fields,
        value_fields,
        filter_fields,
        calculated_fields: Vec::new(),
        calculated_items: Vec::new(),
        layout,
        subtotals,
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

fn map_show_data_as(show_data_as: Option<&str>) -> Option<ShowAsType> {
    let show_data_as = show_data_as?.trim();
    if show_data_as.is_empty() {
        return None;
    }

    match show_data_as.to_ascii_lowercase().as_str() {
        "normal" => Some(ShowAsType::Normal),
        "percentofgrandtotal" => Some(ShowAsType::PercentOfGrandTotal),
        "percentofrowtotal" => Some(ShowAsType::PercentOfRowTotal),
        "percentofcolumntotal" => Some(ShowAsType::PercentOfColumnTotal),
        "percentof" => Some(ShowAsType::PercentOf),
        "percentdifferencefrom" => Some(ShowAsType::PercentDifferenceFrom),
        "runningtotal" => Some(ShowAsType::RunningTotal),
        "rankascending" => Some(ShowAsType::RankAscending),
        "rankdescending" => Some(ShowAsType::RankDescending),
        _ => None,
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

#[cfg(test)]
mod tests {
    use super::*;

    use crate::pivots::PivotCacheField;

    use pretty_assertions::assert_eq;

    #[test]
    fn map_show_data_as_handles_known_strings_case_insensitively() {
        let cases = [
            ("normal", Some(ShowAsType::Normal)),
            ("Normal", Some(ShowAsType::Normal)),
            ("percentOfGrandTotal", Some(ShowAsType::PercentOfGrandTotal)),
            ("PERCENTOFGRANDTOTAL", Some(ShowAsType::PercentOfGrandTotal)),
            ("percentOfRowTotal", Some(ShowAsType::PercentOfRowTotal)),
            ("percentOfColumnTotal", Some(ShowAsType::PercentOfColumnTotal)),
            ("percentOf", Some(ShowAsType::PercentOf)),
            ("percentDifferenceFrom", Some(ShowAsType::PercentDifferenceFrom)),
            ("runningTotal", Some(ShowAsType::RunningTotal)),
            ("rankAscending", Some(ShowAsType::RankAscending)),
            ("rankDescending", Some(ShowAsType::RankDescending)),
            ("unknownValue", None),
            ("", None),
            ("   ", None),
        ];

        for (raw, expected) in cases {
            assert_eq!(map_show_data_as(Some(raw)), expected, "showDataAs={raw:?}");
        }
        assert_eq!(map_show_data_as(None), None);
    }

    #[test]
    fn maps_show_data_as_percent_of_grand_total() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataFields count="1">
    <dataField fld="0" showDataAs="percentOfGrandTotal"/>
  </dataFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");
        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Sales".to_string(),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(cfg.value_fields.len(), 1);
        assert_eq!(
            cfg.value_fields[0].show_as,
            Some(ShowAsType::PercentOfGrandTotal)
        );
    }

    #[test]
    fn maps_base_field_to_cache_field_name() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataFields count="1">
    <dataField fld="0" showDataAs="percentOf" baseField="1"/>
  </dataFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");
        let cache_def = PivotCacheDefinition {
            cache_fields: vec![
                PivotCacheField {
                    name: "Sales".to_string(),
                    ..Default::default()
                },
                PivotCacheField {
                    name: "Region".to_string(),
                    ..Default::default()
                },
            ],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(cfg.value_fields.len(), 1);
        assert_eq!(cfg.value_fields[0].base_field.as_deref(), Some("Region"));
    }

    #[test]
    fn pivot_table_to_engine_config_maps_layout_subtotals_and_page_fields() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1"
  cacheId="1"
  compact="1"
  subtotalLocation="AtTop">
  <pageFields count="2">
    <pageField fld="0"/>
    <pageField fld="2"/>
  </pageFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");

        let cache_def = PivotCacheDefinition {
            cache_fields: vec![
                PivotCacheField {
                    name: "Region".to_string(),
                    ..Default::default()
                },
                PivotCacheField {
                    name: "Product".to_string(),
                    ..Default::default()
                },
                PivotCacheField {
                    name: "Sales".to_string(),
                    ..Default::default()
                },
            ],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);

        assert_eq!(cfg.layout, Layout::Compact);
        assert_eq!(cfg.subtotals, SubtotalPosition::Top);
        assert_eq!(
            cfg.filter_fields,
            vec![
                FilterField {
                    source_field: "Region".to_string(),
                    allowed: None
                },
                FilterField {
                    source_field: "Sales".to_string(),
                    allowed: None
                }
            ]
        );
    }
}
