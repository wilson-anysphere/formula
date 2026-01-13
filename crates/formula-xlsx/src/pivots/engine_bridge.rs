//! Conversions from parsed XLSX pivot parts into `formula_engine::pivot` types.
//!
//! The `formula-xlsx` crate's core responsibility is high-fidelity import/export.
//! For in-app pivot computation we also need to turn pivot cache/table metadata
//! into the engine's self-contained pivot types.

use formula_engine::pivot::{
    AggregationType, FilterField, GrandTotals, Layout, PivotConfig, PivotField, PivotValue,
    ShowAsType, SubtotalPosition, ValueField,
};
use formula_engine::pivot::{PivotKeyPart, SortOrder};
use std::collections::HashSet;

use super::cache_records::pivot_cache_datetime_to_naive_date;
use super::{PivotCacheDefinition, PivotCacheValue, PivotTableDefinition, PivotTableFieldItem};

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
            row.push(pivot_cache_value_to_engine(def, field_idx, value));
        }
        out.push(row);
    }

    out
}

fn pivot_cache_value_to_engine(
    def: &PivotCacheDefinition,
    field_idx: usize,
    value: PivotCacheValue,
) -> PivotValue {
    // Pivot caches can encode record values via a per-field "shared items" table (written as
    // `<x v="..."/>` indices in `pivotCacheRecords*.xml`). Resolve those indices using the *field
    // position* in the record (not the field name).
    let value = def.resolve_record_value(field_idx, value);
    pivot_cache_value_to_engine_inner(value)
}

fn pivot_cache_value_to_engine_inner(value: PivotCacheValue) -> PivotValue {
    match value {
        PivotCacheValue::String(s) => PivotValue::Text(s),
        PivotCacheValue::Number(n) => PivotValue::Number(n),
        PivotCacheValue::Bool(b) => PivotValue::Bool(b),
        PivotCacheValue::Missing => PivotValue::Blank,
        PivotCacheValue::Error(_) => PivotValue::Blank,
        PivotCacheValue::DateTime(s) => {
            if s.is_empty() {
                PivotValue::Blank
            } else if let Some(date) = pivot_cache_datetime_to_naive_date(&s) {
                PivotValue::Date(date)
            } else {
                PivotValue::Text(s)
            }
        }
        PivotCacheValue::Index(_) => PivotValue::Blank,
    }
}

fn pivot_value_to_key_part(value: PivotValue) -> PivotKeyPart {
    match value {
        PivotValue::Blank => PivotKeyPart::Blank,
        PivotValue::Number(n) => {
            // Match the pivot engine's internal canonicalization:
            // - treat -0.0 and 0.0 as identical
            // - canonicalize all NaNs to a single representation
            let bits = if n == 0.0 {
                0.0_f64.to_bits()
            } else if n.is_nan() {
                f64::NAN.to_bits()
            } else {
                n.to_bits()
            };
            PivotKeyPart::Number(bits)
        }
        PivotValue::Date(d) => PivotKeyPart::Date(d),
        PivotValue::Text(s) => PivotKeyPart::Text(s),
        PivotValue::Bool(b) => PivotKeyPart::Bool(b),
    }
}

fn pivot_key_display_string(value: PivotValue) -> String {
    match value {
        PivotValue::Blank => "(blank)".to_string(),
        PivotValue::Number(n) => {
            // Keep it simple; Excel has more nuanced formatting.
            if n.fract() == 0.0 {
                format!("{}", n as i64)
            } else {
                format!("{n}")
            }
        }
        PivotValue::Date(d) => d.to_string(),
        PivotValue::Text(s) => s,
        PivotValue::Bool(b) => b.to_string(),
    }
}

fn pivot_value_to_key_part(value: PivotValue) -> PivotKeyPart {
    fn canonical_number_bits(n: f64) -> u64 {
        if n == 0.0 {
            // Treat -0.0 and 0.0 as the same pivot item (Excel displays them identically).
            return 0.0_f64.to_bits();
        }
        if n.is_nan() {
            // Canonicalize all NaNs to a single representation so we don't emit multiple "NaN" rows.
            return f64::NAN.to_bits();
        }
        n.to_bits()
    }

    match value {
        PivotValue::Blank => PivotKeyPart::Blank,
        PivotValue::Number(n) => PivotKeyPart::Number(canonical_number_bits(n)),
        PivotValue::Date(d) => PivotKeyPart::Date(d),
        PivotValue::Text(s) => PivotKeyPart::Text(s),
        PivotValue::Bool(b) => PivotKeyPart::Bool(b),
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
        .filter_map(|idx| pivot_table_field_to_engine(table, cache_def, *idx))
        .collect::<Vec<_>>();

    let column_fields = table
        .col_fields
        .iter()
        .filter_map(|idx| pivot_table_field_to_engine(table, cache_def, *idx))
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
            let base_item = df.base_field.zip(df.base_item).and_then(|(field_idx, item_idx)| {
                let shared_items = cache_def
                    .cache_fields
                    .get(field_idx as usize)?
                    .shared_items
                    .as_ref()?;
                let item = shared_items.get(item_idx as usize)?.clone();
                Some(pivot_key_display_string(pivot_cache_value_to_engine(
                    cache_def,
                    field_idx as usize,
                    item,
                )))
            });
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
        .page_field_entries
        .iter()
        .filter_map(|page_field| {
            let field_idx = page_field.fld as usize;
            let cache_field = cache_def.cache_fields.get(field_idx)?;
            let source_field = cache_field.name.clone();

            // `pageField@item` is typically a shared-item index for the field, with `-1` meaning
            // "(All)". We currently model report filters as a single-selection allowed-set.
            let allowed = page_field.item.and_then(|item| {
                if item < 0 {
                    return None;
                }
                let item_idx = usize::try_from(item).ok()?;
                let shared_items = cache_field.shared_items.as_ref()?;
                let item = shared_items.get(item_idx)?.clone();
                let pivot_value = pivot_cache_value_to_engine(cache_def, field_idx, item);
                let key_part = pivot_value_to_key_part(pivot_value);
                let mut set = HashSet::new();
                set.insert(key_part);
                Some(set)
            });

            Some(FilterField {
                source_field,
                allowed,
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

fn pivot_table_field_to_engine(
    table: &PivotTableDefinition,
    cache_def: &PivotCacheDefinition,
    field_idx: u32,
) -> Option<PivotField> {
    let cache_field = cache_def.cache_fields.get(field_idx as usize)?;

    let mut field = PivotField::new(cache_field.name.clone());
    let table_field = table.pivot_fields.get(field_idx as usize);

    if let Some(table_field) = table_field {
        if let Some(sort_type) = table_field.sort_type.as_deref() {
            match sort_type.to_ascii_lowercase().as_str() {
                "descending" => field.sort_order = SortOrder::Descending,
                "manual" => field.sort_order = SortOrder::Manual,
                // Default: keep engine default (ascending).
                _ => {}
            }
        }

        if field.sort_order == SortOrder::Manual {
            field.manual_sort = table_field
                .manual_sort_items
                .as_ref()
                .and_then(|items| {
                    let mut out: Vec<PivotKeyPart> = items
                        .iter()
                        .filter_map(|item| match item {
                            PivotTableFieldItem::Name(name) => {
                                Some(PivotKeyPart::Text(name.clone()))
                            }
                            PivotTableFieldItem::Index(item_idx) => cache_def
                                .cache_fields
                                .get(field_idx as usize)
                                .and_then(|f| f.shared_items.as_ref())
                                .and_then(|items| items.get(*item_idx as usize))
                                .cloned()
                                .map(|v| {
                                    pivot_value_to_key_part(pivot_cache_value_to_engine(
                                        cache_def,
                                        field_idx as usize,
                                        v,
                                    ))
                                }),
                        })
                        .collect();
                    if out.is_empty() {
                        None
                    } else {
                        // De-dupe while preserving order (Excel seems to treat duplicates as no-ops).
                        let mut seen: HashSet<PivotKeyPart> = HashSet::new();
                        out.retain(|p| seen.insert(p.clone()));
                        Some(out)
                    }
                });
        }
    }

    Some(field)
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
    fn maps_base_item_from_shared_items_when_available() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataFields count="1">
    <dataField fld="0" showDataAs="percentOf" baseField="1" baseItem="0"/>
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
                    shared_items: Some(vec![
                        PivotCacheValue::String("East".to_string()),
                        PivotCacheValue::String("West".to_string()),
                    ]),
                    ..Default::default()
                },
            ],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(cfg.value_fields.len(), 1);
        assert_eq!(cfg.value_fields[0].base_field.as_deref(), Some("Region"));
        assert_eq!(cfg.value_fields[0].base_item.as_deref(), Some("East"));
    }

    #[test]
    fn pivot_cache_to_engine_source_resolves_shared_item_indices() {
        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Region".to_string(),
                shared_items: Some(vec![
                    PivotCacheValue::String("East".to_string()),
                    PivotCacheValue::String("West".to_string()),
                ]),
                ..Default::default()
            }],
            ..Default::default()
        };

        let source = pivot_cache_to_engine_source(
            &cache_def,
            vec![vec![PivotCacheValue::Index(1)]].into_iter(),
        );
        assert_eq!(source.len(), 2, "header + one record");
        assert_eq!(source[0], vec![PivotValue::Text("Region".to_string())]);
        assert_eq!(source[1], vec![PivotValue::Text("West".to_string())]);
    }

    #[test]
    fn pivot_table_to_engine_config_maps_manual_sort_items_via_shared_item_indices() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pivotFields count="1">
    <pivotField sortType="manual">
      <items count="2">
        <item x="1"/>
        <item x="0"/>
      </items>
    </pivotField>
  </pivotFields>
  <rowFields count="1">
    <field x="0"/>
  </rowFields>
  <dataFields count="1">
    <dataField fld="0"/>
  </dataFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");
        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Region".to_string(),
                shared_items: Some(vec![
                    PivotCacheValue::String("East".to_string()),
                    PivotCacheValue::String("West".to_string()),
                ]),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(cfg.row_fields.len(), 1);
        assert_eq!(cfg.row_fields[0].sort_order, SortOrder::Manual);
        assert_eq!(
            cfg.row_fields[0].manual_sort,
            Some(vec![
                PivotKeyPart::Text("West".to_string()),
                PivotKeyPart::Text("East".to_string()),
            ])
        );
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

    #[test]
    fn pivot_table_to_engine_config_respects_page_field_item_selection() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pageFields count="1">
    <pageField fld="0" item="1"/>
  </pageFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");

        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Region".to_string(),
                shared_items: Some(vec![
                    PivotCacheValue::String("East".to_string()),
                    PivotCacheValue::String("West".to_string()),
                ]),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);

        let mut allowed = HashSet::new();
        allowed.insert(PivotKeyPart::Text("West".to_string()));

        assert_eq!(
            cfg.filter_fields,
            vec![FilterField {
                source_field: "Region".to_string(),
                allowed: Some(allowed),
            }]
        );
    }

    #[test]
    fn maps_descending_sort_type_into_engine_field() {
        let table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pivotFields count="1">
    <pivotField axis="axisRow" sortType="descending"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
</pivotTableDefinition>"#;

        let table = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", table_xml)
            .expect("parse pivot table definition");

        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Region".to_string(),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(cfg.row_fields.len(), 1);
        assert_eq!(cfg.row_fields[0].sort_order, SortOrder::Descending);
    }

    #[test]
    fn maps_named_manual_sort_items_into_engine_field() {
        let table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pivotFields count="1">
    <pivotField axis="axisRow" sortType="manual">
      <items count="3">
        <item n="B"/>
        <item n="A"/>
        <item n="C"/>
      </items>
    </pivotField>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
</pivotTableDefinition>"#;

        let table = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", table_xml)
            .expect("parse pivot table definition");

        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Region".to_string(),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(cfg.row_fields.len(), 1);
        assert_eq!(cfg.row_fields[0].sort_order, SortOrder::Manual);
        assert_eq!(
            cfg.row_fields[0].manual_sort.as_deref(),
            Some(&[
                PivotKeyPart::Text("B".to_string()),
                PivotKeyPart::Text("A".to_string()),
                PivotKeyPart::Text("C".to_string()),
            ][..])
        );
    }
}
