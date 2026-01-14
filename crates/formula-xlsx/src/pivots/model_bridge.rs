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
    fn cache_field_ref(cache_def: &PivotCacheDefinition, name: String) -> PivotFieldRef {
        // Worksheet-backed caches should treat cache field names as literal header text. Parsing
        // DAX-like strings (e.g. `Table[Column]`) is reserved for Data Model / external pivots.
        match &cache_def.cache_source_type {
            super::PivotCacheSourceType::Worksheet => PivotFieldRef::CacheFieldName(name),
            _ => name.into(),
        }
    }

    table
        .data_fields
        .iter()
        .filter_map(|df| {
            let field_idx = df.fld? as usize;
            let source_field_name = cache_def.cache_fields.get(field_idx)?.name.clone();
            let source_field = cache_field_ref(cache_def, source_field_name.clone());
            let aggregation = map_subtotal(df.subtotal.as_deref());
            let default_name = {
                // For Data Model measures, prefer a human-friendly caption without DAX brackets
                // (Excel displays the measure as `Total Sales`, not `[Total Sales]`, in the default
                // "Sum of ..." label).
                let label = match &source_field {
                    PivotFieldRef::DataModelMeasure(measure) => measure.clone(),
                    _ => source_field.canonical_name().into_owned(),
                };
                format!("{} of {}", aggregation_display_name(aggregation), label)
            };
            let name = df
                .name
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or(default_name);

            let number_format = df
                .num_fmt_id
                .and_then(|id| resolve_pivot_num_fmt_id(id, styles));

            Some(ValueField {
                source_field,
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pivots::{PivotCacheField, PivotCacheSourceType};
    use pretty_assertions::assert_eq;

    #[test]
    fn worksheet_pivots_preserve_cache_field_names_that_look_like_dax_in_model_bridge() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataFields count="1">
    <dataField fld="0"/>
  </dataFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");
        let cache_def = PivotCacheDefinition {
            cache_source_type: PivotCacheSourceType::Worksheet,
            cache_fields: vec![PivotCacheField {
                name: "Table[Column]".to_string(),
                ..Default::default()
            }],
            ..Default::default()
        };

        let mut workbook = formula_model::Workbook::new();
        let styles =
            StylesPart::parse_or_default(None, &mut workbook.styles).expect("styles default");

        let value_fields = pivot_table_to_model_value_fields(&table, &cache_def, &styles);
        assert_eq!(value_fields.len(), 1);
        assert_eq!(
            value_fields[0].source_field.as_cache_field_name(),
            Some("Table[Column]")
        );
        assert_eq!(value_fields[0].name, "Sum of Table[Column]");
    }

    #[test]
    fn data_model_pivots_default_value_field_name_uses_measure_without_brackets_in_model_bridge() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataFields count="1">
    <dataField fld="0"/>
  </dataFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");
        let cache_def = PivotCacheDefinition {
            cache_source_type: PivotCacheSourceType::External,
            cache_fields: vec![PivotCacheField {
                name: "[Total Sales]".to_string(),
                ..Default::default()
            }],
            ..Default::default()
        };

        let mut workbook = formula_model::Workbook::new();
        let styles =
            StylesPart::parse_or_default(None, &mut workbook.styles).expect("styles default");

        let value_fields = pivot_table_to_model_value_fields(&table, &cache_def, &styles);
        assert_eq!(value_fields.len(), 1);
        assert_eq!(
            value_fields[0].source_field,
            PivotFieldRef::DataModelMeasure("Total Sales".to_string())
        );
        assert_eq!(value_fields[0].name, "Sum of Total Sales");
    }
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
