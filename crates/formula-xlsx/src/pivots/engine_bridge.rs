//! Conversions from parsed XLSX pivot parts into `formula_engine::pivot` types.
//!
//! The `formula-xlsx` crate's core responsibility is high-fidelity import/export.
//! For in-app pivot computation we also need to turn pivot cache/table metadata
//! into the engine's self-contained pivot types.

use chrono::NaiveDate;
use formula_engine::pivot::{
    AggregationType, CalculatedField, FilterField, GrandTotals, Layout, PivotCache, PivotConfig,
    PivotField, PivotFieldRef, PivotKeyPart, PivotValue, ShowAsType, SortOrder, SubtotalPosition,
    ValueField,
};
use formula_model::pivots::ScalarValue;
use std::collections::HashSet;

use crate::styles::StylesPart;

use super::cache_records::pivot_cache_datetime_to_naive_date;
use super::{
    PivotCacheDefinition, PivotCacheSourceType, PivotCacheValue, PivotTableDefinition,
    PivotTableFieldItem,
};
use crate::pivots::slicers::{PivotSlicerParts, SlicerSelectionState, TimelineSelectionState};

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

    let field_count = def.cache_fields.len();
    for record in records {
        let mut row = Vec::with_capacity(field_count);
        // Pivot caches can encode record values via a per-field "shared items" table (written as
        // `<x v="..."/>` indices in `pivotCacheRecords*.xml`). Resolve those indices using the
        // field position in the record (not the field name).
        row.extend(
            def.resolve_record_values(record)
                .take(field_count)
                .map(pivot_cache_value_to_engine_inner),
        );
        if row.len() < field_count {
            row.resize(field_count, PivotValue::Blank);
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
    // `<x v="..."/>` indices in `pivotCacheRecords*.xml`). Resolve those indices using the field
    // position in the record (not the field name).
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
        PivotCacheValue::DateTime(s) => pivot_cache_datetime_to_naive_date(&s)
            .map(PivotValue::Date)
            .unwrap_or_else(|| {
                if s.is_empty() {
                    PivotValue::Blank
                } else {
                    PivotValue::Text(s)
                }
            }),
        PivotCacheValue::Index(_) => PivotValue::Blank,
    }
}

fn pivot_key_display_string(value: PivotValue) -> String {
    value.to_key_part().display_string()
}

fn scalar_value_to_engine_key_part(value: &ScalarValue) -> PivotKeyPart {
    match value {
        ScalarValue::Blank => PivotKeyPart::Blank,
        ScalarValue::Text(s) => PivotKeyPart::Text(s.clone()),
        ScalarValue::Number(n) => PivotKeyPart::Number(PivotValue::canonical_number_bits(n.0)),
        ScalarValue::Date(d) => PivotKeyPart::Date(*d),
        ScalarValue::Bool(b) => PivotKeyPart::Bool(*b),
    }
}

fn slicer_cache_field_idx_best_effort(
    slicer: &crate::pivots::slicers::SlicerDefinition,
    cache_def: &PivotCacheDefinition,
    cache: &PivotCache,
) -> Option<usize> {
    // Prefer explicit field metadata when present.
    if let Some(field_name) = slicer.field_name.as_deref() {
        if let Some(idx) = cache_def.cache_fields.iter().position(|f| f.name == field_name) {
            return Some(idx);
        }

        // Some producers differ in case; try a case-insensitive match.
        let folded = field_name.trim().to_ascii_lowercase();
        if !folded.is_empty() {
            if let Some(idx) = cache_def
                .cache_fields
                .iter()
                .position(|f| f.name.to_ascii_lowercase() == folded)
            {
                return Some(idx);
            }
        }
    }

    if let Some(source_name) = slicer.source_name.as_deref() {
        if let Some(idx) = cache_def
            .cache_fields
            .iter()
            .position(|f| f.name == source_name)
        {
            return Some(idx);
        }

        // Some producers differ in case; try a case-insensitive match.
        let folded = source_name.trim().to_ascii_lowercase();
        if !folded.is_empty() {
            if let Some(idx) = cache_def
                .cache_fields
                .iter()
                .position(|f| f.name.to_ascii_lowercase() == folded)
            {
                return Some(idx);
            }
        }
    }

    // Fall back to inferring the field by matching slicer item keys against the cache's
    // per-field unique values.
    //
    // This is intentionally best-effort: OOXML slicer cache parts often omit a stable
    // field mapping (especially in simplified fixtures), but the item values typically
    // correspond to one cache field. When multiple fields share the same items (or the
    // slicer uses index keys), we may not be able to infer the correct field.
    let item_iter: Box<dyn Iterator<Item = &String> + '_> =
        if !slicer.selection.available_items.is_empty() {
            Box::new(slicer.selection.available_items.iter())
        } else if let Some(items) = &slicer.selection.selected_items {
            Box::new(items.iter())
        } else {
            return None;
        };

    let items = item_iter
        .map(|s| s.trim().to_ascii_lowercase())
        .filter(|s| !s.is_empty())
        .collect::<Vec<_>>();
    if items.is_empty() {
        return None;
    }

    let cache_name_folded = slicer
        .cache_name
        .as_deref()
        .unwrap_or_default()
        .to_ascii_lowercase();
    let slicer_name_folded = slicer
        .name
        .as_deref()
        .unwrap_or_default()
        .to_ascii_lowercase();

    let mut best: Option<(usize, usize, bool)> = None; // (field_idx, match_count, name_hint)
    for (field_idx, field) in cache_def.cache_fields.iter().enumerate() {
        let Some(values) = cache.unique_values.get(&field.name) else {
            continue;
        };

        let mut value_strings = HashSet::with_capacity(values.len());
        for value in values {
            value_strings.insert(value.to_key_part().display_string().to_ascii_lowercase());
        }

        let mut match_count = 0usize;
        for item in &items {
            if value_strings.contains(item) {
                match_count += 1;
            }
        }
        if match_count == 0 {
            continue;
        }

        let field_folded = field.name.to_ascii_lowercase();
        let name_hint = !field_folded.is_empty()
            && (cache_name_folded.contains(&field_folded)
                || slicer_name_folded.contains(&field_folded));

        let is_better = match best {
            None => true,
            Some((_best_idx, best_matches, best_hint)) => {
                match_count > best_matches
                    || (match_count == best_matches && name_hint && !best_hint)
            }
        };
        if is_better {
            best = Some((field_idx, match_count, name_hint));
        }
    }

    best.map(|(idx, _, _)| idx)
}

fn timeline_cache_field_name_best_effort(
    timeline: &crate::pivots::slicers::TimelineDefinition,
    cache_def: &PivotCacheDefinition,
    cache: &PivotCache,
) -> Option<String> {
    if let Some(idx) = timeline.base_field {
        if let Some(field) = cache_def.cache_fields.get(idx as usize) {
            return Some(field.name.clone());
        }
    }

    if let Some(field_name) = timeline.field_name.as_deref() {
        if cache.unique_values.contains_key(field_name) {
            return Some(field_name.to_string());
        }
        let folded = field_name.trim().to_ascii_lowercase();
        if !folded.is_empty() {
            if let Some(name) = cache
                .unique_values
                .keys()
                .find(|k| k.to_ascii_lowercase() == folded)
            {
                return Some(name.clone());
            }
        }
    }

    if let Some(source_name) = timeline.source_name.as_deref() {
        if cache.unique_values.contains_key(source_name) {
            return Some(source_name.to_string());
        }
        let folded = source_name.trim().to_ascii_lowercase();
        if !folded.is_empty() {
            if let Some(name) = cache
                .unique_values
                .keys()
                .find(|k| k.to_ascii_lowercase() == folded)
            {
                return Some(name.clone());
            }
        }
    }

    let start = timeline.selection.start.as_deref().and_then(parse_iso_ymd);
    let end = timeline.selection.end.as_deref().and_then(parse_iso_ymd);
    if start.is_none() && end.is_none() {
        return None;
    }

    let cache_name_folded = timeline
        .cache_name
        .as_deref()
        .unwrap_or_default()
        .to_ascii_lowercase();
    let timeline_name_folded = timeline
        .name
        .as_deref()
        .unwrap_or_default()
        .to_ascii_lowercase();

    let mut best: Option<(String, usize, bool, usize)> = None;
    for field in &cache_def.cache_fields {
        let Some(values) = cache.unique_values.get(&field.name) else {
            continue;
        };

        let mut date_values = Vec::new();
        for value in values {
            if let PivotValue::Date(d) = value {
                date_values.push(*d);
            }
        }
        if date_values.is_empty() {
            continue;
        }

        let mut in_range = 0usize;
        for date in &date_values {
            if start.is_some_and(|s| *date < s) {
                continue;
            }
            if end.is_some_and(|e| *date > e) {
                continue;
            }
            in_range += 1;
        }

        let field_folded = field.name.to_ascii_lowercase();
        let name_hint = !field_folded.is_empty()
            && (cache_name_folded.contains(&field_folded)
                || timeline_name_folded.contains(&field_folded));

        let score = in_range;
        let is_better = match &best {
            None => true,
            Some((_name, best_score, best_hint, best_date_count)) => {
                score > *best_score
                    || (score == *best_score && name_hint && !*best_hint)
                    || (score == *best_score && name_hint == *best_hint && date_values.len() > *best_date_count)
            }
        };
        if is_better {
            best = Some((field.name.clone(), score, name_hint, date_values.len()));
        }
    }

    best.map(|(name, _score, _hint, _date_count)| name)
}

/// Convert a parsed slicer selection into a pivot-engine filter field.
///
/// Callers can supply a resolver that maps slicer item keys (often stored as `x` indices) back into
/// typed values.
pub fn slicer_selection_to_engine_filter_with_resolver<F>(
    field: impl Into<String>,
    selection: &SlicerSelectionState,
    mut resolve: F,
) -> FilterField
where
    F: FnMut(&str) -> Option<ScalarValue>,
{
    let source_field = PivotFieldRef::CacheFieldName(field.into());
    let allowed = match &selection.selected_items {
        None => None,
        Some(items) => {
            let mut allowed = HashSet::with_capacity(items.len());
            for item in items {
                let scalar = resolve(item).unwrap_or_else(|| ScalarValue::from(item.as_str()));
                allowed.insert(scalar_value_to_engine_key_part(&scalar));
            }
            Some(allowed)
        }
    };

    FilterField { source_field, allowed }
}

/// Convert a parsed timeline selection into a pivot-engine filter field.
///
/// The pivot engine currently supports only "allowed-set" filters. We implement timeline date
/// ranges by scanning the pivot cache's unique values for the field and selecting the ones that
/// fall within the inclusive `[start, end]` range.
pub fn timeline_selection_to_engine_filter(
    field: impl Into<String>,
    selection: &TimelineSelectionState,
    cache: &PivotCache,
) -> Option<FilterField> {
    let field_name = field.into();
    let start = selection.start.as_deref().and_then(parse_iso_ymd);
    let end = selection.end.as_deref().and_then(parse_iso_ymd);

    if start.is_none() && end.is_none() {
        return None;
    }

    let values = cache.unique_values.get(&field_name)?;
    let mut allowed = HashSet::new();
    for value in values {
        let PivotValue::Date(date) = value else {
            continue;
        };
        if start.is_some_and(|s| *date < s) {
            continue;
        }
        if end.is_some_and(|e| *date > e) {
            continue;
        }
        allowed.insert(PivotKeyPart::Date(*date));
    }

    Some(FilterField {
        source_field: PivotFieldRef::CacheFieldName(field_name),
        allowed: Some(allowed),
    })
}

/// Compute pivot-engine filter fields for slicers/timelines connected to a given pivot table.
pub fn pivot_slicer_parts_to_engine_filters(
    pivot_table_part: &str,
    cache_def: &PivotCacheDefinition,
    cache: &PivotCache,
    parts: &PivotSlicerParts,
) -> Vec<FilterField> {
    let mut out = Vec::new();

    for slicer in &parts.slicers {
        if !slicer
            .connected_pivot_tables
            .iter()
            .any(|p| p == pivot_table_part)
        {
            continue;
        }
        if slicer.selection.selected_items.is_none() {
            continue;
        }
        let Some(field_idx) = slicer_cache_field_idx_best_effort(slicer, cache_def, cache) else {
            continue;
        };
        let Some(field) = cache_def
            .cache_fields
            .get(field_idx)
            .map(|f| f.name.clone())
        else {
            continue;
        };

        let filter =
            slicer_selection_to_engine_filter_with_resolver(field, &slicer.selection, |key| {
                let idx = key.trim().parse::<u32>().ok()?;
                cache_def.resolve_shared_item(field_idx, idx)
            });
        out.push(filter);
    }

    for timeline in &parts.timelines {
        if !timeline
            .connected_pivot_tables
            .iter()
            .any(|p| p == pivot_table_part)
        {
            continue;
        }

        let field = timeline_cache_field_name_best_effort(timeline, cache_def, cache);
        let Some(field) = field else {
            continue;
        };

        if let Some(filter) = timeline_selection_to_engine_filter(field, &timeline.selection, cache)
        {
            out.push(filter);
        }
    }

    out
}

/// Extend an existing pivot-engine config with slicer/timeline filters for a given pivot table.
pub fn apply_pivot_slicer_parts_to_engine_config(
    cfg: &mut PivotConfig,
    pivot_table_part: &str,
    cache_def: &PivotCacheDefinition,
    cache: &PivotCache,
    parts: &PivotSlicerParts,
) {
    cfg.filter_fields
        .extend(pivot_slicer_parts_to_engine_filters(
            pivot_table_part,
            cache_def,
            cache,
            parts,
        ));
}

fn parse_iso_ymd(value: &str) -> Option<NaiveDate> {
    let trimmed = value.trim();
    let ymd = trimmed.get(..10).unwrap_or(trimmed);
    NaiveDate::parse_from_str(ymd, "%Y-%m-%d").ok()
}
/// Convert a parsed pivot table definition into a pivot-engine config.
///
/// This is a best-effort conversion; unsupported layout / display options are
/// ignored.
pub fn pivot_table_to_engine_config(
    table: &PivotTableDefinition,
    cache_def: &PivotCacheDefinition,
) -> PivotConfig {
    pivot_table_to_engine_config_with_styles(table, cache_def, None)
}

/// Convert a parsed pivot table definition into a [`PivotConfig`], optionally resolving number
/// format ids (`numFmtId`) using a parsed `styles.xml` [`StylesPart`].
///
/// When `styles` is `None`, this uses built-in format mappings when available and otherwise
/// preserves the id using a `__builtin_numFmtId:*` placeholder string.
pub fn pivot_table_to_engine_config_with_styles(
    table: &PivotTableDefinition,
    cache_def: &PivotCacheDefinition,
    styles: Option<&StylesPart>,
) -> PivotConfig {
    fn cache_field_ref(cache_def: &PivotCacheDefinition, name: String) -> PivotFieldRef {
        // Worksheet-backed caches should treat cache field names as literal header text. Parsing
        // DAX-like strings (e.g. `Table[Column]`) is reserved for Data Model / external pivots.
        match &cache_def.cache_source_type {
            PivotCacheSourceType::Worksheet => PivotFieldRef::CacheFieldName(name),
            _ => name.into(),
        }
    }

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
            let source_field_name = cache_def.cache_fields.get(field_idx)?.name.clone();
            let aggregation = map_subtotal(df.subtotal.as_deref());
            let default_name = format!(
                "{} of {}",
                aggregation_display_name(aggregation),
                &source_field_name
            );
            let name = df
                .name
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or(default_name);

            let show_as = map_show_data_as(df.show_data_as.as_deref());

            // `dataField@baseField` is an index into the cache fields list.
            let base_field = df.base_field.and_then(|base_field_idx| {
                cache_def
                    .cache_fields
                    .get(base_field_idx as usize)
                    .map(|f| cache_field_ref(cache_def, f.name.clone()))
            });

            // `dataField@baseItem` refers to an item within `baseField`'s shared-items table.
            let base_item = df
                .base_field
                .zip(df.base_item)
                .and_then(|(field_idx, item_idx)| {
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
                source_field: source_field_name.into(),
                name,
                aggregation,
                number_format: df.num_fmt_id.and_then(|id| resolve_pivot_num_fmt_id(id, styles)),
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
            let source_field = cache_field_ref(cache_def, cache_field.name.clone());

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
                let key_part = pivot_value.to_key_part();
                let mut set = HashSet::new();
                set.insert(key_part);
                Some(set)
            });

            Some(FilterField { source_field, allowed })
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
    let calculated_fields = cache_def
        .calculated_fields()
        .into_iter()
        .map(|cf| CalculatedField {
            name: cf.name,
            formula: cf.formula,
        })
        .collect();

    PivotConfig {
        row_fields,
        column_fields,
        value_fields,
        filter_fields,
        calculated_fields,
        calculated_items: Vec::new(),
        layout,
        subtotals,
        grand_totals,
    }
}

fn resolve_pivot_num_fmt_id(num_fmt_id: u32, styles: Option<&StylesPart>) -> Option<String> {
    if num_fmt_id == 0 {
        return None;
    }

    if num_fmt_id <= u16::MAX as u32 {
        let id_u16 = num_fmt_id as u16;
        if let Some(styles) = styles {
            if let Some(code) = styles.num_fmt_code_for_id(id_u16) {
                return Some(code.to_string());
            }
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

fn pivot_table_field_to_engine(
    table: &PivotTableDefinition,
    cache_def: &PivotCacheDefinition,
    field_idx: u32,
) -> Option<PivotField> {
    let cache_field = cache_def.cache_fields.get(field_idx as usize)?;

    let mut field = PivotField::new(match &cache_def.cache_source_type {
        PivotCacheSourceType::Worksheet => PivotFieldRef::CacheFieldName(cache_field.name.clone()),
        _ => cache_field.name.clone().into(),
    });
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
            field.manual_sort = table_field.manual_sort_items.as_ref().and_then(|items| {
                let mut out: Vec<PivotKeyPart> = items
                    .iter()
                    .filter_map(|item| match item {
                        PivotTableFieldItem::Name(name) => Some(PivotKeyPart::Text(name.clone())),
                        PivotTableFieldItem::Index(item_idx) => cache_def
                            .cache_fields
                            .get(field_idx as usize)
                            .and_then(|f| f.shared_items.as_ref())
                            .and_then(|items| items.get(*item_idx as usize))
                            .cloned()
                            .map(|v| {
                                pivot_cache_value_to_engine(cache_def, field_idx as usize, v)
                                    .to_key_part()
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

    use crate::pivots::slicers::SlicerDefinition;
    use crate::pivots::slicers::TimelineDefinition;
    use crate::pivots::PivotCacheField;
    use pretty_assertions::assert_eq;
    use std::collections::HashSet;

    fn cache_field(name: &str) -> PivotFieldRef {
        PivotFieldRef::CacheFieldName(name.to_string())
    }

    #[test]
    fn map_show_data_as_handles_known_strings_case_insensitively() {
        let cases = [
            ("normal", Some(ShowAsType::Normal)),
            ("Normal", Some(ShowAsType::Normal)),
            ("percentOfGrandTotal", Some(ShowAsType::PercentOfGrandTotal)),
            ("PERCENTOFGRANDTOTAL", Some(ShowAsType::PercentOfGrandTotal)),
            ("percentOfRowTotal", Some(ShowAsType::PercentOfRowTotal)),
            (
                "percentOfColumnTotal",
                Some(ShowAsType::PercentOfColumnTotal),
            ),
            ("percentOf", Some(ShowAsType::PercentOf)),
            (
                "percentDifferenceFrom",
                Some(ShowAsType::PercentDifferenceFrom),
            ),
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
        assert_eq!(cfg.value_fields[0].base_field, Some(cache_field("Region")));
    }

    #[test]
    fn worksheet_pivots_preserve_cache_field_names_that_look_like_dax() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pivotFields count="1">
    <pivotField/>
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
            cache_source_type: PivotCacheSourceType::Worksheet,
            cache_fields: vec![PivotCacheField {
                name: "Table[Column]".to_string(),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(cfg.row_fields.len(), 1);
        assert_eq!(
            cfg.row_fields[0].source_field.as_cache_field_name(),
            Some("Table[Column]")
        );
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
        assert_eq!(cfg.value_fields[0].base_field, Some(cache_field("Region")));
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
                    source_field: cache_field("Region"),
                    allowed: None
                },
                FilterField {
                    source_field: cache_field("Sales"),
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
                source_field: cache_field("Region"),
                allowed: Some(allowed),
            }]
        );
    }

    #[test]
    fn pivot_table_to_engine_config_treats_page_field_negative_item_as_all() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pageFields count="1">
    <pageField fld="0" item="-1"/>
  </pageFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");

        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Region".to_string(),
                shared_items: Some(vec![PivotCacheValue::String("East".to_string())]),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(
            cfg.filter_fields,
            vec![FilterField {
                source_field: cache_field("Region"),
                allowed: None,
            }]
        );
    }

    #[test]
    fn pivot_table_to_engine_config_page_field_out_of_range_item_falls_back_to_all() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pageFields count="1">
    <pageField fld="0" item="99"/>
  </pageFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");

        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Region".to_string(),
                shared_items: Some(vec![PivotCacheValue::String("East".to_string())]),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(
            cfg.filter_fields,
            vec![FilterField {
                source_field: cache_field("Region"),
                allowed: None,
            }]
        );
    }

    #[test]
    fn pivot_table_to_engine_config_canonicalizes_numeric_page_field_items() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pageFields count="1">
    <pageField fld="0" item="0"/>
  </pageFields>
</pivotTableDefinition>"#;

        let table =
            PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml).expect("parse");

        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "Metric".to_string(),
                shared_items: Some(vec![PivotCacheValue::Number(-0.0)]),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);

        let allowed = cfg.filter_fields[0]
            .allowed
            .as_ref()
            .expect("expected allowed set for selected page field item");

        assert!(allowed.contains(&PivotKeyPart::Number(0.0_f64.to_bits())));
        assert!(!allowed.contains(&PivotKeyPart::Number((-0.0_f64).to_bits())));
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
            Some(
                &[
                    PivotKeyPart::Text("B".to_string()),
                    PivotKeyPart::Text("A".to_string()),
                    PivotKeyPart::Text("C".to_string()),
                ][..]
            )
        );
    }

    #[test]
    fn maps_indexed_manual_sort_items_using_cache_shared_items() {
        let table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pivotFields count="1">
    <pivotField axis="axisRow" sortType="manual">
      <items count="3">
        <item x="2"/>
        <item x="0"/>
        <item x="1"/>
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
                shared_items: Some(vec![
                    PivotCacheValue::String("East".to_string()),
                    PivotCacheValue::String("West".to_string()),
                    PivotCacheValue::String("North".to_string()),
                ]),
                ..Default::default()
            }],
            ..Default::default()
        };

        let cfg = pivot_table_to_engine_config(&table, &cache_def);
        assert_eq!(cfg.row_fields.len(), 1);
        assert_eq!(cfg.row_fields[0].sort_order, SortOrder::Manual);
        assert_eq!(
            cfg.row_fields[0].manual_sort.as_deref(),
            Some(
                &[
                    PivotKeyPart::Text("North".to_string()),
                    PivotKeyPart::Text("East".to_string()),
                    PivotKeyPart::Text("West".to_string()),
                ][..]
            )
        );
    }

    #[test]
    fn pivot_slicer_parts_to_engine_filters_infers_field_from_item_values_when_source_name_is_pivot_name(
    ) {
        let source = vec![
            vec![
                PivotValue::Text("Region".to_string()),
                PivotValue::Text("Product".to_string()),
                PivotValue::Text("Sales".to_string()),
            ],
            vec![
                PivotValue::Text("East".to_string()),
                PivotValue::Text("A".to_string()),
                PivotValue::Number(10.0),
            ],
            vec![
                PivotValue::Text("West".to_string()),
                PivotValue::Text("A".to_string()),
                PivotValue::Number(20.0),
            ],
            vec![
                PivotValue::Text("East".to_string()),
                PivotValue::Text("B".to_string()),
                PivotValue::Number(30.0),
            ],
        ];
        let cache = PivotCache::from_range(&source).expect("build pivot cache");
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

        let mut selected_items = HashSet::new();
        selected_items.insert("East".to_string());
        let slicer = SlicerDefinition {
            part_name: "xl/slicers/slicer1.xml".to_string(),
            name: Some("RegionSlicer".to_string()),
            uid: None,
            cache_part: None,
            cache_name: Some("RegionSlicerCache".to_string()),
            // Some producers persist the pivot table name in `sourceName` instead of the field name.
            source_name: Some("PivotTable1".to_string()),
            field_name: None,
            connected_pivot_tables: vec!["xl/pivotTables/pivotTable1.xml".to_string()],
            connected_tables: vec![],
            placed_on_drawings: vec![],
            placed_on_sheets: vec![],
            placed_on_sheet_names: vec![],
            selection: SlicerSelectionState {
                available_items: vec!["East".to_string(), "West".to_string()],
                selected_items: Some(selected_items),
            },
        };

        let parts = PivotSlicerParts {
            slicers: vec![slicer],
            timelines: vec![],
        };

        let filters = pivot_slicer_parts_to_engine_filters(
            "xl/pivotTables/pivotTable1.xml",
            &cache_def,
            &cache,
            &parts,
        );
        assert_eq!(filters.len(), 1);
        assert_eq!(filters[0].source_field, cache_field("Region"));
        assert_eq!(
            filters[0].allowed.as_ref(),
            Some(&HashSet::from([PivotKeyPart::Text("East".to_string())]))
        );
    }

    #[test]
    fn maps_indexed_manual_sort_items_into_engine_field_using_cache_shared_items() {
        let table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pivotFields count="1">
    <pivotField axis="axisRow" sortType="manual">
      <items count="3">
        <item x="2"/>
        <item x="0"/>
        <item x="1"/>
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
                shared_items: Some(vec![
                    PivotCacheValue::String("A".to_string()),
                    PivotCacheValue::String("B".to_string()),
                    PivotCacheValue::String("C".to_string()),
                ]),
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
                PivotKeyPart::Text("C".to_string()),
                PivotKeyPart::Text("A".to_string()),
                PivotKeyPart::Text("B".to_string()),
            ][..])
        );
    }

    #[test]
    fn pivot_slicer_parts_to_engine_filters_infers_timeline_field_when_base_field_missing() {
        let source = vec![
            vec![
                PivotValue::Text("Date".to_string()),
                PivotValue::Text("Sales".to_string()),
            ],
            vec![
                PivotValue::Date(NaiveDate::from_ymd_opt(2024, 1, 1).unwrap()),
                1.into(),
            ],
            vec![
                PivotValue::Date(NaiveDate::from_ymd_opt(2024, 1, 15).unwrap()),
                2.into(),
            ],
            vec![
                PivotValue::Date(NaiveDate::from_ymd_opt(2024, 2, 1).unwrap()),
                3.into(),
            ],
        ];
        let cache = PivotCache::from_range(&source).expect("build pivot cache");

        let cache_def = PivotCacheDefinition {
            cache_fields: vec![
                PivotCacheField {
                    name: "Date".to_string(),
                    ..Default::default()
                },
                PivotCacheField {
                    name: "Sales".to_string(),
                    ..Default::default()
                },
            ],
            ..Default::default()
        };

        let timeline = TimelineDefinition {
            part_name: "xl/timelines/timeline1.xml".to_string(),
            name: Some("DateTimeline".to_string()),
            uid: None,
            cache_part: None,
            cache_name: Some("DateTimelineCache".to_string()),
            source_name: None,
            field_name: None,
            base_field: None,
            level: None,
            connected_pivot_tables: vec!["xl/pivotTables/pivotTable1.xml".to_string()],
            placed_on_drawings: vec![],
            placed_on_sheets: vec![],
            placed_on_sheet_names: vec![],
            selection: TimelineSelectionState {
                start: Some("2024-01-01".to_string()),
                end: Some("2024-01-31".to_string()),
            },
        };

        let parts = PivotSlicerParts {
            slicers: vec![],
            timelines: vec![timeline],
        };

        let filters = pivot_slicer_parts_to_engine_filters(
            "xl/pivotTables/pivotTable1.xml",
            &cache_def,
            &cache,
            &parts,
        );
        assert_eq!(filters.len(), 1);
        assert_eq!(
            filters[0].source_field,
            PivotFieldRef::CacheFieldName("Date".to_string())
        );
        assert_eq!(
            filters[0].allowed.as_ref(),
            Some(&HashSet::from([
                PivotKeyPart::Date(NaiveDate::from_ymd_opt(2024, 1, 1).unwrap()),
                PivotKeyPart::Date(NaiveDate::from_ymd_opt(2024, 1, 15).unwrap()),
            ]))
        );
    }
}
