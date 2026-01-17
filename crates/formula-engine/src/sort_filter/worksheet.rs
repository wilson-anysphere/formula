use crate::locale::ValueLocaleConfig;
use crate::sort_filter::sort::{
    compute_header_rows_with_value_locale, compute_row_permutation_with_value_locale,
};
use crate::sort_filter::{
    apply_autofilter_with_value_locale, AutoFilter, CellValue, FilterResult, RowPermutation,
    SortSpec,
};
use crate::{parse_formula, CellAddr, LocaleConfig, ParseOptions, SerializeOptions};
use formula_model::{
    CellRef, CellValue as ModelCellValue, Outline, Range, RowProperties, SheetAutoFilter, Worksheet,
};

pub fn sort_worksheet_range(
    sheet: &mut Worksheet,
    range: Range,
    spec: &SortSpec,
) -> RowPermutation {
    sort_worksheet_range_with_value_locale(sheet, range, spec, ValueLocaleConfig::en_us())
}

pub fn sort_worksheet_range_with_value_locale(
    sheet: &mut Worksheet,
    range: Range,
    spec: &SortSpec,
    value_locale: ValueLocaleConfig,
) -> RowPermutation {
    let row_count = range.height() as usize;
    if row_count <= 1 || spec.keys.is_empty() {
        return RowPermutation {
            new_to_old: (0..row_count).collect(),
            old_to_new: (0..row_count).collect(),
        };
    }

    let start_row = range.start.row;
    let start_col = range.start.col;

    let width = range.width() as usize;
    let cell_at = |sheet: &Worksheet, local_row: usize, local_col: usize| -> CellValue {
        if local_col >= width {
            return CellValue::Blank;
        }
        let row = start_row + local_row as u32;
        let col = start_col + local_col as u32;
        model_cell_value_to_sort_value(&sheet.value(CellRef::new(row, col)))
    };

    let header_rows = compute_header_rows_with_value_locale(
        row_count,
        spec.header,
        &spec.keys,
        value_locale,
        |r, c| cell_at(sheet, r, c),
    );
    let perm = compute_row_permutation_with_value_locale(
        row_count,
        header_rows,
        &spec.keys,
        value_locale,
        |r, c| cell_at(sheet, r, c),
    );

    // Nothing to permute (e.g. header-only range).
    if header_rows >= row_count {
        return perm;
    }

    let data_start_row = start_row + header_rows as u32;
    if data_start_row > range.end.row {
        return perm;
    }

    let data_range = Range::new(
        CellRef::new(data_start_row, range.start.col),
        CellRef::new(range.end.row, range.end.col),
    );

    let it = sheet.iter_cells_in_range(data_range);
    let mut moved_cells = Vec::with_capacity(it.size_hint().0);
    for (cell_ref, cell) in it {
        moved_cells.push((cell_ref, cell.clone()));
    }

    sheet.clear_range(data_range);

    for (cell_ref, mut cell) in moved_cells {
        let local_old_row = (cell_ref.row - start_row) as usize;
        let local_new_row = perm.old_to_new[local_old_row];
        let new_row = start_row + local_new_row as u32;

        if let Some(formula) = cell.formula.as_deref() {
            if let Some(rewritten) = rewrite_formula_for_move(
                formula,
                CellAddr::new(cell_ref.row, cell_ref.col),
                CellAddr::new(new_row, cell_ref.col),
            ) {
                cell.formula = Some(rewritten);
            }
        }

        sheet.set_cell(CellRef::new(new_row, cell_ref.col), cell);
    }

    permute_row_properties(sheet, start_row, header_rows, range.end.row, &perm);

    perm
}

pub fn apply_autofilter_to_outline(
    sheet: &Worksheet,
    outline: &mut Outline,
    range: Range,
    filter: Option<&SheetAutoFilter>,
) -> FilterResult {
    apply_autofilter_to_outline_with_value_locale(
        sheet,
        outline,
        range,
        filter,
        ValueLocaleConfig::en_us(),
    )
}

pub fn apply_autofilter_to_outline_with_value_locale(
    sheet: &Worksheet,
    outline: &mut Outline,
    range: Range,
    filter: Option<&SheetAutoFilter>,
    value_locale: ValueLocaleConfig,
) -> FilterResult {
    let row_count = range.height() as usize;
    let col_count = range.width() as usize;

    if row_count == 0 || col_count == 0 {
        return FilterResult {
            visible_rows: Vec::new(),
            hidden_sheet_rows: Vec::new(),
        };
    }

    // Always clear any existing filter-hidden flags for the data rows within the range.
    // AutoFilter treats the first row as the header row, so we never set filter hidden on it.
    // Use saturating arithmetic so we don't panic on overflow for very large row indices.
    // (Note: Outline metadata is stored as `u32` indices today, so callers cannot represent
    // a 1-based row number greater than `u32::MAX` anyway.)
    let header_row_1based = range.start.row.saturating_add(1);
    let data_start_row_1based = header_row_1based.saturating_add(1);
    let end_row_1based = range.end.row.saturating_add(1);
    if data_start_row_1based <= end_row_1based {
        outline
            .rows
            .clear_filter_hidden_range(data_start_row_1based, end_row_1based);
    }

    let Some(filter) = filter else {
        return FilterResult {
            visible_rows: vec![true; row_count],
            hidden_sheet_rows: Vec::new(),
        };
    };

    let Ok(filter) = AutoFilter::try_from_model_with_value_locale(filter, value_locale) else {
        return FilterResult {
            visible_rows: vec![true; row_count],
            hidden_sheet_rows: Vec::new(),
        };
    };

    let range_ref = crate::sort_filter::RangeRef {
        start_row: range.start.row as usize,
        start_col: range.start.col as usize,
        end_row: range.end.row as usize,
        end_col: range.end.col as usize,
    };

    let mut rows = Vec::with_capacity(row_count);
    for local_row in 0..row_count {
        let row_idx = range.start.row + local_row as u32;
        let mut row = Vec::with_capacity(col_count);
        for local_col in 0..col_count {
            let col_idx = range.start.col + local_col as u32;
            let value = sheet.value(CellRef::new(row_idx, col_idx));
            row.push(model_cell_value_to_sort_value(&value));
        }
        rows.push(row);
    }

    let range_data = crate::sort_filter::RangeData::new(range_ref, rows)
        .expect("worksheet range should always produce rectangular RangeData");

    let result = apply_autofilter_with_value_locale(&range_data, &filter, value_locale);

    for hidden_row_0based in &result.hidden_sheet_rows {
        let row_1based = u32::try_from(*hidden_row_0based)
            .unwrap_or(u32::MAX)
            .saturating_add(1);
        outline.rows.set_filter_hidden(row_1based, true);
    }

    result
}

fn permute_row_properties(
    sheet: &mut Worksheet,
    start_row: u32,
    header_rows: usize,
    end_row: u32,
    perm: &RowPermutation,
) {
    let data_start = start_row + header_rows as u32;
    if data_start > end_row {
        return;
    }

    let mut extracted: Vec<(u32, RowProperties)> = Vec::new();
    for row in data_start..=end_row {
        if let Some(props) = sheet.row_properties.remove(&row) {
            extracted.push((row, props));
        }
    }

    for (old_row, props) in extracted {
        let local_old = (old_row - start_row) as usize;
        let local_new = perm.old_to_new[local_old];
        let new_row = start_row + local_new as u32;
        sheet.row_properties.insert(new_row, props);
    }
}

fn rewrite_formula_for_move(formula: &str, from: CellAddr, to: CellAddr) -> Option<String> {
    let opts = ParseOptions {
        locale: LocaleConfig::en_us(),
        reference_style: crate::ReferenceStyle::A1,
        normalize_relative_to: Some(from),
    };
    let ast = parse_formula(formula, opts).ok()?;

    let out_opts = SerializeOptions {
        locale: LocaleConfig::en_us(),
        reference_style: crate::ReferenceStyle::A1,
        include_xlfn_prefix: true,
        origin: Some(to),
        // `formula-model` stores formulas without a leading '='.
        omit_equals: true,
    };

    ast.to_string(out_opts).ok()
}

fn model_cell_value_to_sort_value(value: &ModelCellValue) -> CellValue {
    match value {
        ModelCellValue::Empty => CellValue::Blank,
        ModelCellValue::Number(n) => CellValue::Number(*n),
        ModelCellValue::String(s) => CellValue::Text(s.clone()),
        ModelCellValue::Boolean(b) => CellValue::Bool(*b),
        ModelCellValue::Error(err) => CellValue::Error(*err),
        ModelCellValue::RichText(rt) => CellValue::Text(rt.plain_text().to_string()),
        ModelCellValue::Entity(entity) => CellValue::Text(entity.display_value.clone()),
        ModelCellValue::Image(image) => image_alt_text_to_sort_value(image.alt_text.as_deref()),
        ModelCellValue::Record(record) => {
            if let Some(display_field) = record.display_field.as_deref() {
                if let Some(value) = record.fields.get(display_field) {
                    let display_value = match value {
                        ModelCellValue::Empty => Some(CellValue::Blank),
                        ModelCellValue::Number(n) => Some(CellValue::Number(*n)),
                        ModelCellValue::String(s) => Some(CellValue::Text(s.clone())),
                        ModelCellValue::Boolean(b) => Some(CellValue::Bool(*b)),
                        ModelCellValue::Error(err) => Some(CellValue::Error(*err)),
                        ModelCellValue::RichText(rt) => {
                            Some(CellValue::Text(rt.plain_text().to_string()))
                        }
                        // If the display field points at another rich value, degrade it to the same
                        // display string the user sees in the grid. This matches `formula-model`'s
                        // `RecordValue::Display` semantics.
                        ModelCellValue::Entity(entity) => (!entity.display_value.is_empty())
                            .then(|| CellValue::Text(entity.display_value.clone()))
                            .or(Some(CellValue::Blank)),
                        ModelCellValue::Record(record) => {
                            let display = record.to_string();
                            if display.is_empty() {
                                Some(CellValue::Blank)
                            } else {
                                Some(CellValue::Text(display))
                            }
                        }
                        ModelCellValue::Image(image) => {
                            Some(image_alt_text_to_sort_value(image.alt_text.as_deref()))
                        }
                        _ => None,
                    };
                    if let Some(value) = display_value {
                        return value;
                    }
                }
            }

            if record.display_value.is_empty() {
                CellValue::Blank
            } else {
                CellValue::Text(record.display_value.clone())
            }
        }
        // Rich value variants are represented as `{type, value}` in `formula-model` for stable IPC.
        //
        // Keep a wildcard arm for forward-compatibility with new `formula-model::CellValue`
        // variants. Best-effort: attempt to degrade the value to a scalar sort/filter value.
        other => match other {
            ModelCellValue::Array(_) | ModelCellValue::Spill(_) => CellValue::Blank,
            _ => rich_model_cell_value_to_sort_value(other).unwrap_or(CellValue::Blank),
        },
    }
}

fn image_alt_text_to_sort_value(alt_text: Option<&str>) -> CellValue {
    let display = alt_text.filter(|s| !s.is_empty()).unwrap_or("[Image]");
    CellValue::Text(display.to_string())
}

fn image_payload_to_sort_value(payload: Option<&serde_json::Value>) -> CellValue {
    let alt_text = payload.and_then(|value| {
        value
            .get("altText")
            .or_else(|| value.get("alt_text"))
            .and_then(|v| v.as_str())
    });
    image_alt_text_to_sort_value(alt_text)
}

fn rich_model_cell_value_to_sort_value(value: &ModelCellValue) -> Option<CellValue> {
    let serialized = serde_json::to_value(value).ok()?;
    let value_type = serialized.get("type")?.as_str()?;

    match value_type {
        "entity" => {
            let value = serialized.get("value")?;
            let display_value = value
                // `formula-model` uses camelCase for rich value payloads, but we also accept legacy
                // snake_case / display aliases for robustness.
                .get("displayValue")
                // Back-compat: earlier payloads used snake_case.
                .or_else(|| value.get("display_value"))
                .or_else(|| value.get("display"))
                .and_then(|v| v.as_str())?;
            Some(CellValue::Text(display_value.to_string()))
        }
        "image" => Some(image_payload_to_sort_value(serialized.get("value"))),
        "record" => {
            let record = serialized.get("value")?;
            // Some legacy IPC payloads represented record values as a simple display string.
            if let Some(display) = record.as_str() {
                return if display.is_empty() {
                    Some(CellValue::Blank)
                } else {
                    Some(CellValue::Text(display.to_string()))
                };
            }

            // Prefer using `displayField` when it points at a scalar, since this preserves the
            // value's natural sort order (e.g. numbers sort numerically).
            let display_field = record
                .get("displayField")
                .or_else(|| record.get("display_field"))
                .and_then(|v| v.as_str());
            if let Some(display_field) = display_field {
                if let Some(fields) = record.get("fields").and_then(|v| v.as_object()) {
                    if let Some(display_value) = fields.get(display_field) {
                        if let Some(display_value_type) =
                            display_value.get("type").and_then(|v| v.as_str())
                        {
                            let parsed = match display_value_type {
                                "empty" => Some(CellValue::Blank),
                                "number" => display_value
                                    .get("value")
                                    .and_then(|v| v.as_f64())
                                    .map(CellValue::Number),
                                "string" => display_value
                                    .get("value")
                                    .and_then(|v| v.as_str())
                                    .map(|s| CellValue::Text(s.to_string())),
                                "boolean" => display_value
                                    .get("value")
                                    .and_then(|v| v.as_bool())
                                    .map(CellValue::Bool),
                                "error" => display_value.get("value").and_then(|v| v.as_str()).map(
                                    |err_str| {
                                        let err = err_str
                                            .parse::<formula_model::ErrorValue>()
                                            .unwrap_or(formula_model::ErrorValue::Unknown);
                                        CellValue::Error(err)
                                    },
                                ),
                                "rich_text" => display_value
                                    .get("value")
                                    .and_then(|v| v.get("text"))
                                    .and_then(|v| v.as_str())
                                    .map(|s| CellValue::Text(s.to_string())),
                                "image" => {
                                    Some(image_payload_to_sort_value(display_value.get("value")))
                                }
                                // Degrade nested rich values (e.g. records whose display field is
                                // an entity/record) using the same logic as the main conversion.
                                //
                                // Note: `"image"` is handled explicitly above so we can prefer its
                                // alt text without an extra deserialize roundtrip.
                                "entity" | "record" => {
                                    serde_json::from_value(display_value.clone())
                                        .ok()
                                        .map(|v: ModelCellValue| model_cell_value_to_sort_value(&v))
                                }
                                _ => None,
                            };

                            if parsed.is_some() {
                                return parsed;
                            }
                        }
                    }
                }
            }

            // Prefer a precomputed display string when present.
            if let Some(display) = record
                .get("displayValue")
                .or_else(|| record.get("display_value"))
                .or_else(|| record.get("display"))
                .and_then(|v| v.as_str())
            {
                return if display.is_empty() {
                    Some(CellValue::Blank)
                } else {
                    Some(CellValue::Text(display.to_string()))
                };
            }
            None
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::{
        image_payload_to_sort_value, model_cell_value_to_sort_value,
        rich_model_cell_value_to_sort_value,
    };
    use crate::sort_filter::CellValue;
    use formula_model::CellValue as ModelCellValue;
    use formula_model::ErrorValue;
    use serde_json::json;

    #[test]
    fn model_cell_value_to_sort_value_entity_record() {
        fn from_json_or_skip_unknown_variant(json: serde_json::Value) -> Option<ModelCellValue> {
            match serde_json::from_value(json) {
                Ok(value) => Some(value),
                Err(err) => {
                    let msg = err.to_string();
                    if msg.contains("unknown variant") {
                        None
                    } else {
                        panic!("failed to deserialize ModelCellValue: {msg}");
                    }
                }
            }
        }

        let Some(entity) = from_json_or_skip_unknown_variant(json!({
            "type": "entity",
            "value": {
                "display": "Entity display"
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&entity),
            CellValue::Text("Entity display".to_string())
        );

        let Some(image) = from_json_or_skip_unknown_variant(json!({
            "type": "image",
            "value": {
                "imageId": "logo.png",
                "altText": "Logo"
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&image),
            CellValue::Text("Logo".to_string())
        );

        let Some(image_without_alt_text) = from_json_or_skip_unknown_variant(json!({
            "type": "image",
            "value": {
                "imageId": "logo.png"
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&image_without_alt_text),
            CellValue::Text("[Image]".to_string())
        );

        // Canonical camelCase field name (`displayValue`) should also deserialize.
        let Some(entity_camel_case) = from_json_or_skip_unknown_variant(json!({
            "type": "entity",
            "value": {
                "displayValue": "Entity display (camelCase)"
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&entity_camel_case),
            CellValue::Text("Entity display (camelCase)".to_string())
        );
        let Some(record) = from_json_or_skip_unknown_variant(json!({
            "type": "record",
            "value": {
                "displayField": "name",
                "fields": {
                    "name": { "type": "string", "value": "Alice" },
                    "age": { "type": "number", "value": 42.0 }
                }
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&record),
            CellValue::Text("Alice".to_string())
        );

        let Some(record_number) = from_json_or_skip_unknown_variant(json!({
            "type": "record",
            "value": {
                "displayField": "age",
                "fields": {
                    "name": { "type": "string", "value": "Alice" },
                    "age": { "type": "number", "value": 42.0 }
                }
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&record_number),
            CellValue::Number(42.0)
        );

        let Some(record_bool) = from_json_or_skip_unknown_variant(json!({
            "type": "record",
            "value": {
                "displayField": "active",
                "fields": {
                    "active": { "type": "boolean", "value": true }
                }
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&record_bool),
            CellValue::Bool(true)
        );

        let Some(record_error) = from_json_or_skip_unknown_variant(json!({
            "type": "record",
            "value": {
                "displayField": "err",
                "fields": {
                    "err": { "type": "error", "value": "#REF!" }
                }
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&record_error),
            CellValue::Error(ErrorValue::Ref)
        );

        let Some(record_rich_text) = from_json_or_skip_unknown_variant(json!({
            "type": "record",
            "value": {
                "displayField": "rt",
                "fields": {
                    "rt": { "type": "rich_text", "value": { "text": "Hello", "runs": [] } }
                }
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&record_rich_text),
            CellValue::Text("Hello".to_string())
        );

        // Records prefer the display field when it points at a scalar value. Rich values (like
        // entities or nested records) degrade to their display string.
        let record_entity_display_field: ModelCellValue = serde_json::from_value(json!({
            "type": "record",
            "value": {
                "displayField": "entity",
                "fields": {
                    "entity": { "type": "entity", "value": { "display": "Nested entity" } }
                }
            }
        }))
        .expect("record should deserialize");
        assert_eq!(
            model_cell_value_to_sort_value(&record_entity_display_field),
            CellValue::Text("Nested entity".to_string())
        );

        let record_image_display_field: ModelCellValue = serde_json::from_value(json!({
            "type": "record",
            "value": {
                "displayField": "logo",
                "fields": {
                    "logo": { "type": "image", "value": { "imageId": "logo.png", "altText": "Logo" } }
                }
            }
        }))
        .expect("record should deserialize");
        assert_eq!(
            model_cell_value_to_sort_value(&record_image_display_field),
            CellValue::Text("Logo".to_string())
        );

        let record_image_without_alt_text_display_field: ModelCellValue =
            serde_json::from_value(json!({
                "type": "record",
                "value": {
                    "displayField": "logo",
                    "fields": {
                        "logo": { "type": "image", "value": { "imageId": "logo.png" } }
                    }
                }
            }))
            .expect("record should deserialize");
        assert_eq!(
            model_cell_value_to_sort_value(&record_image_without_alt_text_display_field),
            CellValue::Text("[Image]".to_string())
        );

        // Invalid display field falls back to the record's display string (if present),
        // otherwise blank.
        let record_missing_display_field: ModelCellValue = serde_json::from_value(json!({
            "type": "record",
            "value": {
                "displayField": "missing",
                "fields": {
                    "name": { "type": "string", "value": "Alice" }
                }
            }
        }))
        .expect("record should deserialize");
        assert_eq!(
            model_cell_value_to_sort_value(&record_missing_display_field),
            CellValue::Blank
        );

        let Some(record_invalid_display_field_with_display) =
            from_json_or_skip_unknown_variant(json!({
                "type": "record",
                "value": {
                    "displayField": "missing",
                    "displayValue": "Fallback display",
                    "fields": {
                        "name": { "type": "string", "value": "Alice" }
                    }
                }
            }))
        else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&record_invalid_display_field_with_display),
            CellValue::Text("Fallback display".to_string())
        );

        // When the display field resolves to an explicit blank/empty cell value, it should take
        // precedence over any legacy display string.
        let record_empty_display_field_value: ModelCellValue = serde_json::from_value(json!({
            "type": "record",
            "value": {
                "displayField": "name",
                "displayValue": "Fallback display",
                "fields": {
                    "name": { "type": "empty" }
                }
            }
        }))
        .expect("record should deserialize");
        assert_eq!(
            model_cell_value_to_sort_value(&record_empty_display_field_value),
            CellValue::Blank
        );

        // Records degrade nested records/entities to their display strings.
        let record_record_display_field: ModelCellValue = serde_json::from_value(json!({
            "type": "record",
            "value": {
                "displayField": "nested",
                "fields": {
                    "nested": {
                        "type": "record",
                        "value": {
                            "displayField": "name",
                            "fields": {
                                "name": { "type": "string", "value": "Bob" }
                            }
                        }
                    }
                }
            }
        }))
        .expect("record should deserialize");
        assert_eq!(
            model_cell_value_to_sort_value(&record_record_display_field),
            CellValue::Text("Bob".to_string())
        );

        // Records without a display field fall back to their legacy display string.
        let Some(record_display_fallback) = from_json_or_skip_unknown_variant(json!({
            "type": "record",
            "value": {
                "display": "Record display"
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&record_display_fallback),
            CellValue::Text("Record display".to_string())
        );

        // Canonical camelCase field name (`displayValue`) should also deserialize.
        let Some(record_display_value_fallback) = from_json_or_skip_unknown_variant(json!({
            "type": "record",
            "value": {
                "displayValue": "Record display (camelCase)"
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&record_display_value_fallback),
            CellValue::Text("Record display (camelCase)".to_string())
        );

        // Empty display values should degrade to blank.
        let Some(record_blank_display) = from_json_or_skip_unknown_variant(json!({
            "type": "record",
            "value": {
                "displayValue": ""
            }
        })) else {
            return;
        };
        assert_eq!(
            model_cell_value_to_sort_value(&record_blank_display),
            CellValue::Blank
        );
    }

    #[test]
    fn rich_model_cell_value_to_sort_value_image() {
        let image: ModelCellValue = serde_json::from_value(json!({
            "type": "image",
            "value": {
                "imageId": "logo.png",
                "altText": "Logo"
            }
        }))
        .expect("image should deserialize");
        assert_eq!(
            rich_model_cell_value_to_sort_value(&image),
            Some(CellValue::Text("Logo".to_string()))
        );

        let image_no_alt: ModelCellValue = serde_json::from_value(json!({
            "type": "image",
            "value": {
                "imageId": "logo.png"
            }
        }))
        .expect("image should deserialize");
        assert_eq!(
            rich_model_cell_value_to_sort_value(&image_no_alt),
            Some(CellValue::Text("[Image]".to_string()))
        );

        let record_display_field_image: ModelCellValue = serde_json::from_value(json!({
            "type": "record",
            "value": {
                "displayField": "logo",
                "fields": {
                    "logo": { "type": "image", "value": { "imageId": "logo.png", "altText": "Logo" } }
                }
            }
        }))
        .expect("record should deserialize");
        assert_eq!(
            rich_model_cell_value_to_sort_value(&record_display_field_image),
            Some(CellValue::Text("Logo".to_string()))
        );

        let record_display_field_image_no_alt: ModelCellValue = serde_json::from_value(json!({
            "type": "record",
            "value": {
                "displayField": "logo",
                "fields": {
                    "logo": { "type": "image", "value": { "imageId": "logo.png" } }
                }
            }
        }))
        .expect("record should deserialize");
        assert_eq!(
            rich_model_cell_value_to_sort_value(&record_display_field_image_no_alt),
            Some(CellValue::Text("[Image]".to_string()))
        );
    }

    #[test]
    fn image_payload_to_sort_value_alt_text_aliases() {
        let payload_camel = json!({
            "imageId": "logo.png",
            "altText": "Logo (camelCase)"
        });
        assert_eq!(
            image_payload_to_sort_value(Some(&payload_camel)),
            CellValue::Text("Logo (camelCase)".to_string())
        );

        let payload_snake = json!({
            "imageId": "logo.png",
            "alt_text": "Logo (snake_case)"
        });
        assert_eq!(
            image_payload_to_sort_value(Some(&payload_snake)),
            CellValue::Text("Logo (snake_case)".to_string())
        );

        let payload_empty_alt = json!({
            "imageId": "logo.png",
            "altText": ""
        });
        assert_eq!(
            image_payload_to_sort_value(Some(&payload_empty_alt)),
            CellValue::Text("[Image]".to_string())
        );
    }
}
