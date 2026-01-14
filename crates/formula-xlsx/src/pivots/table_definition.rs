use std::collections::{BTreeMap, BTreeSet};
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::local_name;
use crate::{XlsxError, XlsxPackage};

/// Metadata extracted from an `xl/pivotTables/pivotTable*.xml` part.
///
/// This struct intentionally captures only the subset of the pivot table definition that we need
/// for sheet rendering and recomputation, while leaving the original XML untouched for round-trip
/// fidelity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTableDefinition {
    /// OPC part path, e.g. `xl/pivotTables/pivotTable1.xml`.
    pub path: String,
    pub name: Option<String>,
    pub cache_id: Option<u32>,
    /// Styling hints from `<pivotTableStyleInfo>`.
    pub style_info: Option<PivotTableStyleInfo>,
    /// `pivotTableDefinition@applyNumberFormats` (if present).
    pub apply_number_formats: Option<bool>,
    /// `pivotTableDefinition@applyBorderFormats` (if present).
    pub apply_border_formats: Option<bool>,
    /// `pivotTableDefinition@applyFontFormats` (if present).
    pub apply_font_formats: Option<bool>,
    /// `pivotTableDefinition@applyPatternFormats` (if present).
    pub apply_pattern_formats: Option<bool>,
    /// `pivotTableDefinition@applyAlignmentFormats` (if present).
    pub apply_alignment_formats: Option<bool>,
    /// `pivotTableDefinition@applyWidthHeightFormats` (if present).
    pub apply_width_height_formats: Option<bool>,
    /// Output range on the destination worksheet (A1-style range).
    pub location_ref: Option<String>,
    pub first_header_row: Option<u32>,
    pub first_data_row: Option<u32>,
    pub first_data_col: Option<u32>,
    /// `pivotTableDefinition@dataOnRows` (defaults to `false`).
    pub data_on_rows: bool,
    /// `pivotTableDefinition@rowGrandTotals` (defaults to `true`).
    pub row_grand_totals: bool,
    /// `pivotTableDefinition@colGrandTotals` (defaults to `true`).
    pub col_grand_totals: bool,
    /// `pivotTableDefinition@outline` (if present).
    pub outline: Option<bool>,
    /// `pivotTableDefinition@compact` (if present).
    pub compact: Option<bool>,
    /// `pivotTableDefinition@compactData` (if present).
    pub compact_data: Option<bool>,
    /// `pivotTableDefinition@subtotalLocation` (if present).
    pub subtotal_location: Option<String>,
    /// `<pivotFields>` -> `<pivotField>` entries (in order).
    pub pivot_fields: Vec<PivotTableField>,
    /// `<rowFields>` -> `<field x="...">` indices (in order).
    pub row_fields: Vec<u32>,
    /// `<colFields>` -> `<field x="...">` indices (in order).
    pub col_fields: Vec<u32>,
    /// `<pageFields>` -> `<pageField>` entries (in order).
    pub page_field_entries: Vec<PivotTablePageField>,
    /// `<pageFields>` -> field indices (in order).
    ///
    /// This is a compatibility view over [`PivotTableDefinition::page_field_entries`].
    pub page_fields: Vec<u32>,
    /// `<dataFields>` -> `<dataField>` entries (in order).
    pub data_fields: Vec<PivotTableDataField>,
}

impl PivotTableDefinition {
    /// Parse a pivot table definition part (`xl/pivotTables/pivotTable*.xml`).
    pub fn parse(path: &str, xml: &[u8]) -> Result<Self, XlsxError> {
        let mut reader = Reader::from_reader(Cursor::new(xml));
        reader.config_mut().trim_text(true);

        let mut def = PivotTableDefinition {
            path: path.to_string(),
            name: None,
            cache_id: None,
            style_info: None,
            apply_number_formats: None,
            apply_border_formats: None,
            apply_font_formats: None,
            apply_pattern_formats: None,
            apply_alignment_formats: None,
            apply_width_height_formats: None,
            location_ref: None,
            first_header_row: None,
            first_data_row: None,
            first_data_col: None,
            data_on_rows: false,
            row_grand_totals: true,
            col_grand_totals: true,
            outline: None,
            compact: None,
            compact_data: None,
            subtotal_location: None,
            pivot_fields: Vec::new(),
            row_fields: Vec::new(),
            col_fields: Vec::new(),
            page_field_entries: Vec::new(),
            page_fields: Vec::new(),
            data_fields: Vec::new(),
        };

        let mut buf = Vec::new();
        let mut parsed_root = false;
        let mut context: Option<FieldContext> = None;
        let mut pivot_field_idx: Option<usize> = None;

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Start(start) => {
                    parse_start_element(&mut def, &start, &mut parsed_root)?;
                    handle_start_element(&mut def, &start, &mut context, &mut pivot_field_idx, true)?;
                }
                Event::Empty(start) => {
                    parse_start_element(&mut def, &start, &mut parsed_root)?;
                    handle_start_element(&mut def, &start, &mut context, &mut pivot_field_idx, false)?;
                }
                Event::End(end) => {
                    let name = end.name();
                    let tag = local_name(name.as_ref());
                    if tag.eq_ignore_ascii_case(b"rowFields")
                        || tag.eq_ignore_ascii_case(b"colFields")
                        || tag.eq_ignore_ascii_case(b"pageFields")
                        || tag.eq_ignore_ascii_case(b"dataFields")
                    {
                        context = None;
                    } else if tag.eq_ignore_ascii_case(b"pivotField") {
                        pivot_field_idx = None;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        // Keep legacy `page_fields` in sync with the richer `page_field_entries`.
        def.page_fields = def.page_field_entries.iter().map(|pf| pf.fld).collect();

        Ok(def)
    }

    /// Returns the output range of this pivot table (as declared by `<location ref="...">`).
    ///
    /// The returned range is in worksheet coordinates (0-indexed) and is parsed using
    /// [`formula_model::Range::from_a1`]. If the pivot table definition does not specify a
    /// location or the `ref` string is invalid, returns `None`.
    pub fn location_range(&self) -> Option<formula_model::Range> {
        let a1 = self.location_ref.as_deref()?;
        formula_model::Range::from_a1(a1).ok()
    }

    /// Returns the top-left cell of the pivot table output range.
    ///
    /// This is equivalent to `self.location_range().map(|r| r.start)`.
    pub fn location_top_left(&self) -> Option<formula_model::CellRef> {
        self.location_range().map(|r| r.start)
    }
}

impl XlsxPackage {
    /// Parse a pivot table definition part (e.g. `xl/pivotTables/pivotTable1.xml`).
    pub fn pivot_table_definition(&self, part_name: &str) -> Result<PivotTableDefinition, XlsxError> {
        let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
        let xml = self
            .part(part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.to_string()))?;
        PivotTableDefinition::parse(part_name, xml)
    }

    /// Parse all pivot table definition parts in the package.
    pub fn pivot_table_definitions(&self) -> Result<Vec<PivotTableDefinition>, XlsxError> {
        let mut paths: BTreeSet<String> = BTreeSet::new();
        for name in self.part_names() {
            if name.starts_with("xl/pivotTables/") && name.ends_with(".xml") {
                paths.insert(name.to_string());
            }
        }

        let mut out = Vec::with_capacity(paths.len());
        for path in paths {
            out.push(self.pivot_table_definition(&path)?);
        }
        Ok(out)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PivotTableField {
    pub axis: Option<String>,
    pub show_all: Option<bool>,
    pub default_subtotal: Option<bool>,
    /// `pivotField@sortType` (if present).
    ///
    /// Common values include `ascending`, `descending`, and `manual`.
    pub sort_type: Option<String>,
    /// Best-effort representation of manual pivot item ordering, typically sourced from
    /// `<pivotField><items><item .../></items></pivotField>`.
    ///
    /// Note: Excel's schema usually encodes items as indices (`item@x`) into the cache field's
    /// shared items list; some producers may also emit names.
    pub manual_sort_items: Option<Vec<PivotTableFieldItem>>,
    /// Explicit subtotal flags such as `sumSubtotal`, `countSubtotal`, etc.
    ///
    /// Keys are stored exactly as they appear in the XML (minus any namespace prefix).
    pub subtotals: BTreeMap<String, bool>,
}

/// Best-effort representation of an `<item>` entry inside `<pivotField><items>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PivotTableFieldItem {
    /// Shared-item index (`item@x`).
    Index(u32),
    /// Producer-specific item label (`item@n` / `item@name`).
    Name(String),
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PivotTableStyleInfo {
    pub name: Option<String>,
    pub show_row_headers: Option<bool>,
    pub show_col_headers: Option<bool>,
    pub show_row_stripes: Option<bool>,
    pub show_col_stripes: Option<bool>,
    pub show_last_column: Option<bool>,
    pub show_first_column: Option<bool>,
}

/// An entry inside `<pageFields>` describing a report filter ("page field").
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PivotTablePageField {
    /// Field index (`pageField@fld`).
    pub fld: u32,
    /// Selected item index (`pageField@item`), often `-1` for `(All)`.
    pub item: Option<i32>,
    /// Hierarchy index (`pageField@hier`) when present.
    pub hier: Option<u32>,
    /// Display name (`pageField@name`) when present.
    pub name: Option<String>,
    /// `pageField@hierarchical` when present (non-standard).
    pub hierarchical: Option<bool>,
    /// `pageField@multipleItemSelectionAllowed` when present.
    pub multiple_item_selection_allowed: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PivotTableDataField {
    pub fld: Option<u32>,
    pub name: Option<String>,
    pub subtotal: Option<String>,
    pub num_fmt_id: Option<u32>,
    pub base_field: Option<u32>,
    pub base_item: Option<u32>,
    pub show_data_as: Option<String>,
    pub calculated: Option<bool>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FieldContext {
    Row,
    Col,
    Page,
    Data,
}

fn handle_start_element(
    def: &mut PivotTableDefinition,
    start: &quick_xml::events::BytesStart<'_>,
    context: &mut Option<FieldContext>,
    pivot_field_idx: &mut Option<usize>,
    open_container: bool,
) -> Result<(), XlsxError> {
    let name = start.name();
    let tag = local_name(name.as_ref());

    if open_container {
        if tag.eq_ignore_ascii_case(b"rowFields") {
            *context = Some(FieldContext::Row);
            return Ok(());
        }
        if tag.eq_ignore_ascii_case(b"colFields") {
            *context = Some(FieldContext::Col);
            return Ok(());
        }
        if tag.eq_ignore_ascii_case(b"pageFields") {
            *context = Some(FieldContext::Page);
            return Ok(());
        }
        if tag.eq_ignore_ascii_case(b"dataFields") {
            *context = Some(FieldContext::Data);
            return Ok(());
        }
    }

    if tag.eq_ignore_ascii_case(b"pivotField") {
        let mut field = PivotTableField::default();
        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            let value = attr.unescape_value()?.into_owned();

            if key.eq_ignore_ascii_case(b"axis") {
                field.axis = Some(value);
            } else if key.eq_ignore_ascii_case(b"showAll") {
                field.show_all = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"defaultSubtotal") {
                field.default_subtotal = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"sortType") {
                field.sort_type = Some(value);
            } else if key.eq_ignore_ascii_case(b"sortOrder") {
                // Non-standard producers sometimes emit `sortOrder` instead of `sortType`.
                field.sort_type.get_or_insert(value);
            } else if key.eq_ignore_ascii_case(b"sortAscending") {
                // Another non-standard alias; map to the canonical sort type strings.
                if field.sort_type.is_none() {
                    if let Some(v) = parse_bool(&value) {
                        field.sort_type = Some(if v {
                            "ascending".to_string()
                        } else {
                            "descending".to_string()
                        });
                    }
                }
            } else if key.len() >= b"Subtotal".len()
                && key[key.len() - b"Subtotal".len()..].eq_ignore_ascii_case(b"Subtotal")
                && !key.eq_ignore_ascii_case(b"defaultSubtotal")
            {
                if let Some(v) = parse_bool(&value) {
                    field
                        .subtotals
                        .insert(String::from_utf8_lossy(key).to_string(), v);
                }
            }
        }
        def.pivot_fields.push(field);
        if open_container {
            *pivot_field_idx = def.pivot_fields.len().checked_sub(1);
        } else {
            *pivot_field_idx = None;
        }
        return Ok(());
    }

    // Manual sort order is commonly represented as a sequence of `<item>` entries in the pivot
    // field's `<items>` container. We record these in the order they appear.
    if tag.eq_ignore_ascii_case(b"item") {
        let Some(field_idx) = *pivot_field_idx else {
            return Ok(());
        };
        let Some(field) = def.pivot_fields.get_mut(field_idx) else {
            return Ok(());
        };

        let mut item_index: Option<u32> = None;
        let mut item_name: Option<String> = None;

        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            let value = attr.unescape_value()?.into_owned();

            if key.eq_ignore_ascii_case(b"x") {
                item_index = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"n") || key.eq_ignore_ascii_case(b"name") {
                item_name = Some(value);
            }
        }

        let item = if let Some(name) = item_name {
            Some(PivotTableFieldItem::Name(name))
        } else if let Some(idx) = item_index {
            Some(PivotTableFieldItem::Index(idx))
        } else {
            None
        };

        if let Some(item) = item {
            field
                .manual_sort_items
                .get_or_insert_with(Vec::new)
                .push(item);
        }
        return Ok(());
    }

    if tag.eq_ignore_ascii_case(b"pivotTableStyleInfo") {
        let mut style = PivotTableStyleInfo::default();
        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            let value = attr.unescape_value()?.into_owned();

            if key.eq_ignore_ascii_case(b"name") {
                style.name = Some(value);
            } else if key.eq_ignore_ascii_case(b"showRowHeaders") {
                style.show_row_headers = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"showColHeaders") {
                style.show_col_headers = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"showRowStripes") {
                style.show_row_stripes = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"showColStripes") {
                style.show_col_stripes = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"showLastColumn") {
                style.show_last_column = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"showFirstColumn") {
                style.show_first_column = parse_bool(&value);
            }
        }
        def.style_info = Some(style);
        return Ok(());
    }

    if tag.eq_ignore_ascii_case(b"field") {
        let Some(ctx) = *context else {
            return Ok(());
        };
        if ctx == FieldContext::Data {
            return Ok(());
        }
        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"x") {
                if let Ok(v) = attr.unescape_value()?.trim().parse::<u32>() {
                    match ctx {
                        FieldContext::Row => def.row_fields.push(v),
                        FieldContext::Col => def.col_fields.push(v),
                        FieldContext::Page => def.page_field_entries.push(PivotTablePageField {
                            fld: v,
                            ..PivotTablePageField::default()
                        }),
                        FieldContext::Data => {}
                    }
                }
            }
        }
        return Ok(());
    }

    // Some pivot tables use `<pageField fld="...">` for page fields.
    if tag.eq_ignore_ascii_case(b"pageField") && *context == Some(FieldContext::Page) {
        let mut fld: Option<u32> = None;
        let mut item: Option<i32> = None;
        let mut hier: Option<u32> = None;
        let mut name: Option<String> = None;
        let mut hierarchical: Option<bool> = None;
        let mut multiple_item_selection_allowed: Option<bool> = None;
        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            let value = attr.unescape_value()?.into_owned();
            if key.eq_ignore_ascii_case(b"fld") {
                fld = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"item") {
                item = value.trim().parse::<i32>().ok();
            } else if key.eq_ignore_ascii_case(b"hier") {
                hier = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"name") {
                name = Some(value);
            } else if key.eq_ignore_ascii_case(b"hierarchical") {
                hierarchical = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"multipleItemSelectionAllowed") {
                multiple_item_selection_allowed = parse_bool(&value);
            }
        }
        if let Some(fld) = fld {
            def.page_field_entries.push(PivotTablePageField {
                fld,
                item,
                hier,
                name,
                hierarchical,
                multiple_item_selection_allowed,
            });
        }
        return Ok(());
    }

    if tag.eq_ignore_ascii_case(b"dataField") && *context == Some(FieldContext::Data) {
        let mut field = PivotTableDataField::default();
        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            let value = attr.unescape_value()?.into_owned();

            if key.eq_ignore_ascii_case(b"fld") {
                field.fld = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"name") {
                field.name = Some(value);
            } else if key.eq_ignore_ascii_case(b"subtotal") {
                field.subtotal = Some(value);
            } else if key.eq_ignore_ascii_case(b"numFmtId") {
                field.num_fmt_id = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"baseField") {
                field.base_field = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"baseItem") {
                field.base_item = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"showDataAs") {
                field.show_data_as = Some(value);
            } else if key.eq_ignore_ascii_case(b"calculated") {
                field.calculated = parse_bool(&value);
            }
        }
        def.data_fields.push(field);
        return Ok(());
    }

    Ok(())
}

fn parse_start_element(
    def: &mut PivotTableDefinition,
    start: &quick_xml::events::BytesStart<'_>,
    parsed_root: &mut bool,
) -> Result<(), XlsxError> {
    let name = start.name();
    let tag = local_name(name.as_ref());
    if !*parsed_root && tag.eq_ignore_ascii_case(b"pivotTableDefinition") {
        *parsed_root = true;
        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            let value = attr.unescape_value()?.into_owned();

            if key.eq_ignore_ascii_case(b"name") {
                def.name = Some(value);
            } else if key.eq_ignore_ascii_case(b"cacheId") {
                def.cache_id = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"dataOnRows") {
                if let Some(v) = parse_bool(&value) {
                    def.data_on_rows = v;
                }
            } else if key.eq_ignore_ascii_case(b"applyNumberFormats") {
                def.apply_number_formats = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"applyBorderFormats") {
                def.apply_border_formats = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"applyFontFormats") {
                def.apply_font_formats = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"applyPatternFormats") {
                def.apply_pattern_formats = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"applyAlignmentFormats") {
                def.apply_alignment_formats = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"applyWidthHeightFormats") {
                def.apply_width_height_formats = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"rowGrandTotals") {
                if let Some(v) = parse_bool(&value) {
                    def.row_grand_totals = v;
                }
            } else if key.eq_ignore_ascii_case(b"colGrandTotals") {
                if let Some(v) = parse_bool(&value) {
                    def.col_grand_totals = v;
                }
            } else if key.eq_ignore_ascii_case(b"outline") {
                def.outline = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"compact") {
                def.compact = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"compactData") {
                def.compact_data = parse_bool(&value);
            } else if key.eq_ignore_ascii_case(b"subtotalLocation") {
                def.subtotal_location = Some(value);
            }
        }
        return Ok(());
    }

    if tag.eq_ignore_ascii_case(b"location") {
        for attr in start.attributes().with_checks(false) {
            let attr = attr?;
            let key = local_name(attr.key.as_ref());
            let value = attr.unescape_value()?.into_owned();

            if key.eq_ignore_ascii_case(b"ref") {
                def.location_ref = Some(value);
            } else if key.eq_ignore_ascii_case(b"firstHeaderRow") {
                def.first_header_row = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"firstDataRow") {
                def.first_data_row = value.trim().parse::<u32>().ok();
            } else if key.eq_ignore_ascii_case(b"firstDataCol") {
                def.first_data_col = value.trim().parse::<u32>().ok();
            }
        }
    }

    Ok(())
}

fn parse_bool(value: &str) -> Option<bool> {
    match value.trim() {
        "1" | "true" | "TRUE" | "True" => Some(true),
        "0" | "false" | "FALSE" | "False" => Some(false),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use formula_model::{CellRef, Range};
    use pretty_assertions::assert_eq;

    #[test]
    fn parses_location_and_layout_flags() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:pivotTableDefinition xmlns:p="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1"
  cacheId=" 7 "
  dataOnRows="1"
  rowGrandTotals="0"
  colGrandTotals="1"
  outline="1"
  compact="0"
  compactData="1">
  <p:location ref="B3:F20" firstHeaderRow=" 2 " firstDataRow=" 3 " firstDataCol=" 2 "/>
</p:pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");

        assert_eq!(parsed.name.as_deref(), Some("PivotTable1"));
        assert_eq!(parsed.cache_id, Some(7));
        assert_eq!(parsed.location_ref.as_deref(), Some("B3:F20"));
        assert_eq!(parsed.first_header_row, Some(2));
        assert_eq!(parsed.first_data_row, Some(3));
        assert_eq!(parsed.first_data_col, Some(2));
        assert_eq!(parsed.data_on_rows, true);
        assert_eq!(parsed.row_grand_totals, false);
        assert_eq!(parsed.col_grand_totals, true);
        assert_eq!(parsed.outline, Some(true));
        assert_eq!(parsed.compact, Some(false));
        assert_eq!(parsed.compact_data, Some(true));
    }

    #[test]
    fn parses_apply_format_flags() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1"
  cacheId="1"
  applyNumberFormats="1"
  applyBorderFormats="0"
  applyFontFormats="true"
  applyPatternFormats="false"
  applyAlignmentFormats="1"
  applyWidthHeightFormats="0"/>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");

        assert_eq!(parsed.apply_number_formats, Some(true));
        assert_eq!(parsed.apply_border_formats, Some(false));
        assert_eq!(parsed.apply_font_formats, Some(true));
        assert_eq!(parsed.apply_pattern_formats, Some(false));
        assert_eq!(parsed.apply_alignment_formats, Some(true));
        assert_eq!(parsed.apply_width_height_formats, Some(false));
    }

    #[test]
    fn parses_pivot_table_style_info_empty_element() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:pivotTableDefinition xmlns:p="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1"
  cacheId="1">
  <p:pivotTableStyleInfo name="PivotStyleMedium9"
    showRowHeaders="1"
    showColHeaders="0"
    showRowStripes="1"
    showColStripes="0"
    showLastColumn="1"/>
</p:pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");
        let style = parsed.style_info.expect("style info parsed");

        assert_eq!(style.name.as_deref(), Some("PivotStyleMedium9"));
        assert_eq!(style.show_row_headers, Some(true));
        assert_eq!(style.show_col_headers, Some(false));
        assert_eq!(style.show_row_stripes, Some(true));
        assert_eq!(style.show_col_stripes, Some(false));
        assert_eq!(style.show_last_column, Some(true));
    }

    #[test]
    fn parses_pivot_table_style_info_start_end_element_with_prefixes() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:pivotTableDefinition xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1"
  cacheId="1">
  <x:pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="true" showColHeaders="false"></x:pivotTableStyleInfo>
 </x:pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");
        let style = parsed.style_info.expect("style info parsed");

        assert_eq!(style.name.as_deref(), Some("PivotStyleLight16"));
        assert_eq!(style.show_row_headers, Some(true));
        assert_eq!(style.show_col_headers, Some(false));
    }

    #[test]
    fn parses_page_field_item_selection() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pageFields count="2">
    <pageField fld="2" item="3"/>
    <pageField fld="5" item="-1"/>
  </pageFields>
</pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");

        assert_eq!(
            parsed.page_field_entries,
            vec![
                PivotTablePageField {
                    fld: 2,
                    item: Some(3),
                    hier: None,
                    name: None,
                    hierarchical: None,
                    multiple_item_selection_allowed: None,
                },
                PivotTablePageField {
                    fld: 5,
                    item: Some(-1),
                    hier: None,
                    name: None,
                    hierarchical: None,
                    multiple_item_selection_allowed: None,
                },
            ]
        );
        assert_eq!(parsed.page_fields, vec![2, 5]);
    }

    #[test]
    fn location_range_parses_valid_a1() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <location ref="B3:F20"/>
</pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");

        assert_eq!(parsed.location_range(), Some(Range::from_a1("B3:F20").unwrap()));
        assert_eq!(parsed.location_top_left(), Some(CellRef::new(2, 1)));
    }

    #[test]
    fn location_range_invalid_ref_returns_none() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <location ref="not a range"/>
</pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");

        assert_eq!(parsed.location_range(), None);
        assert_eq!(parsed.location_top_left(), None);
    }

    #[test]
    fn parses_pivot_field_sort_type() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pivotFields count="2">
    <pivotField axis="axisRow" sortType="ascending"/>
    <pivotField axis="axisRow" sortType="descending"/>
  </pivotFields>
</pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");

        assert_eq!(parsed.pivot_fields.len(), 2);
        assert_eq!(parsed.pivot_fields[0].sort_type.as_deref(), Some("ascending"));
        assert_eq!(parsed.pivot_fields[1].sort_type.as_deref(), Some("descending"));
    }

    #[test]
    fn parses_manual_sort_items_from_pivot_field_items() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");

        assert_eq!(parsed.pivot_fields.len(), 1);
        assert_eq!(parsed.pivot_fields[0].sort_type.as_deref(), Some("manual"));
        assert_eq!(
            parsed.pivot_fields[0].manual_sort_items.as_deref(),
            Some(&[
                PivotTableFieldItem::Index(2),
                PivotTableFieldItem::Index(0),
                PivotTableFieldItem::Index(1)
            ][..])
        );
    }

    #[test]
    fn parses_manual_sort_items_by_name_when_present() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</pivotTableDefinition>"#;

        let parsed = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", xml)
            .expect("parse pivotTableDefinition");

        assert_eq!(parsed.pivot_fields.len(), 1);
        assert_eq!(
            parsed.pivot_fields[0].manual_sort_items.as_deref(),
            Some(&[
                PivotTableFieldItem::Name("B".to_string()),
                PivotTableFieldItem::Name("A".to_string()),
                PivotTableFieldItem::Name("C".to_string()),
            ][..])
        );
    }
}
