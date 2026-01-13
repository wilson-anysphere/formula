use crate::openxml::{
    local_name, parse_relationships, rels_part_name, resolve_relationship_target, resolve_target,
};
use crate::package::{XlsxError, XlsxPackage};
use crate::sheet_metadata::parse_workbook_sheets;
use crate::DateSystem;
use chrono::NaiveDate;
use formula_engine::date::{serial_to_ymd, ExcelDateSystem};
use formula_model::pivots::slicers::{RowFilter, SlicerSelection, TimelineSelection};
use formula_model::pivots::ScalarValue;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::io::Cursor;

/// Best-effort slicer selection state extracted from `xl/slicerCaches/slicerCache*.xml`.
///
/// When Excel does not persist explicit selection state, slicers behave as "All selected".
/// We represent that as `selected_items: None` (even if `available_items` is known).
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SlicerSelectionState {
    /// Item keys in the order they appear in the cache.
    pub available_items: Vec<String>,
    /// Explicitly selected items. `None` means "All selected".
    pub selected_items: Option<HashSet<String>>,
}

/// Best-effort timeline selection state extracted from timeline parts/caches.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TimelineSelectionState {
    /// Inclusive start date for the selection, formatted as `YYYY-MM-DD`.
    pub start: Option<String>,
    /// Inclusive end date for the selection, formatted as `YYYY-MM-DD`.
    pub end: Option<String>,
}

/// Convert a parsed slicer selection into a model-level row filter.
///
/// This function is intentionally conservative: item keys are treated as text unless
/// the caller provides a resolver via [`slicer_selection_to_row_filter_with_resolver`].
pub fn slicer_selection_to_row_filter(
    field: impl Into<String>,
    selection: &SlicerSelectionState,
) -> RowFilter {
    slicer_selection_to_row_filter_with_resolver(field, selection, |_| None)
}

/// Convert a parsed slicer selection into a pivot-engine filter field.
///
/// This is the pivot-engine equivalent of [`slicer_selection_to_row_filter`]: item keys are treated
/// as text unless the caller provides a resolver via
/// [`slicer_selection_to_engine_filter_field_with_resolver`].
pub fn slicer_selection_to_engine_filter_field(
    field: impl Into<String>,
    selection: &SlicerSelectionState,
) -> formula_engine::pivot::FilterField {
    slicer_selection_to_engine_filter_field_with_resolver(field, selection, |_| None)
}

/// Convert a parsed slicer selection into a model-level row filter, using `resolve` to map
/// slicer item keys to typed [`ScalarValue`]s.
///
/// This is useful when slicer cache items are stored as indices (`x`) into a pivot cache
/// shared-items table; callers can resolve those indices to typed values when available.
pub fn slicer_selection_to_row_filter_with_resolver<F>(
    field: impl Into<String>,
    selection: &SlicerSelectionState,
    mut resolve: F,
) -> RowFilter
where
    F: FnMut(&str) -> Option<ScalarValue>,
{
    let selection = match &selection.selected_items {
        None => SlicerSelection::All,
        Some(items) => {
            let mut selected = HashSet::with_capacity(items.len());
            for item in items {
                selected.insert(resolve(item).unwrap_or_else(|| ScalarValue::from(item.as_str())));
            }
            SlicerSelection::Items(selected)
        }
    };

    RowFilter::Slicer {
        field: field.into(),
        selection,
    }
}

/// Convert a parsed slicer selection into a pivot-engine filter field, using `resolve` to map
/// slicer item keys to typed [`formula_engine::pivot::PivotKeyPart`]s.
///
/// This is useful when slicer cache items are stored as indices (`x`) into a pivot cache
/// shared-items table; callers can resolve those indices to typed values when available.
///
/// If Excel does not persist explicit selection state, slicers behave as "All selected".
/// This is represented by `selection.selected_items = None`, which becomes
/// `FilterField { allowed: None }`.
pub fn slicer_selection_to_engine_filter_field_with_resolver(
    field: impl Into<String>,
    selection: &SlicerSelectionState,
    mut resolve: impl FnMut(&str) -> Option<formula_engine::pivot::PivotKeyPart>,
) -> formula_engine::pivot::FilterField {
    let allowed = match &selection.selected_items {
        None => None,
        Some(items) => {
            let mut selected = HashSet::with_capacity(items.len());
            for item in items {
                selected.insert(resolve(item).unwrap_or_else(|| {
                    formula_engine::pivot::PivotKeyPart::Text(item.clone())
                }));
            }
            Some(selected)
        }
    };

    formula_engine::pivot::FilterField {
        source_field: field.into(),
        allowed,
    }
}

/// Convert a parsed timeline selection into a pivot-engine filter field.
///
/// The pivot engine's filter system currently only supports an allow-list of discrete values.
/// Timeline selections are defined as a date range, which cannot be represented without knowing
/// the set of date values in the underlying cache. For now we return `allowed: None` and leave
/// range filtering to the caller.
pub fn timeline_selection_to_engine_filter_field(
    field: impl Into<String>,
    _selection: &TimelineSelectionState,
) -> formula_engine::pivot::FilterField {
    formula_engine::pivot::FilterField {
        source_field: field.into(),
        allowed: None,
    }
}

/// Convert a parsed timeline selection into a model-level row filter.
///
/// If the ISO date strings cannot be parsed, the corresponding endpoint is left unset.
pub fn timeline_selection_to_row_filter(
    field: impl Into<String>,
    selection: &TimelineSelectionState,
) -> RowFilter {
    let start = selection.start.as_deref().and_then(parse_iso_ymd);
    let end = selection.end.as_deref().and_then(parse_iso_ymd);

    RowFilter::Timeline {
        field: field.into(),
        selection: TimelineSelection { start, end },
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SlicerDefinition {
    pub part_name: String,
    pub name: Option<String>,
    pub uid: Option<String>,
    pub cache_part: Option<String>,
    pub cache_name: Option<String>,
    pub source_name: Option<String>,
    pub connected_pivot_tables: Vec<String>,
    /// Table parts (`xl/tables/table*.xml`) referenced by the slicer cache relationships.
    pub connected_tables: Vec<String>,
    pub placed_on_drawings: Vec<String>,
    /// Sheet parts (worksheets or chartsheets) that host any drawing containing this slicer (e.g.
    /// `xl/worksheets/sheet1.xml`, `xl/chartsheets/sheet1.xml`).
    pub placed_on_sheets: Vec<String>,
    /// Workbook sheet names for [`Self::placed_on_sheets`] when resolvable from `xl/workbook.xml`.
    pub placed_on_sheet_names: Vec<String>,
    pub selection: SlicerSelectionState,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TimelineDefinition {
    pub part_name: String,
    pub name: Option<String>,
    pub uid: Option<String>,
    pub cache_part: Option<String>,
    pub cache_name: Option<String>,
    pub source_name: Option<String>,
    pub base_field: Option<u32>,
    pub level: Option<u32>,
    pub connected_pivot_tables: Vec<String>,
    pub placed_on_drawings: Vec<String>,
    /// Sheet parts (worksheets or chartsheets) that host any drawing containing this timeline (e.g.
    /// `xl/worksheets/sheet1.xml`, `xl/chartsheets/sheet1.xml`).
    pub placed_on_sheets: Vec<String>,
    /// Workbook sheet names for [`Self::placed_on_sheets`] when resolvable from `xl/workbook.xml`.
    pub placed_on_sheet_names: Vec<String>,
    pub selection: TimelineSelectionState,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct PivotSlicerParts {
    pub slicers: Vec<SlicerDefinition>,
    pub timelines: Vec<TimelineDefinition>,
}

impl XlsxPackage {
    /// Parse slicers and timelines out of an XLSX package.
    ///
    /// This parser is intentionally conservative: it extracts the minimum metadata needed to
    /// wire up the UX layer, while leaving the XML untouched for round-trip fidelity.
    pub fn pivot_slicer_parts(&self) -> Result<PivotSlicerParts, XlsxError> {
        parse_pivot_slicer_parts(self)
    }
}

fn parse_pivot_slicer_parts(package: &XlsxPackage) -> Result<PivotSlicerParts, XlsxError> {
    let date_system = package
        .part("xl/workbook.xml")
        .and_then(|bytes| parse_workbook_date_system(bytes).ok())
        .unwrap_or_default();
    let excel_date_system = date_system.to_engine_date_system();

    let slicer_parts = package
        .part_names()
        .filter(|name| name.starts_with("xl/slicers/") && name.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();
    let timeline_parts = package
        .part_names()
        .filter(|name| name.starts_with("xl/timelines/") && name.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let drawing_rels = package
        .part_names()
        .filter(|name| name.starts_with("xl/drawings/_rels/") && name.ends_with(".rels"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let worksheet_rels = package
        .part_names()
        .filter(|name| name.starts_with("xl/worksheets/_rels/") && name.ends_with(".rels"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let chartsheet_rels = package
        .part_names()
        .filter(|name| name.starts_with("xl/chartsheets/_rels/") && name.ends_with(".rels"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let mut slicer_to_drawings: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    let mut timeline_to_drawings: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();

    for rels_name in drawing_rels {
        let Some(rels_bytes) = package.part(&rels_name) else {
            continue;
        };
        let relationships = match parse_relationships(rels_bytes) {
            Ok(relationships) => relationships,
            Err(_) => continue,
        };
        let drawing_part = drawing_part_name_from_rels(&rels_name);
        for rel in relationships {
            let target = resolve_target(&drawing_part, &rel.target);
            if target.starts_with("xl/slicers/") {
                slicer_to_drawings
                    .entry(target)
                    .or_default()
                    .insert(drawing_part.clone());
            } else if target.starts_with("xl/timelines/") {
                timeline_to_drawings
                    .entry(target)
                    .or_default()
                    .insert(drawing_part.clone());
            }
        }
    }

    let mut drawing_to_sheets: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    for rels_name in worksheet_rels.into_iter().chain(chartsheet_rels) {
        let Some(rels_bytes) = package.part(&rels_name) else {
            continue;
        };
        // Best-effort: malformed `.rels` parts are ignored instead of failing slicer parsing.
        let relationships = match parse_relationships(rels_bytes) {
            Ok(relationships) => relationships,
            Err(_) => continue,
        };

        let sheet_part = if rels_name.starts_with("xl/worksheets/_rels/") {
            worksheet_part_name_from_rels(&rels_name)
        } else {
            chartsheet_part_name_from_rels(&rels_name)
        };

        for rel in relationships {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }
            if !is_drawing_relationship_type(&rel.type_uri) {
                continue;
            }
            let target = resolve_target(&sheet_part, &rel.target);
            if target.starts_with("xl/drawings/") {
                drawing_to_sheets
                    .entry(target)
                    .or_default()
                    .insert(sheet_part.clone());
            }
        }
    }

    let sheet_name_by_part = sheet_name_by_part(package);

    let mut slicers = Vec::with_capacity(slicer_parts.len());
    for part_name in slicer_parts {
        let xml = package
            .part(&part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.clone()))?;
        let parsed = parse_slicer_xml(xml)?;

        let cache_part = match parsed.cache_rid.as_deref() {
            Some(rid) => resolve_relationship_target(package, &part_name, rid)?,
            None => None,
        };

        let (cache_name, source_name, connected_pivot_tables, connected_tables) =
            if let Some(cache_part) = cache_part.as_deref() {
                let resolved = resolve_slicer_cache_definition(package, cache_part)?;
                (
                    resolved.cache_name,
                    resolved.source_name,
                    resolved.connected_pivot_tables,
                    resolved.connected_tables,
                )
            } else {
                (None, None, Vec::new(), Vec::new())
            };

        let placed_on_drawings = slicer_to_drawings
            .get(&part_name)
            .map(|drawings| drawings.iter().cloned().collect::<Vec<_>>())
            .unwrap_or_default();

        let (placed_on_sheets, placed_on_sheet_names) = placement_sheet_info(
            &placed_on_drawings,
            &drawing_to_sheets,
            &sheet_name_by_part,
        );

        let selection = cache_part
            .as_deref()
            .and_then(|cache_part| package.part(cache_part))
            .and_then(|bytes| parse_slicer_cache_selection(bytes).ok())
            .unwrap_or_default();

        slicers.push(SlicerDefinition {
            part_name: part_name.clone(),
            name: parsed.name,
            uid: parsed.uid,
            cache_part,
            cache_name,
            source_name,
            connected_pivot_tables,
            connected_tables,
            placed_on_drawings,
            placed_on_sheets,
            placed_on_sheet_names,
            selection,
        });
    }

    let mut timelines = Vec::with_capacity(timeline_parts.len());
    for part_name in timeline_parts {
        let xml = package
            .part(&part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.clone()))?;
        let parsed = parse_timeline_xml(xml)?;

        let cache_part = match parsed.cache_rid.as_deref() {
            Some(rid) => resolve_relationship_target(package, &part_name, rid)?,
            None => None,
        };

        let (cache_name, source_name, base_field, level, connected_pivot_tables) =
            if let Some(cache_part) = cache_part.as_deref() {
                let resolved = resolve_timeline_cache_definition(package, cache_part)?;
                (
                    resolved.cache_name,
                    resolved.source_name,
                    resolved.base_field,
                    resolved.level,
                    resolved.connected_pivot_tables,
                )
            } else {
                (None, None, None, None, Vec::new())
            };

        let placed_on_drawings = timeline_to_drawings
            .get(&part_name)
            .map(|drawings| drawings.iter().cloned().collect::<Vec<_>>())
            .unwrap_or_default();

        let (placed_on_sheets, placed_on_sheet_names) = placement_sheet_info(
            &placed_on_drawings,
            &drawing_to_sheets,
            &sheet_name_by_part,
        );

        let mut selection = parse_timeline_selection(xml, excel_date_system).unwrap_or_default();
        if (selection.start.is_none() || selection.end.is_none()) && cache_part.is_some() {
            if let Some(cache_part) = cache_part.as_deref() {
                if let Some(bytes) = package.part(cache_part) {
                    if let Ok(cache_selection) = parse_timeline_selection(bytes, excel_date_system)
                    {
                        if selection.start.is_none() {
                            selection.start = cache_selection.start;
                        }
                        if selection.end.is_none() {
                            selection.end = cache_selection.end;
                        }
                    }
                }
            }
        }

        timelines.push(TimelineDefinition {
            part_name: part_name.clone(),
            name: parsed.name,
            uid: parsed.uid,
            cache_part,
            cache_name,
            source_name,
            base_field,
            level,
            connected_pivot_tables,
            placed_on_drawings,
            placed_on_sheets,
            placed_on_sheet_names,
            selection,
        });
    }

    Ok(PivotSlicerParts { slicers, timelines })
}

fn worksheet_part_name_from_rels(rels_name: &str) -> String {
    // Example: xl/worksheets/_rels/sheet1.xml.rels -> xl/worksheets/sheet1.xml
    let trimmed = rels_name
        .strip_prefix("xl/worksheets/_rels/")
        .unwrap_or(rels_name);
    let trimmed = trimmed.strip_suffix(".rels").unwrap_or(trimmed);
    format!("xl/worksheets/{trimmed}")
}

fn chartsheet_part_name_from_rels(rels_name: &str) -> String {
    // Example: xl/chartsheets/_rels/sheet1.xml.rels -> xl/chartsheets/sheet1.xml
    let trimmed = rels_name
        .strip_prefix("xl/chartsheets/_rels/")
        .unwrap_or(rels_name);
    let trimmed = trimmed.strip_suffix(".rels").unwrap_or(trimmed);
    format!("xl/chartsheets/{trimmed}")
}

fn is_drawing_relationship_type(type_uri: &str) -> bool {
    // Most producers use the canonical OfficeDocument relationship URI, but some third-party
    // tools may vary the prefix. Since we only need to locate drawing parts, match by suffix.
    type_uri.trim_end().ends_with("/drawing")
}

fn sheet_name_by_part(package: &XlsxPackage) -> BTreeMap<String, String> {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = match package.part(workbook_part) {
        Some(bytes) => bytes,
        None => return BTreeMap::new(),
    };
    let workbook_xml = match String::from_utf8(workbook_xml.to_vec()) {
        Ok(xml) => xml,
        Err(_) => return BTreeMap::new(),
    };
    let sheets = match parse_workbook_sheets(&workbook_xml) {
        Ok(sheets) => sheets,
        Err(_) => return BTreeMap::new(),
    };

    let mut out = BTreeMap::new();
    for sheet in sheets {
        let resolved = resolve_relationship_target(package, workbook_part, &sheet.rel_id)
            .ok()
            .flatten()
            .or_else(|| {
                let guess_ws = format!("xl/worksheets/sheet{}.xml", sheet.sheet_id);
                if package.part(&guess_ws).is_some() {
                    return Some(guess_ws);
                }
                let guess_cs = format!("xl/chartsheets/sheet{}.xml", sheet.sheet_id);
                package.part(&guess_cs).map(|_| guess_cs)
            });
        if let Some(part) = resolved {
            out.insert(part, sheet.name);
        }
    }

    out
}

fn placement_sheet_info(
    placed_on_drawings: &[String],
    drawing_to_sheets: &BTreeMap<String, BTreeSet<String>>,
    sheet_name_by_part: &BTreeMap<String, String>,
) -> (Vec<String>, Vec<String>) {
    let mut placed_on_sheets: BTreeSet<String> = BTreeSet::new();
    for drawing in placed_on_drawings {
        if let Some(sheets) = drawing_to_sheets.get(drawing) {
            placed_on_sheets.extend(sheets.iter().cloned());
        }
    }
    let placed_on_sheets = placed_on_sheets.into_iter().collect::<Vec<_>>();

    let mut placed_on_sheet_names: BTreeSet<String> = BTreeSet::new();
    for sheet_part in &placed_on_sheets {
        if let Some(name) = sheet_name_by_part.get(sheet_part) {
            placed_on_sheet_names.insert(name.clone());
        }
    }

    (
        placed_on_sheets,
        placed_on_sheet_names.into_iter().collect::<Vec<_>>(),
    )
}
fn parse_iso_ymd(value: &str) -> Option<NaiveDate> {
    let trimmed = value.trim();
    let ymd = trimmed.get(..10).unwrap_or(trimmed);
    NaiveDate::parse_from_str(ymd, "%Y-%m-%d").ok()
}

fn parse_excel_bool(value: &str) -> Option<bool> {
    let trimmed = value.trim();
    if trimmed == "1" {
        return Some(true);
    }
    if trimmed == "0" {
        return Some(false);
    }
    if trimmed.eq_ignore_ascii_case("true") {
        return Some(true);
    }
    if trimmed.eq_ignore_ascii_case("yes") {
        return Some(true);
    }
    if trimmed.eq_ignore_ascii_case("false") {
        return Some(false);
    }
    if trimmed.eq_ignore_ascii_case("no") {
        return Some(false);
    }
    None
}

fn parse_slicer_cache_selection(xml: &[u8]) -> Result<SlicerSelectionState, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut inner_buf = Vec::new();

    let mut available_items = Vec::new();
    let mut seen_items: HashSet<String> = HashSet::new();
    let mut selected_items: HashSet<String> = HashSet::new();
    let mut saw_selection_attr = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"slicerCacheItem") {
                    // Some third-party generators nest the item key as text (often within a `<t>`
                    // element) rather than persisting a `name`/`n` attribute. Consume the element
                    // and treat the first nested text node as the key if no usable attribute is
                    // present.
                    let nested_text = read_slicer_cache_item_text(&mut reader, &mut inner_buf)?;
                    let (key, selected, saw_attr) =
                        parse_slicer_cache_item(&start, nested_text.as_deref())?;
                    if key.is_empty() {
                        continue;
                    }
                    if saw_attr {
                        saw_selection_attr = true;
                    }
                    if seen_items.insert(key.clone()) {
                        available_items.push(key.clone());
                    }
                    if selected {
                        selected_items.insert(key);
                    }
                }
            }
            Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"slicerCacheItem") {
                    let (key, selected, saw_attr) = parse_slicer_cache_item(&start, None)?;
                    if key.is_empty() {
                        continue;
                    }
                    if saw_attr {
                        saw_selection_attr = true;
                    }
                    if seen_items.insert(key.clone()) {
                        available_items.push(key.clone());
                    }
                    if selected {
                        selected_items.insert(key);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    let selected_items = if available_items.is_empty() || !saw_selection_attr {
        None
    } else if selected_items.len() == available_items.len() {
        None
    } else {
        Some(selected_items)
    };

    Ok(SlicerSelectionState {
        available_items,
        selected_items,
    })
}

fn read_slicer_cache_item_text<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
) -> Result<Option<String>, XlsxError> {
    let mut depth = 0u32;
    let mut text = None;

    loop {
        match reader.read_event_into(buf)? {
            Event::Start(start) => {
                if local_name(start.name().as_ref()).eq_ignore_ascii_case(b"slicerCacheItem") {
                    depth = depth.saturating_add(1);
                }
            }
            Event::End(end) => {
                if local_name(end.name().as_ref()).eq_ignore_ascii_case(b"slicerCacheItem") {
                    if depth == 0 {
                        buf.clear();
                        break;
                    }
                    depth -= 1;
                }
            }
            Event::Text(value) => {
                if text.is_none() {
                    let value = value.unescape()?.into_owned();
                    if !value.is_empty() {
                        text = Some(value);
                    }
                }
            }
            Event::CData(value) => {
                if text.is_none() {
                    let value = String::from_utf8_lossy(value.as_ref()).into_owned();
                    if !value.is_empty() {
                        text = Some(value);
                    }
                }
            }
            Event::Eof => {
                buf.clear();
                break;
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(text)
}

fn parse_slicer_cache_item(
    start: &quick_xml::events::BytesStart<'_>,
    fallback_text: Option<&str>,
) -> Result<(String, bool, bool), XlsxError> {
    let mut key_n = None;
    let mut key_name = None;
    let mut key_item_name = None;
    let mut key_caption = None;
    let mut key_unique_name = None;
    let mut key_v = None;
    let mut index_key = None;
    let mut selected = None;
    let mut saw_selection_attr = false;

    for attr in start.attributes().with_checks(false) {
        let attr = attr?;
        let attr_key = local_name(attr.key.as_ref());
        let value = attr.unescape_value()?.into_owned();

        if attr_key.eq_ignore_ascii_case(b"n") && key_n.is_none() && !value.is_empty() {
            key_n = Some(value);
        } else if attr_key.eq_ignore_ascii_case(b"name") && key_name.is_none() && !value.is_empty()
        {
            key_name = Some(value);
        } else if attr_key.eq_ignore_ascii_case(b"itemName")
            && key_item_name.is_none()
            && !value.is_empty()
        {
            key_item_name = Some(value);
        } else if attr_key.eq_ignore_ascii_case(b"caption") && key_caption.is_none() && !value.is_empty()
        {
            key_caption = Some(value);
        } else if attr_key.eq_ignore_ascii_case(b"uniqueName")
            && key_unique_name.is_none()
            && !value.is_empty()
        {
            key_unique_name = Some(value);
        } else if attr_key.eq_ignore_ascii_case(b"v") && key_v.is_none() && !value.is_empty() {
            key_v = Some(value);
        } else if attr_key.eq_ignore_ascii_case(b"x") && !value.is_empty() {
            index_key = Some(value);
        } else if attr_key.eq_ignore_ascii_case(b"s") || attr_key.eq_ignore_ascii_case(b"selected")
        {
            saw_selection_attr = true;
            if selected.is_none() {
                selected = parse_excel_bool(&value);
            }
        }
    }

    let mut key = key_n
        .or(key_name)
        .or(key_item_name)
        .or(key_caption)
        .or(key_unique_name)
        .or(key_v)
        .or(index_key)
        .unwrap_or_default();
    if key.is_empty() {
        if let Some(value) = fallback_text {
            if !value.trim().is_empty() {
                key = value.to_string();
            }
        }
    }
    let selected = selected.unwrap_or(true);

    Ok((key, selected, saw_selection_attr))
}

#[cfg(test)]
mod slicer_cache_selection_tests {
    use super::*;

    use std::collections::HashSet;

    #[test]
    fn parse_slicer_cache_selection_attribute_items() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache>
  <slicerCacheItems>
    <slicerCacheItem n="East" s="1"/>
    <slicerCacheItem n="West" s="0"/>
  </slicerCacheItems>
</slicerCache>"#;

        let selection = parse_slicer_cache_selection(xml).expect("parse selection");
        assert_eq!(selection.available_items, vec!["East", "West"]);

        let selected = selection.selected_items.expect("explicit selection");
        let expected: HashSet<String> = ["East".to_string()].into_iter().collect();
        assert_eq!(selected, expected);
    }

    #[test]
    fn parse_slicer_cache_selection_nested_text_items() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache>
  <slicerCacheItems>
    <slicerCacheItem s="1"><t>East</t></slicerCacheItem>
    <slicerCacheItem s="0"><t>West</t></slicerCacheItem>
  </slicerCacheItems>
</slicerCache>"#;

        let selection = parse_slicer_cache_selection(xml).expect("parse selection");
        assert_eq!(selection.available_items, vec!["East", "West"]);

        let selected = selection.selected_items.expect("explicit selection");
        let expected: HashSet<String> = ["East".to_string()].into_iter().collect();
        assert_eq!(selected, expected);
    }

    #[test]
    fn parse_slicer_cache_selection_dedupes_keys() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache>
  <slicerCacheItems>
    <slicerCacheItem n="East" s="1"/>
    <slicerCacheItem n="East" s="1"/>
    <slicerCacheItem n="West" s="0"/>
    <slicerCacheItem n="West" s="0"/>
  </slicerCacheItems>
</slicerCache>"#;

        let selection = parse_slicer_cache_selection(xml).expect("parse selection");
        assert_eq!(selection.available_items, vec!["East", "West"]);

        let selected = selection.selected_items.expect("explicit selection");
        let expected: HashSet<String> = ["East".to_string()].into_iter().collect();
        assert_eq!(selected, expected);
    }

    #[test]
    fn parse_slicer_cache_selection_item_name_attribute() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache>
  <slicerCacheItems>
    <slicerCacheItem itemName="East" s="1"/>
    <slicerCacheItem itemName="West" s="0"/>
  </slicerCacheItems>
</slicerCache>"#;

        let selection = parse_slicer_cache_selection(xml).expect("parse selection");
        assert_eq!(selection.available_items, vec!["East", "West"]);

        let selected = selection.selected_items.expect("explicit selection");
        let expected: HashSet<String> = ["East".to_string()].into_iter().collect();
        assert_eq!(selected, expected);
    }

    #[test]
    fn parse_slicer_cache_selection_prefers_n_over_caption_regardless_of_attr_order() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache>
  <slicerCacheItems>
    <slicerCacheItem caption="CaptionFirst" n="NameSecond" s="1"/>
    <slicerCacheItem caption="OtherCaption" n="OtherName" s="0"/>
  </slicerCacheItems>
</slicerCache>"#;

        let selection = parse_slicer_cache_selection(xml).expect("parse selection");
        assert_eq!(selection.available_items, vec!["NameSecond", "OtherName"]);

        let selected = selection.selected_items.expect("explicit selection");
        let expected: HashSet<String> = ["NameSecond".to_string()].into_iter().collect();
        assert_eq!(selected, expected);
    }
}

fn parse_timeline_selection(
    xml: &[u8],
    date_system: ExcelDateSystem,
) -> Result<TimelineSelectionState, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut selection = TimelineSelectionState::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                let is_start_el =
                    tag.eq_ignore_ascii_case(b"start") || tag.eq_ignore_ascii_case(b"startDate");
                let is_end_el =
                    tag.eq_ignore_ascii_case(b"end") || tag.eq_ignore_ascii_case(b"endDate");

                for attr in start.attributes().with_checks(false) {
                    let attr = attr?;
                    let key = local_name(attr.key.as_ref());
                    let value = attr.unescape_value()?.into_owned();

                    if selection.start.is_none()
                        && (key.eq_ignore_ascii_case(b"start")
                            || key.eq_ignore_ascii_case(b"startDate")
                            || key.eq_ignore_ascii_case(b"selectionStart")
                            || key.eq_ignore_ascii_case(b"selectionStartDate")
                            || (is_start_el
                                && (key.eq_ignore_ascii_case(b"val")
                                    || key.eq_ignore_ascii_case(b"value"))))
                    {
                        selection.start = normalize_timeline_date(&value, date_system);
                    }

                    if selection.end.is_none()
                        && (key.eq_ignore_ascii_case(b"end")
                            || key.eq_ignore_ascii_case(b"endDate")
                            || key.eq_ignore_ascii_case(b"selectionEnd")
                            || key.eq_ignore_ascii_case(b"selectionEndDate")
                            || (is_end_el
                                && (key.eq_ignore_ascii_case(b"val")
                                    || key.eq_ignore_ascii_case(b"value"))))
                    {
                        selection.end = normalize_timeline_date(&value, date_system);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(selection)
}

fn normalize_timeline_date(value: &str, date_system: ExcelDateSystem) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return None;
    }

    // Common case: already ISO `YYYY-MM-DD` (or `YYYY-MM-DDTHH:MM:SS...`).
    if trimmed.len() >= 10 {
        let prefix = &trimmed[..10];
        let bytes = prefix.as_bytes();
        if bytes.len() == 10
            && bytes[4] == b'-'
            && bytes[7] == b'-'
            && bytes[..4].iter().all(|b| b.is_ascii_digit())
            && bytes[5..7].iter().all(|b| b.is_ascii_digit())
            && bytes[8..10].iter().all(|b| b.is_ascii_digit())
        {
            return Some(prefix.to_string());
        }
    }

    // Another common representation: `YYYYMMDD`.
    if trimmed.len() == 8 && trimmed.as_bytes().iter().all(|b| b.is_ascii_digit()) {
        return Some(format!(
            "{}-{}-{}",
            &trimmed[..4],
            &trimmed[4..6],
            &trimmed[6..8]
        ));
    }

    // Fallback: interpret as an Excel serial date using the workbook date system.
    if let Ok(serial) = trimmed.parse::<i32>() {
        if let Ok(date) = serial_to_ymd(serial, date_system) {
            return Some(format!(
                "{:04}-{:02}-{:02}",
                date.year, date.month, date.day
            ));
        }
    }

    None
}

fn parse_workbook_date_system(xml: &[u8]) -> Result<DateSystem, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut date_system = DateSystem::V1900;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"workbookPr") {
                    for attr in start.attributes().with_checks(false) {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        if key.eq_ignore_ascii_case(b"date1904") {
                            let value = attr.unescape_value()?.into_owned();
                            if parse_excel_bool(&value).unwrap_or(false) {
                                date_system = DateSystem::V1904;
                                break;
                            }
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(date_system)
}

#[derive(Debug)]
struct ResolvedSlicerCacheDefinition {
    cache_name: Option<String>,
    source_name: Option<String>,
    connected_pivot_tables: Vec<String>,
    connected_tables: Vec<String>,
}

fn resolve_slicer_cache_definition(
    package: &XlsxPackage,
    cache_part: &str,
) -> Result<ResolvedSlicerCacheDefinition, XlsxError> {
    let cache_bytes = package
        .part(cache_part)
        .ok_or_else(|| XlsxError::MissingPart(cache_part.to_string()))?;
    let parsed = parse_slicer_cache_xml(cache_bytes)?;

    // Slicer caches can connect to both pivot tables and Excel Tables (ListObjects). Pivot table
    // connections are referenced by relationship id inside the cache XML, while table connections
    // are typically represented solely by relationships of type `.../table`.
    //
    // Excel generally emits `xl/slicerCaches/_rels/slicerCache*.xml.rels`, but the part can be
    // missing or malformed in real-world files. This code is best-effort: if relationships cannot
    // be parsed we fall back to empty connection lists rather than failing slicer discovery.
    let (rel_by_id, connected_tables) = parse_slicer_cache_rels_best_effort(package, cache_part);

    let mut connected_pivot_tables = BTreeSet::new();
    for rid in parsed.pivot_table_rids {
        let Some(rel) = rel_by_id.get(&rid) else {
            continue;
        };
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        let resolved = resolve_target(cache_part, &rel.target);
        if !resolved.is_empty() {
            connected_pivot_tables.insert(resolved);
        }
    }

    Ok(ResolvedSlicerCacheDefinition {
        cache_name: parsed.cache_name,
        source_name: parsed.source_name,
        connected_pivot_tables: connected_pivot_tables.into_iter().collect(),
        connected_tables,
    })
}

fn parse_slicer_cache_rels_best_effort(
    package: &XlsxPackage,
    cache_part: &str,
) -> (HashMap<String, crate::openxml::Relationship>, Vec<String>) {
    const TABLE_REL_TYPE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

    let rels_name = rels_part_name(cache_part);
    let relationships = match package.part(&rels_name) {
        Some(bytes) => parse_relationships(bytes).unwrap_or_else(|_| Vec::new()),
        None => Vec::new(),
    };

    let mut rel_by_id = HashMap::with_capacity(relationships.len());
    let mut connected_tables = BTreeSet::new();

    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            // External relationships are not OPC parts.
            continue;
        }

        let type_uri = rel.type_uri.trim();
        if type_uri == TABLE_REL_TYPE || type_uri.ends_with("/table") {
            let resolved = resolve_target(cache_part, &rel.target);
            if !resolved.is_empty() {
                connected_tables.insert(resolved);
            }
        }

        rel_by_id.insert(rel.id.clone(), rel);
    }

    (rel_by_id, connected_tables.into_iter().collect())
}

#[derive(Debug)]
struct ResolvedTimelineCacheDefinition {
    cache_name: Option<String>,
    source_name: Option<String>,
    base_field: Option<u32>,
    level: Option<u32>,
    connected_pivot_tables: Vec<String>,
}

fn resolve_timeline_cache_definition(
    package: &XlsxPackage,
    cache_part: &str,
) -> Result<ResolvedTimelineCacheDefinition, XlsxError> {
    let cache_bytes = package
        .part(cache_part)
        .ok_or_else(|| XlsxError::MissingPart(cache_part.to_string()))?;
    let parsed = parse_timeline_cache_xml(cache_bytes)?;
    let connected_pivot_tables =
        resolve_relationship_targets(package, cache_part, parsed.pivot_table_rids)?;

    Ok(ResolvedTimelineCacheDefinition {
        cache_name: parsed.cache_name,
        source_name: parsed.source_name,
        base_field: parsed.base_field,
        level: parsed.level,
        connected_pivot_tables,
    })
}

fn resolve_relationship_targets(
    package: &XlsxPackage,
    base_part: &str,
    relationship_ids: Vec<String>,
) -> Result<Vec<String>, XlsxError> {
    let mut targets = BTreeSet::new();
    for rid in relationship_ids {
        if let Some(target) = resolve_relationship_target(package, base_part, &rid)? {
            targets.insert(target);
        }
    }
    Ok(targets.into_iter().collect())
}

fn drawing_part_name_from_rels(rels_name: &str) -> String {
    // Example: xl/drawings/_rels/drawing1.xml.rels -> xl/drawings/drawing1.xml
    let trimmed = rels_name
        .strip_prefix("xl/drawings/_rels/")
        .unwrap_or(rels_name);
    let trimmed = trimmed.strip_suffix(".rels").unwrap_or(trimmed);
    format!("xl/drawings/{trimmed}")
}

#[derive(Debug)]
struct ParsedSlicerXml {
    name: Option<String>,
    uid: Option<String>,
    cache_rid: Option<String>,
}

fn parse_slicer_xml(xml: &[u8]) -> Result<ParsedSlicerXml, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut name = None;
    let mut uid = None;
    let mut cache_rid = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"slicer") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"name") {
                            name = Some(value);
                        } else if key.eq_ignore_ascii_case(b"uid") {
                            uid = Some(value);
                        }
                    }
                } else if tag.eq_ignore_ascii_case(b"slicerCache") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            cache_rid = Some(attr.unescape_value()?.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(ParsedSlicerXml {
        name,
        uid,
        cache_rid,
    })
}

#[derive(Debug)]
struct ParsedSlicerCacheXml {
    cache_name: Option<String>,
    source_name: Option<String>,
    pivot_table_rids: Vec<String>,
}

fn parse_slicer_cache_xml(xml: &[u8]) -> Result<ParsedSlicerCacheXml, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut cache_name = None;
    let mut source_name = None;
    let mut pivot_table_rids = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"slicerCache") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"name") {
                            cache_name = Some(value);
                        } else if key.eq_ignore_ascii_case(b"sourceName") {
                            source_name = Some(value);
                        }
                    }
                } else if tag.eq_ignore_ascii_case(b"slicerCachePivotTable") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            pivot_table_rids.push(attr.unescape_value()?.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(ParsedSlicerCacheXml {
        cache_name,
        source_name,
        pivot_table_rids,
    })
}

#[derive(Debug)]
struct ParsedTimelineXml {
    name: Option<String>,
    uid: Option<String>,
    cache_rid: Option<String>,
}

fn parse_timeline_xml(xml: &[u8]) -> Result<ParsedTimelineXml, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut name = None;
    let mut uid = None;
    let mut cache_rid = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"timeline") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"name") {
                            name = Some(value);
                        } else if key.eq_ignore_ascii_case(b"uid") {
                            uid = Some(value);
                        }
                    }
                } else if tag.eq_ignore_ascii_case(b"timelineCache") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            cache_rid = Some(attr.unescape_value()?.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(ParsedTimelineXml {
        name,
        uid,
        cache_rid,
    })
}

#[derive(Debug)]
struct ParsedTimelineCacheXml {
    cache_name: Option<String>,
    source_name: Option<String>,
    base_field: Option<u32>,
    level: Option<u32>,
    pivot_table_rids: Vec<String>,
}

fn parse_timeline_cache_xml(xml: &[u8]) -> Result<ParsedTimelineCacheXml, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut cache_name = None;
    let mut source_name = None;
    let mut base_field = None;
    let mut level = None;
    let mut pivot_table_rids = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"timelineCacheDefinition") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"name") {
                            cache_name = Some(value);
                        } else if key.eq_ignore_ascii_case(b"sourceName") {
                            source_name = Some(value);
                        } else if key.eq_ignore_ascii_case(b"baseField") {
                            base_field = value.parse::<u32>().ok();
                        } else if key.eq_ignore_ascii_case(b"level") {
                            level = value.parse::<u32>().ok();
                        }
                    }
                } else if tag.eq_ignore_ascii_case(b"pivotTable") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            pivot_table_rids.push(attr.unescape_value()?.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(ParsedTimelineCacheXml {
        cache_name,
        source_name,
        base_field,
        level,
        pivot_table_rids,
    })
}

#[cfg(test)]
mod engine_filter_field_tests {
    use super::*;
    use std::collections::HashSet;

    #[test]
    fn slicer_selection_to_engine_filter_field_items() {
        let mut selected_items = HashSet::new();
        selected_items.insert("East".to_string());
        let selection = SlicerSelectionState {
            available_items: Vec::new(),
            selected_items: Some(selected_items),
        };

        let actual =
            slicer_selection_to_engine_filter_field_with_resolver("Region", &selection, |_| None);

        let mut expected_allowed = HashSet::new();
        expected_allowed.insert(formula_engine::pivot::PivotKeyPart::Text("East".to_string()));
        let expected = formula_engine::pivot::FilterField {
            source_field: "Region".to_string(),
            allowed: Some(expected_allowed),
        };

        assert_eq!(actual, expected);
    }

    #[test]
    fn slicer_selection_to_engine_filter_field_all() {
        let selection = SlicerSelectionState {
            available_items: Vec::new(),
            selected_items: None,
        };

        let actual =
            slicer_selection_to_engine_filter_field_with_resolver("Region", &selection, |_| None);

        let expected = formula_engine::pivot::FilterField {
            source_field: "Region".to_string(),
            allowed: None,
        };

        assert_eq!(actual, expected);
    }
}
