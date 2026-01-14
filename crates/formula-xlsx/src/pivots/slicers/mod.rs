use crate::openxml::{
    local_name, parse_relationships, rels_part_name, resolve_relationship_target, resolve_target,
};
use crate::package::{XlsxError, XlsxPackage};
use crate::sheet_metadata::parse_workbook_sheets;
use crate::{DateSystem, XlsxDocument};
use super::{PivotCacheDefinition, PivotCacheValue};
use formula_engine::pivot::{FilterField, PivotFieldRef, PivotKeyPart};
use chrono::{Datelike, NaiveDate};
use formula_engine::date::{serial_to_ymd, ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_model::pivots::slicers::{RowFilter, SlicerSelection, TimelineSelection};
use formula_model::pivots::ScalarValue;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::Reader;
use quick_xml::Writer;
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

/// Create a resolver that interprets slicer item keys as indices (`x="..."`) into the pivot
/// cache field's `<sharedItems>` table.
///
/// This is useful when `SlicerSelectionState` keys are numeric strings (e.g. `"0"`), which
/// Excel uses for slicer caches bound to pivot caches that store items as shared-item indices.
pub fn shared_item_key_resolver(
    cache_def: &PivotCacheDefinition,
    field_idx: usize,
) -> impl FnMut(&str) -> Option<ScalarValue> + '_ {
    move |key| resolve_slicer_item_key(cache_def, field_idx, key)
}

/// Resolve a slicer item key into a typed [`ScalarValue`] using a pivot cache field's
/// `<sharedItems>` table.
///
/// Returns `None` when `key` is not numeric, or when the shared item cannot be found.
pub fn resolve_slicer_item_key(
    cache_def: &PivotCacheDefinition,
    field_idx: usize,
    key: &str,
) -> Option<ScalarValue> {
    let idx = key.parse::<u32>().ok()? as usize;
    let field = cache_def.cache_fields.get(field_idx)?;
    let item = field.shared_items.as_ref()?.get(idx)?;

    match item {
        PivotCacheValue::String(s) => Some(ScalarValue::Text(s.clone())),
        PivotCacheValue::Number(n) => Some(ScalarValue::from(*n)),
        PivotCacheValue::Bool(b) => Some(ScalarValue::Bool(*b)),
        PivotCacheValue::DateTime(s) => {
            if let Some(date) = parse_iso_ymd(s) {
                Some(ScalarValue::Date(date))
            } else {
                Some(ScalarValue::Text(s.clone()))
            }
        }
        PivotCacheValue::Missing | PivotCacheValue::Error(_) => Some(ScalarValue::Blank),
        PivotCacheValue::Index(_) => None,
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
    let field = field.into();
    let allowed = match &selection.selected_items {
        None => None,
        Some(items) => {
            let mut selected = HashSet::with_capacity(items.len());
            for item in items {
                selected
                    .insert(resolve(item).unwrap_or_else(|| {
                        formula_engine::pivot::PivotKeyPart::Text(item.clone())
                    }));
            }
            Some(selected)
        }
    };

    formula_engine::pivot::FilterField {
        source_field: PivotFieldRef::CacheFieldName(field),
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
    let field = field.into();
    formula_engine::pivot::FilterField {
        source_field: PivotFieldRef::CacheFieldName(field),
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

/// Convert a parsed timeline selection into a pivot-engine filter field by enumerating the
/// allowed date items in the pivot cache definition.
///
/// This relies on the cache field's `<sharedItems>` list being available. If shared items are
/// missing (or the timeline selection has no effective bounds), the returned [`FilterField`]
/// will have `allowed: None`, meaning "allow all".
pub fn timeline_selection_to_engine_filter_field_with_cache(
    cache_def: &PivotCacheDefinition,
    base_field: u32,
    selection: &TimelineSelectionState,
) -> FilterField {
    let source_field = PivotFieldRef::CacheFieldName(
        cache_def
            .cache_fields
            .get(base_field as usize)
            .map(|f| f.name.clone())
            .unwrap_or_default(),
    );

    if selection.start.is_none() && selection.end.is_none() {
        return FilterField {
            source_field,
            allowed: None,
        };
    }

    let start = selection.start.as_deref().and_then(parse_iso_ymd);
    let end = selection.end.as_deref().and_then(parse_iso_ymd);

    // If the persisted endpoints are not parseable, avoid applying a potentially destructive
    // empty filter.
    if start.is_none() && end.is_none() {
        return FilterField {
            source_field,
            allowed: None,
        };
    }

    let Some(field) = cache_def.cache_fields.get(base_field as usize) else {
        return FilterField {
            source_field,
            allowed: None,
        };
    };
    let Some(shared_items) = field.shared_items.as_ref() else {
        return FilterField {
            source_field,
            allowed: None,
        };
    };

    let date_system = infer_date_system_from_timeline_shared_items(shared_items, start, end);

    let mut allowed = HashSet::new();
    for item in shared_items {
        let Some(date) = pivot_cache_shared_item_to_naive_date(item, date_system) else {
            continue;
        };
        if let Some(start) = start {
            if date < start {
                continue;
            }
        }
        if let Some(end) = end {
            if date > end {
                continue;
            }
        }
        allowed.insert(PivotKeyPart::Date(date));
    }

    FilterField {
        source_field,
        allowed: Some(allowed),
    }
}

fn pivot_cache_shared_item_to_naive_date(
    item: &PivotCacheValue,
    date_system: ExcelDateSystem,
) -> Option<NaiveDate> {
    match item {
        PivotCacheValue::DateTime(v) | PivotCacheValue::String(v) => {
            // Timeline caches normalize dates using the workbook's date system, but this helper is
            // purely cache-definition based and does not have access to the workbook metadata.
            // We still accept a `date_system` parameter to support best-effort inference from the
            // selection range when shared items are numeric serials.
            normalize_timeline_date(v, date_system).as_deref().and_then(parse_iso_ymd)
        }
        PivotCacheValue::Number(n) => {
            // Excel stores dates as serial numbers; shared items for date fields should generally
            // use `<d>`, but accept numbers as a best-effort fallback.
            let serial = n.floor() as i32;
            serial_to_ymd(serial, date_system)
                .ok()
                .and_then(|ymd| NaiveDate::from_ymd_opt(ymd.year, ymd.month as u32, ymd.day as u32))
        }
        _ => None,
    }
}

fn date_in_range(date: NaiveDate, start: Option<NaiveDate>, end: Option<NaiveDate>) -> bool {
    if let Some(start) = start {
        if date < start {
            return false;
        }
    }
    if let Some(end) = end {
        if date > end {
            return false;
        }
    }
    true
}

fn infer_date_system_from_timeline_shared_items(
    shared_items: &[PivotCacheValue],
    start: Option<NaiveDate>,
    end: Option<NaiveDate>,
) -> ExcelDateSystem {
    // Without at least one parseable bound we cannot infer which date system the cached serials
    // use (and callers explicitly treat that scenario as "allow all").
    if start.is_none() && end.is_none() {
        return ExcelDateSystem::EXCEL_1900;
    }

    let mut matches_1900 = 0usize;
    let mut matches_1904 = 0usize;
    let mut saw_disambiguating_value = false;

    for item in shared_items {
        let d1900 = pivot_cache_shared_item_to_naive_date(item, ExcelDateSystem::EXCEL_1900);
        let d1904 = pivot_cache_shared_item_to_naive_date(item, ExcelDateSystem::Excel1904);

        if d1900 == d1904 {
            continue;
        }
        saw_disambiguating_value = true;

        if d1900.is_some_and(|d| date_in_range(d, start, end)) {
            matches_1900 += 1;
        }
        if d1904.is_some_and(|d| date_in_range(d, start, end)) {
            matches_1904 += 1;
        }
    }

    if saw_disambiguating_value && matches_1904 > matches_1900 {
        ExcelDateSystem::Excel1904
    } else {
        ExcelDateSystem::EXCEL_1900
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
    /// Best-effort resolved pivot-cache field/column name this slicer filters.
    ///
    /// Excel often only persists a `sourceName` (frequently the pivot table name) and does not
    /// explicitly store the filtered field. When available, we map slicer/timeline metadata back to
    /// the connected pivot cache definition to recover the underlying field name.
    pub field_name: Option<String>,
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
    /// Best-effort resolved pivot-cache field/column name this timeline filters.
    pub field_name: Option<String>,
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
        parse_pivot_slicer_parts_with(
            self.part_names(),
            |name| self.part(name),
            |base, rid| crate::openxml::resolve_relationship_target(self, base, rid),
        )
    }

    /// Update the persisted selection state for a slicer cache (and optionally its slicer
    /// definition part).
    ///
    /// Slicer selections are persisted on the `xl/slicerCaches/slicerCache*.xml` items via
    /// per-`<slicerCacheItem>` selection flags. This helper patches the provided part and, when
    /// given a slicer definition part (`xl/slicers/slicer*.xml`), also attempts to patch the
    /// referenced cache part.
    ///
    /// Relationship traversal is best-effort: missing or malformed `.rels` parts will not prevent
    /// patching the explicitly provided part (when it exists).
    pub fn set_slicer_selection(
        &mut self,
        slicer_cache_or_slicer_part: &str,
        selection: &SlicerSelectionState,
    ) -> Result<(), XlsxError> {
        let canonical = slicer_cache_or_slicer_part
            .trim()
            .trim_start_matches('/')
            .to_string();

        let mut explicit_targets: BTreeSet<String> = BTreeSet::new();
        if !canonical.is_empty() {
            explicit_targets.insert(canonical.clone());
        }

        // Only attempt relationship traversal when the caller passes a slicer definition part.
        let mut inferred_cache_targets: BTreeSet<String> = BTreeSet::new();
        if canonical.starts_with("xl/slicers/") && canonical.ends_with(".xml") {
            if let Some(xml) = self.part(&canonical) {
                if let Ok(parsed) = parse_slicer_xml(xml) {
                    if let Some(rid) = parsed.cache_rid.as_deref() {
                        match resolve_relationship_target(self, &canonical, rid) {
                            Ok(Some(cache_part)) => {
                                inferred_cache_targets.insert(cache_part);
                            }
                            // Best-effort: ignore missing or malformed relationship parts.
                            _ => {}
                        }
                    }
                }
            }
        }

        // Patch explicit parts (erroring when missing).
        for target in explicit_targets {
            let bytes = self
                .part(&target)
                .ok_or_else(|| XlsxError::MissingPart(target.clone()))?
                .to_vec();
            let updated = patch_slicer_selection_xml(&bytes, selection)?;
            self.set_part(target, updated);
        }

        // Patch inferred cache parts (best-effort: skip missing).
        for target in inferred_cache_targets {
            let Some(bytes) = self.part(&target) else {
                continue;
            };
            let updated = patch_slicer_selection_xml(bytes, selection)?;
            self.set_part(target, updated);
        }

        Ok(())
    }

    /// Update the persisted selection state for a timeline cache (and any connected timeline
    /// definition parts).
    ///
    /// Excel persists timeline selections in a few different ways depending on workbook
    /// version and producer. This helper patches the provided part (typically a
    /// `xl/timelineCaches/timelineCacheDefinition*.xml`) and also attempts to update any
    /// `xl/timelines/timeline*.xml` parts that reference the same cache so the workbook reopens
    /// with the requested selection.
    pub fn set_timeline_selection(
        &mut self,
        timeline_cache_part: &str,
        selection: &TimelineSelectionState,
    ) -> Result<(), XlsxError> {
        let canonical = timeline_cache_part
            .trim()
            .trim_start_matches('/')
            .to_string();
        let date_system = self
            .part("xl/workbook.xml")
            .and_then(|bytes| parse_workbook_date_system(bytes).ok())
            .unwrap_or_default();

        let mut targets: BTreeSet<String> = BTreeSet::new();
        if !canonical.is_empty() {
            targets.insert(canonical.clone());
        }

        if canonical.starts_with("xl/timelineCaches/") {
            // Patch any timeline definition parts that reference this cache.
            for timeline_part in self
                .part_names()
                .filter(|name| name.starts_with("xl/timelines/") && name.ends_with(".xml"))
            {
                let Some(xml) = self.part(timeline_part) else {
                    continue;
                };
                let parsed = parse_timeline_xml(xml).ok();
                let Some(rid) = parsed.and_then(|p| p.cache_rid) else {
                    continue;
                };
                // Best-effort: treat malformed `.rels` parts as unresolved instead of failing the
                // selection update. We still want to patch the explicitly provided cache part.
                let resolved = match resolve_relationship_target(self, timeline_part, &rid) {
                    Ok(resolved) => resolved,
                    Err(_) => None,
                };
                if resolved
                    .as_deref()
                    .is_some_and(|target| target.trim_start_matches('/') == canonical)
                {
                    targets.insert(timeline_part.to_string());
                }
            }
        } else if canonical.starts_with("xl/timelines/") {
            // Patch the referenced cache definition as well (if any).
            if let Some(xml) = self.part(&canonical) {
                if let Ok(parsed) = parse_timeline_xml(xml) {
                    if let Some(rid) = parsed.cache_rid.as_deref() {
                        // Best-effort: a malformed timeline `.rels` part should not prevent us
                        // from patching the explicitly provided timeline part.
                        let cache_part =
                            match resolve_relationship_target(self, &canonical, rid) {
                                Ok(cache_part) => cache_part,
                                Err(_) => None,
                            };
                        if let Some(cache_part) = cache_part {
                            targets.insert(cache_part);
                        }
                    }
                }
            }
        }

        for target in targets {
            let bytes = self
                .part(&target)
                .ok_or_else(|| XlsxError::MissingPart(target.clone()))?
                .to_vec();
            let updated = patch_timeline_selection_xml(&bytes, selection, date_system)?;
            self.set_part(target, updated);
        }

        Ok(())
    }

    /// Update the persisted selection state inside a slicer cache part (`xl/slicerCaches/slicerCache*.xml`).
    ///
    /// Excel persists slicer selection by setting `s="1|0"` (or occasionally `selected="1|0"`)
    /// on each `<slicerCacheItem .../>`.
    ///
    /// - When `selection.selected_items` is `None` ("All selected"), we remove any explicit
    ///   selection attributes so Excel falls back to its default behavior.
    /// - When `Some(set)`, we set the selection attribute to `1` when the item key matches,
    ///   else `0`.
    ///
    /// Matching is tolerant of different key attributes (`n`, `name`, `caption`, `uniqueName`,
    /// `v`, and index keys via `x`).
    pub fn set_slicer_cache_selection(
        &mut self,
        cache_part: &str,
        selection: &SlicerSelectionState,
    ) -> Result<(), XlsxError> {
        let part_key = resolve_part_key(self.parts_map(), cache_part)
            .ok_or_else(|| XlsxError::MissingPart(cache_part.to_string()))?;
        let existing = self
            .parts_map()
            .get(&part_key)
            .ok_or_else(|| XlsxError::MissingPart(cache_part.to_string()))?;
        let updated = slicer_cache_xml_set_selection(existing, selection)?;
        self.parts_map_mut().insert(part_key, updated);
        Ok(())
    }
}

fn resolve_part_key(parts: &BTreeMap<String, Vec<u8>>, name: &str) -> Option<String> {
    if parts.contains_key(name) {
        return Some(name.to_string());
    }
    if let Some(stripped) = name.strip_prefix('/') {
        if parts.contains_key(stripped) {
            return Some(stripped.to_string());
        }
    } else {
        let with_slash = format!("/{name}");
        if parts.contains_key(&with_slash) {
            return Some(with_slash);
        }
    }
    None
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TimelineDateRepr {
    Iso,
    CompactYmd,
    Serial,
}

#[derive(Clone, Debug)]
struct DesiredTimelineDate {
    iso: String,
    date: Option<NaiveDate>,
}

impl DesiredTimelineDate {
    fn parse(value: &str, date_system: DateSystem) -> Option<Self> {
        let trimmed = value.trim();
        if trimmed.is_empty() {
            return None;
        }

        let date = parse_timeline_date_input(trimmed, date_system);
        let iso = if let Some(date) = date {
            format!("{:04}-{:02}-{:02}", date.year(), date.month(), date.day())
        } else if let Some(prefix) = iso_prefix(trimmed) {
            prefix.to_string()
        } else {
            trimmed.to_string()
        };

        Some(Self { iso, date })
    }
}

fn iso_prefix(value: &str) -> Option<&str> {
    let trimmed = value.trim();
    if trimmed.len() < 10 {
        return None;
    }
    let prefix = trimmed.get(..10)?;
    let bytes = prefix.as_bytes();
    if bytes.len() == 10
        && bytes[4] == b'-'
        && bytes[7] == b'-'
        && bytes[..4].iter().all(|b| b.is_ascii_digit())
        && bytes[5..7].iter().all(|b| b.is_ascii_digit())
        && bytes[8..10].iter().all(|b| b.is_ascii_digit())
    {
        Some(prefix)
    } else {
        None
    }
}

fn parse_timeline_date_input(value: &str, date_system: DateSystem) -> Option<NaiveDate> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return None;
    }

    // ISO `YYYY-MM-DD` (or a datetime string that starts with it).
    if let Some(prefix) = iso_prefix(trimmed) {
        return NaiveDate::parse_from_str(prefix, "%Y-%m-%d").ok();
    }

    // Compact `YYYYMMDD`.
    if trimmed.len() == 8 && trimmed.as_bytes().iter().all(|b| b.is_ascii_digit()) {
        let iso = format!("{}-{}-{}", &trimmed[..4], &trimmed[4..6], &trimmed[6..8]);
        return NaiveDate::parse_from_str(&iso, "%Y-%m-%d").ok();
    }

    // Excel serial.
    let Ok(serial) = trimmed.parse::<i32>() else {
        return None;
    };
    let Ok(date) = serial_to_ymd(serial, date_system.to_engine_date_system()) else {
        return None;
    };
    NaiveDate::from_ymd_opt(date.year, u32::from(date.month), u32::from(date.day))
}

fn detect_timeline_date_repr(existing: &str) -> TimelineDateRepr {
    let trimmed = existing.trim();
    if iso_prefix(trimmed).is_some() {
        return TimelineDateRepr::Iso;
    }
    if trimmed.len() == 8 && trimmed.as_bytes().iter().all(|b| b.is_ascii_digit()) {
        return TimelineDateRepr::CompactYmd;
    }
    if !trimmed.is_empty() && trimmed.as_bytes().iter().all(|b| b.is_ascii_digit()) {
        return TimelineDateRepr::Serial;
    }
    TimelineDateRepr::Iso
}

fn apply_desired_date_to_existing(
    existing: &str,
    desired: &DesiredTimelineDate,
    date_system: DateSystem,
) -> String {
    let trimmed = existing.trim();
    match detect_timeline_date_repr(trimmed) {
        TimelineDateRepr::Iso => {
            if trimmed.len() > 10 && iso_prefix(trimmed).is_some() {
                format!("{}{}", desired.iso, &trimmed[10..])
            } else {
                desired.iso.clone()
            }
        }
        TimelineDateRepr::CompactYmd => {
            if desired.iso.len() == 10 {
                desired.iso.replace('-', "")
            } else {
                desired.iso.clone()
            }
        }
        TimelineDateRepr::Serial => match desired.date {
            Some(date) => {
                let excel_date = ExcelDate::new(date.year(), date.month() as u8, date.day() as u8);
                match ymd_to_serial(excel_date, date_system.to_engine_date_system()) {
                    Ok(serial) => serial.to_string(),
                    Err(_) => desired.iso.clone(),
                }
            }
            None => desired.iso.clone(),
        },
    }
}

fn patch_timeline_selection_xml(
    xml: &[u8],
    selection: &TimelineSelectionState,
    date_system: DateSystem,
) -> Result<Vec<u8>, XlsxError> {
    if selection.start.is_none() && selection.end.is_none() {
        return Ok(xml.to_vec());
    }

    let desired_start = selection
        .start
        .as_deref()
        .and_then(|s| DesiredTimelineDate::parse(s, date_system));
    let desired_end = selection
        .end
        .as_deref()
        .and_then(|s| DesiredTimelineDate::parse(s, date_system));

    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + 128));
    let mut buf = Vec::new();

    let mut depth = 0usize;
    let mut root_name: Option<String> = None;
    let mut root_prefix: Option<String> = None;
    let mut wrote_start = false;
    let mut wrote_end = false;
    let mut inserted = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e) => {
                if depth == 0 {
                    let name = e.name();
                    let name = name.as_ref();
                    root_prefix = name
                        .iter()
                        .position(|b| *b == b':')
                        .map(|idx| String::from_utf8_lossy(&name[..idx]).into_owned());
                    root_name = Some(String::from_utf8_lossy(name).into_owned());
                }

                let (patched, did_touch, event_start, event_end) =
                    patch_timeline_selection_start(e, &desired_start, &desired_end, date_system)?;
                wrote_start |= event_start;
                wrote_end |= event_end;
                if did_touch {
                    writer.write_event(Event::Start(patched))?;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
                depth += 1;
            }
            Event::Empty(ref e) => {
                if depth == 0 {
                    let name = e.name();
                    let name = name.as_ref();
                    root_prefix = name
                        .iter()
                        .position(|b| *b == b':')
                        .map(|idx| String::from_utf8_lossy(&name[..idx]).into_owned());
                    let root_name_str = String::from_utf8_lossy(name).into_owned();
                    root_name = Some(root_name_str.clone());

                    let (patched, did_touch, event_start, event_end) =
                        patch_timeline_selection_start(
                            e,
                            &desired_start,
                            &desired_end,
                            date_system,
                        )?;
                    wrote_start |= event_start;
                    wrote_end |= event_end;
                    let needs_insert =
                        should_insert_timeline_selection(&desired_start, &desired_end)
                            && ((desired_start.is_some() && !wrote_start)
                                || (desired_end.is_some() && !wrote_end));

                    if needs_insert {
                        inserted = true;
                        wrote_start |= desired_start.is_some();
                        wrote_end |= desired_end.is_some();
                        // Expand the self-closing root so we can insert the selection element.
                        writer.write_event(Event::Start(patched.into_owned()))?;
                        insert_timeline_selection_element(
                            &mut writer,
                            root_prefix.as_deref(),
                            &desired_start,
                            &desired_end,
                        )?;
                        writer.write_event(Event::End(BytesEnd::new(root_name_str.as_str())))?;
                    } else if did_touch {
                        writer.write_event(Event::Empty(patched))?;
                    } else {
                        writer.write_event(Event::Empty(e.to_owned()))?;
                    }
                } else {
                    let (patched, did_touch, event_start, event_end) =
                        patch_timeline_selection_start(
                            e,
                            &desired_start,
                            &desired_end,
                            date_system,
                        )?;
                    wrote_start |= event_start;
                    wrote_end |= event_end;
                    if did_touch {
                        writer.write_event(Event::Empty(patched))?;
                    } else {
                        writer.write_event(Event::Empty(e.to_owned()))?;
                    }
                }
            }
            Event::End(ref e) => {
                if depth == 1
                    && !inserted
                    && should_insert_timeline_selection(&desired_start, &desired_end)
                    && ((desired_start.is_some() && !wrote_start)
                        || (desired_end.is_some() && !wrote_end))
                {
                    inserted = true;
                    insert_timeline_selection_element(
                        &mut writer,
                        root_prefix.as_deref(),
                        &desired_start,
                        &desired_end,
                    )?;
                    wrote_start |= desired_start.is_some();
                    wrote_end |= desired_end.is_some();
                }

                writer.write_event(Event::End(e.to_owned()))?;
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            other => writer.write_event(other.into_owned())?,
        }

        buf.clear();
    }

    // If we never observed any root element (malformed XML), fall back to returning the original.
    if root_name.is_none() {
        return Ok(xml.to_vec());
    }

    Ok(writer.into_inner())
}

fn should_insert_timeline_selection(
    desired_start: &Option<DesiredTimelineDate>,
    desired_end: &Option<DesiredTimelineDate>,
) -> bool {
    desired_start.is_some() || desired_end.is_some()
}

fn insert_timeline_selection_element<W: std::io::Write>(
    writer: &mut Writer<W>,
    prefix: Option<&str>,
    desired_start: &Option<DesiredTimelineDate>,
    desired_end: &Option<DesiredTimelineDate>,
) -> Result<(), XlsxError> {
    let tag_name = match prefix {
        Some(prefix) => format!("{prefix}:selection"),
        None => "selection".to_string(),
    };
    let mut el = BytesStart::new(tag_name.as_str());
    if let Some(start) = desired_start.as_ref() {
        el.push_attribute(("startDate", start.iso.as_str()));
    }
    if let Some(end) = desired_end.as_ref() {
        el.push_attribute(("endDate", end.iso.as_str()));
    }
    writer.write_event(Event::Empty(el))?;
    Ok(())
}

fn patch_timeline_selection_start(
    e: &BytesStart<'_>,
    desired_start: &Option<DesiredTimelineDate>,
    desired_end: &Option<DesiredTimelineDate>,
    date_system: DateSystem,
) -> Result<(BytesStart<'static>, bool, bool, bool), XlsxError> {
    let name = e.name();
    let name = name.as_ref();
    let tag_local = local_name(name);
    let is_selection = tag_local.eq_ignore_ascii_case(b"selection");
    let is_start_el =
        tag_local.eq_ignore_ascii_case(b"start") || tag_local.eq_ignore_ascii_case(b"startDate");
    let is_end_el =
        tag_local.eq_ignore_ascii_case(b"end") || tag_local.eq_ignore_ascii_case(b"endDate");

    let mut touched = false;
    let mut saw_start = false;
    let mut saw_end = false;
    let mut wrote_start = false;
    let mut wrote_end = false;

    let tag_name = std::str::from_utf8(name).unwrap_or("selection");
    let mut patched = BytesStart::new(tag_name);

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key_local = local_name(attr.key.as_ref());
        let value = attr.unescape_value()?.into_owned();
        let mut out_value: Option<String> = None;

        if let Some(desired) = desired_start.as_ref() {
            if is_start_attr_key(key_local) {
                saw_start = true;
                touched = true;
                wrote_start = true;
                out_value = Some(apply_desired_date_to_existing(&value, desired, date_system));
            } else if is_start_el
                && (key_local.eq_ignore_ascii_case(b"val")
                    || key_local.eq_ignore_ascii_case(b"value"))
            {
                saw_start = true;
                touched = true;
                wrote_start = true;
                out_value = Some(apply_desired_date_to_existing(&value, desired, date_system));
            }
        }

        if let Some(desired) = desired_end.as_ref() {
            if is_end_attr_key(key_local) {
                saw_end = true;
                touched = true;
                wrote_end = true;
                out_value = Some(apply_desired_date_to_existing(&value, desired, date_system));
            } else if is_end_el
                && (key_local.eq_ignore_ascii_case(b"val")
                    || key_local.eq_ignore_ascii_case(b"value"))
            {
                saw_end = true;
                touched = true;
                wrote_end = true;
                out_value = Some(apply_desired_date_to_existing(&value, desired, date_system));
            }
        }

        if is_start_attr_key(key_local) || is_end_attr_key(key_local) {
            touched = true;
        }

        if let Some(out_value) = out_value {
            patched.push_attribute((attr.key.as_ref(), out_value.as_bytes()));
        } else {
            patched.push_attribute((attr.key.as_ref(), value.as_bytes()));
        }
    }

    if is_selection {
        if let Some(desired) = desired_start.as_ref() {
            if !saw_start {
                patched.push_attribute(("startDate", desired.iso.as_str()));
                touched = true;
                wrote_start = true;
            }
        }
        if let Some(desired) = desired_end.as_ref() {
            if !saw_end {
                patched.push_attribute(("endDate", desired.iso.as_str()));
                touched = true;
                wrote_end = true;
            }
        }
    }

    Ok((patched.into_owned(), touched, wrote_start, wrote_end))
}

fn is_start_attr_key(key: &[u8]) -> bool {
    key.eq_ignore_ascii_case(b"start")
        || key.eq_ignore_ascii_case(b"startDate")
        || key.eq_ignore_ascii_case(b"selectionStart")
        || key.eq_ignore_ascii_case(b"selectionStartDate")
}

fn is_end_attr_key(key: &[u8]) -> bool {
    key.eq_ignore_ascii_case(b"end")
        || key.eq_ignore_ascii_case(b"endDate")
        || key.eq_ignore_ascii_case(b"selectionEnd")
        || key.eq_ignore_ascii_case(b"selectionEndDate")
}

fn canonicalize_part_name_for_discovery(name: &str) -> String {
    name.trim_start_matches(|c| c == '/' || c == '\\')
        .replace('\\', "/")
        .to_ascii_lowercase()
}

impl XlsxDocument {
    /// Parse slicers and timelines out of the preserved parts in an [`XlsxDocument`].
    pub fn pivot_slicer_parts(&self) -> Result<PivotSlicerParts, XlsxError> {
        parse_pivot_slicer_parts_with(
            self.parts().keys(),
            |name| {
                let name = name.strip_prefix('/').unwrap_or(name);
                self.parts().get(name).map(|bytes| bytes.as_slice())
            },
            |base, rid| {
                crate::openxml::resolve_relationship_target_from_parts(
                    |name| {
                        let name = name.strip_prefix('/').unwrap_or(name);
                        self.parts().get(name).map(|bytes| bytes.as_slice())
                    },
                    base,
                    rid,
                )
            },
        )
    }
}

fn parse_pivot_slicer_parts_with<'a, PN, Part, Resolve>(
    part_names: PN,
    part: Part,
    resolve_relationship_target: Resolve,
) -> Result<PivotSlicerParts, XlsxError>
where
    PN: IntoIterator,
    PN::Item: AsRef<str>,
    Part: Fn(&str) -> Option<&'a [u8]>,
    Resolve: Fn(&str, &str) -> Result<Option<String>, XlsxError>,
{
    let date_system = part("xl/workbook.xml")
        .and_then(|bytes| parse_workbook_date_system(bytes).ok())
        .unwrap_or_default();
    let excel_date_system = date_system.to_engine_date_system();
    // Pivot cache resolution is best-effort; slicer/timeline parsing should not fail just because
    // the pivot relationship graph can't be resolved.
    let part_names: Vec<String> = part_names
        .into_iter()
        .map(|name| name.as_ref().strip_prefix('/').unwrap_or(name.as_ref()).to_string())
        .collect();
    let pivot_graph =
        crate::pivots::graph::pivot_graph_with(part_names.iter(), |name| part(name)).ok();

    let mut slicer_parts = Vec::new();
    let mut timeline_parts = Vec::new();
    let mut drawing_rels = Vec::new();
    let mut worksheet_rels = Vec::new();
    let mut chartsheet_rels = Vec::new();

    for name in &part_names {
        let canonical = canonicalize_part_name_for_discovery(name);
        if canonical.starts_with("xl/slicers/") && canonical.ends_with(".xml") {
            slicer_parts.push(canonical);
        } else if canonical.starts_with("xl/timelines/") && canonical.ends_with(".xml") {
            timeline_parts.push(canonical);
        } else if canonical.starts_with("xl/drawings/_rels/") && canonical.ends_with(".rels") {
            drawing_rels.push(canonical);
        } else if canonical.starts_with("xl/worksheets/_rels/") && canonical.ends_with(".rels") {
            worksheet_rels.push(canonical);
        } else if canonical.starts_with("xl/chartsheets/_rels/") && canonical.ends_with(".rels") {
            chartsheet_rels.push(canonical);
        }
    }

    // Ensure deterministic output and avoid duplicate canonical names when producers emit multiple
    // equivalent part names (e.g. different casing or path separators).
    slicer_parts.sort();
    slicer_parts.dedup();
    timeline_parts.sort();
    timeline_parts.dedup();
    drawing_rels.sort();
    drawing_rels.dedup();
    worksheet_rels.sort();
    worksheet_rels.dedup();
    chartsheet_rels.sort();
    chartsheet_rels.dedup();

    let mut slicer_to_drawings: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    let mut timeline_to_drawings: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();

    for rels_name in drawing_rels {
        // Real-world workbooks can contain malformed relationship XML. Pivot slicer discovery is
        // best-effort, so treat any missing/malformed `.rels` part as empty.
        let Some(rels_bytes) = part(&rels_name) else {
            continue;
        };
        let relationships = match parse_relationships(rels_bytes) {
            Ok(relationships) => relationships,
            Err(_) => continue,
        };
        let drawing_part = drawing_part_name_from_rels(&rels_name);
        for rel in relationships {
            let target = canonicalize_part_name_for_discovery(&resolve_target(&drawing_part, &rel.target));
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
        let Some(rels_bytes) = part(&rels_name) else {
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
            let target = canonicalize_part_name_for_discovery(&resolve_target(&sheet_part, &rel.target));
            if target.starts_with("xl/drawings/") {
                drawing_to_sheets
                    .entry(target)
                    .or_default()
                    .insert(sheet_part.clone());
            }
        }
    }

    let sheet_name_by_part = sheet_name_by_part_with(&part, &resolve_relationship_target);

    let mut slicers = Vec::with_capacity(slicer_parts.len());
    for part_name in slicer_parts {
        let xml = part(&part_name).ok_or_else(|| XlsxError::MissingPart(part_name.clone()))?;
        let parsed = parse_slicer_xml(xml)?;

        // Best-effort: malformed `.rels` parts should not prevent slicer discovery.
        let cache_part = match parsed.cache_rid.as_deref() {
            Some(rid) => resolve_relationship_target(&part_name, rid).ok().flatten(),
            None => None,
        };

        let (cache_name, source_name, connected_pivot_tables, connected_tables) =
            if let Some(cache_part) = cache_part.as_deref() {
                match resolve_slicer_cache_definition(&part, cache_part) {
                    Ok(resolved) => (
                        resolved.cache_name,
                        resolved.source_name,
                        resolved.connected_pivot_tables,
                        resolved.connected_tables,
                    ),
                    Err(_) => {
                        // Best-effort: if the slicer cache XML is malformed, still surface the
                        // cache part and any connected Excel Tables referenced via relationships.
                        let (_, connected_tables) =
                            parse_slicer_cache_rels_best_effort(&part, cache_part);
                        (None, None, Vec::new(), connected_tables)
                    }
                }
            } else {
                (None, None, Vec::new(), Vec::new())
            };

        let placed_on_drawings = slicer_to_drawings
            .get(&part_name)
            .map(|drawings| drawings.iter().cloned().collect::<Vec<_>>())
            .unwrap_or_default();

        let (placed_on_sheets, placed_on_sheet_names) =
            placement_sheet_info(&placed_on_drawings, &drawing_to_sheets, &sheet_name_by_part);

        let selection = cache_part
            .as_deref()
            .and_then(|cache_part| part(cache_part))
            .and_then(|bytes| parse_slicer_cache_selection(bytes).ok())
            .unwrap_or_default();

        let field_name = resolve_slicer_field_name(
            &part,
            pivot_graph.as_ref(),
            cache_part.as_deref(),
            &connected_pivot_tables,
            &selection,
        );

        slicers.push(SlicerDefinition {
            part_name: part_name.clone(),
            name: parsed.name,
            uid: parsed.uid,
            cache_part,
            cache_name,
            source_name,
            field_name,
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
        let xml = part(&part_name).ok_or_else(|| XlsxError::MissingPart(part_name.clone()))?;
        let parsed = parse_timeline_xml(xml)?;

        // Best-effort: malformed `.rels` parts should not prevent timeline discovery.
        let cache_part = match parsed.cache_rid.as_deref() {
            Some(rid) => resolve_relationship_target(&part_name, rid).ok().flatten(),
            None => None,
        };

        let (cache_name, source_name, base_field, level, connected_pivot_tables) =
            if let Some(cache_part) = cache_part.as_deref() {
                match resolve_timeline_cache_definition(&part, &resolve_relationship_target, cache_part) {
                    Ok(resolved) => (
                        resolved.cache_name,
                        resolved.source_name,
                        resolved.base_field,
                        resolved.level,
                        resolved.connected_pivot_tables,
                    ),
                    Err(_) => (None, None, None, None, Vec::new()),
                }
            } else {
                (None, None, None, None, Vec::new())
            };

        let placed_on_drawings = timeline_to_drawings
            .get(&part_name)
            .map(|drawings| drawings.iter().cloned().collect::<Vec<_>>())
            .unwrap_or_default();

        let (placed_on_sheets, placed_on_sheet_names) =
            placement_sheet_info(&placed_on_drawings, &drawing_to_sheets, &sheet_name_by_part);

        let mut selection = parse_timeline_selection(xml, excel_date_system).unwrap_or_default();
        if (selection.start.is_none() || selection.end.is_none()) && cache_part.is_some() {
            if let Some(cache_part) = cache_part.as_deref() {
                if let Some(bytes) = part(cache_part) {
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

        let field_name = resolve_timeline_field_name(
            &part,
            pivot_graph.as_ref(),
            base_field,
            &connected_pivot_tables,
        );

        timelines.push(TimelineDefinition {
            part_name: part_name.clone(),
            name: parsed.name,
            uid: parsed.uid,
            cache_part,
            cache_name,
            source_name,
            field_name,
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

fn sheet_name_by_part_with<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    resolve_relationship_target: &impl Fn(&str, &str) -> Result<Option<String>, XlsxError>,
) -> BTreeMap<String, String> {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = match part(workbook_part) {
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
        let resolved = resolve_relationship_target(workbook_part, &sheet.rel_id)
            .ok()
            .flatten()
            .or_else(|| {
                let guess_ws = format!("xl/worksheets/sheet{}.xml", sheet.sheet_id);
                if part(&guess_ws).is_some() {
                    return Some(guess_ws);
                }
                let guess_cs = format!("xl/chartsheets/sheet{}.xml", sheet.sheet_id);
                part(&guess_cs).map(|_| guess_cs)
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

fn patch_slicer_selection_xml(
    xml: &[u8],
    selection: &SlicerSelectionState,
) -> Result<Vec<u8>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + 128));
    let mut buf = Vec::new();

    let selected_items = selection.selected_items.as_ref();

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) => {
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"slicerCacheItem") {
                    if let Some(selected_items) = selected_items {
                        let (key, _, _) = parse_slicer_cache_item(&e, None)?;
                        if key.is_empty() {
                            let start_owned = e.into_owned();
                            buf.clear();
                            let (payload, nested_text) =
                                read_slicer_cache_item_payload(&mut reader, &mut buf)?;
                            let (key, _, _) =
                                parse_slicer_cache_item(&start_owned, nested_text.as_deref())?;
                            let desired = selected_items.contains(&key);
                            let patched =
                                patch_slicer_cache_item_start(&start_owned, Some(desired))?;
                            writer.write_event(Event::Start(patched))?;
                            for ev in payload {
                                writer.write_event(ev)?;
                            }
                        } else {
                            let desired = selected_items.contains(&key);
                            let patched = patch_slicer_cache_item_start(&e, Some(desired))?;
                            writer.write_event(Event::Start(patched))?;
                        }
                    } else {
                        // Clear explicit subset selections by removing any selection attributes.
                        let patched = patch_slicer_cache_item_start(&e, None)?;
                        writer.write_event(Event::Start(patched))?;
                    }
                } else {
                    writer.write_event(Event::Start(e))?;
                }
            }
            Event::Empty(e) => {
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"slicerCacheItem") {
                    if let Some(selected_items) = selected_items {
                        let (key, _, _) = parse_slicer_cache_item(&e, None)?;
                        let desired = selected_items.contains(&key);
                        let patched = patch_slicer_cache_item_start(&e, Some(desired))?;
                        writer.write_event(Event::Empty(patched))?;
                    } else {
                        let patched = patch_slicer_cache_item_start(&e, None)?;
                        writer.write_event(Event::Empty(patched))?;
                    }
                } else {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Event::Eof => break,
            other => writer.write_event(other.into_owned())?,
        }

        buf.clear();
    }

    Ok(writer.into_inner())
}

fn patch_slicer_cache_item_start(
    e: &BytesStart<'_>,
    desired_selected: Option<bool>,
) -> Result<BytesStart<'static>, XlsxError> {
    let name = e.name();
    let name = name.as_ref();
    let tag_name = std::str::from_utf8(name).unwrap_or("slicerCacheItem");
    let mut patched = BytesStart::new(tag_name);

    let mut saw_s = false;
    let mut saw_selected = false;

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key_local = local_name(attr.key.as_ref());

        if key_local.eq_ignore_ascii_case(b"s") {
            saw_s = true;
            if let Some(desired) = desired_selected {
                let value = if desired { "1" } else { "0" };
                patched.push_attribute((attr.key.as_ref(), value.as_bytes()));
            }
            // Clearing selection removes the attribute.
            continue;
        }

        if key_local.eq_ignore_ascii_case(b"selected") {
            saw_selected = true;
            if let Some(desired) = desired_selected {
                let value = if desired { "true" } else { "false" };
                patched.push_attribute((attr.key.as_ref(), value.as_bytes()));
            }
            continue;
        }

        let value = attr.unescape_value()?.into_owned();
        patched.push_attribute((attr.key.as_ref(), value.as_bytes()));
    }

    // Ensure every slicerCacheItem has an explicit selection attribute when setting an explicit
    // subset selection.
    if let Some(desired) = desired_selected {
        if !saw_s && !saw_selected {
            patched.push_attribute(("s", if desired { "1" } else { "0" }));
        }
    }

    Ok(patched.into_owned())
}

fn read_slicer_cache_item_payload<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
) -> Result<(Vec<Event<'static>>, Option<String>), XlsxError> {
    let mut depth = 0u32;
    let mut payload: Vec<Event<'static>> = Vec::new();
    let mut text = None;

    loop {
        let event = reader.read_event_into(buf)?;
        let mut done = false;

        match &event {
            Event::Start(start) => {
                if local_name(start.name().as_ref()).eq_ignore_ascii_case(b"slicerCacheItem") {
                    depth = depth.saturating_add(1);
                }
            }
            Event::End(end) => {
                if local_name(end.name().as_ref()).eq_ignore_ascii_case(b"slicerCacheItem") {
                    if depth == 0 {
                        done = true;
                    } else {
                        depth -= 1;
                    }
                }
            }
            Event::Text(value) => {
                if text.is_none() {
                    let value = value.unescape()?.into_owned();
                    if !value.trim().is_empty() {
                        text = Some(value.trim().to_string());
                    }
                }
            }
            Event::CData(value) => {
                if text.is_none() {
                    let value = String::from_utf8_lossy(value.as_ref()).into_owned();
                    if !value.trim().is_empty() {
                        text = Some(value.trim().to_string());
                    }
                }
            }
            Event::Eof => {
                buf.clear();
                break;
            }
            _ => {}
        }

        payload.push(event.into_owned());
        buf.clear();

        if done {
            break;
        }
    }

    Ok((payload, text))
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

fn detect_slicer_cache_item_selection_attr_key(xml: &[u8]) -> Result<Option<Vec<u8>>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if !tag.eq_ignore_ascii_case(b"slicerCacheItem") {
                    continue;
                }
                for attr in start.attributes().with_checks(false) {
                    let attr = attr?;
                    let key = local_name(attr.key.as_ref());
                    if key.eq_ignore_ascii_case(b"s") || key.eq_ignore_ascii_case(b"selected") {
                        return Ok(Some(attr.key.as_ref().to_vec()));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}

fn slicer_cache_xml_set_selection(
    xml: &[u8],
    selection: &SlicerSelectionState,
) -> Result<Vec<u8>, XlsxError> {
    let preferred_selection_attr_key =
        detect_slicer_cache_item_selection_attr_key(xml)?.unwrap_or_else(|| b"s".to_vec());

    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + 128));
    let mut buf = Vec::new();

    loop {
        let ev = reader.read_event_into(&mut buf)?;
        match ev {
            Event::Start(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"slicerCacheItem") =>
            {
                let patched = patch_slicer_cache_item(e, selection, &preferred_selection_attr_key)?;
                writer.write_event(Event::Start(patched))?;
            }
            Event::Empty(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"slicerCacheItem") =>
            {
                let patched = patch_slicer_cache_item(e, selection, &preferred_selection_attr_key)?;
                writer.write_event(Event::Empty(patched))?;
            }
            Event::Eof => break,
            other => writer.write_event(other.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn patch_slicer_cache_item(
    start: &BytesStart<'_>,
    selection: &SlicerSelectionState,
    preferred_selection_attr_key: &[u8],
) -> Result<BytesStart<'static>, XlsxError> {
    let mut attrs_raw: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut key_candidates: Vec<String> = Vec::new();

    for attr in start.attributes().with_checks(false) {
        let attr = attr?;
        let key_bytes = attr.key.as_ref().to_vec();
        let value_bytes = attr.value.as_ref().to_vec();
        attrs_raw.push((key_bytes.clone(), value_bytes));

        let attr_key = local_name(key_bytes.as_slice());
        if attr_key.eq_ignore_ascii_case(b"n")
            || attr_key.eq_ignore_ascii_case(b"name")
            || attr_key.eq_ignore_ascii_case(b"itemName")
            || attr_key.eq_ignore_ascii_case(b"caption")
            || attr_key.eq_ignore_ascii_case(b"uniqueName")
            || attr_key.eq_ignore_ascii_case(b"v")
            || attr_key.eq_ignore_ascii_case(b"x")
        {
            let value = attr.unescape_value()?.into_owned();
            if !value.is_empty() {
                key_candidates.push(value);
            }
        }
    }

    let desired_selected = match &selection.selected_items {
        None => None,
        Some(selected) => Some(
            !key_candidates.is_empty() && key_candidates.iter().any(|key| selected.contains(key)),
        ),
    };

    let element_name = start.name();
    let tag_name = std::str::from_utf8(element_name.as_ref()).unwrap_or("slicerCacheItem");
    let mut patched = BytesStart::new(tag_name);

    let mut wrote_selection = false;
    for (key_bytes, value_bytes) in attrs_raw {
        let attr_key = local_name(key_bytes.as_slice());
        let is_selection_attr =
            attr_key.eq_ignore_ascii_case(b"s") || attr_key.eq_ignore_ascii_case(b"selected");

        match desired_selected {
            None => {
                if is_selection_attr {
                    continue;
                }
                patched.push_attribute((key_bytes.as_slice(), value_bytes.as_slice()));
            }
            Some(selected) => {
                if is_selection_attr {
                    wrote_selection = true;
                    let value = if selected { b"1" } else { b"0" };
                    patched.push_attribute((key_bytes.as_slice(), value.as_slice()));
                } else {
                    patched.push_attribute((key_bytes.as_slice(), value_bytes.as_slice()));
                }
            }
        }
    }

    if let Some(selected) = desired_selected {
        if !wrote_selection {
            let value = if selected { b"1" } else { b"0" };
            patched.push_attribute((preferred_selection_attr_key, value.as_slice()));
        }
    }

    Ok(patched.into_owned())
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
        } else if attr_key.eq_ignore_ascii_case(b"caption")
            && key_caption.is_none()
            && !value.is_empty()
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

#[cfg(test)]
mod slicer_cache_patch_tests {
    use super::*;
    use std::collections::HashSet;

    #[test]
    fn updates_slicer_cache_xml_single_selected_item() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicerCacheData>
    <slicerCacheItem n="East" s="1"/>
    <slicerCacheItem n="West" s="0"/>
    <slicerCacheItem n="North" s="0"/>
  </slicerCacheData>
</slicerCache>"#;

        let selection = SlicerSelectionState {
            available_items: Vec::new(),
            selected_items: Some(HashSet::from(["West".to_string()])),
        };

        let updated = slicer_cache_xml_set_selection(xml, &selection).expect("patch xml");
        let parsed = parse_slicer_cache_selection(&updated).expect("parse updated");
        assert_eq!(
            parsed.selected_items,
            Some(HashSet::from(["West".to_string()]))
        );
    }

    #[test]
    fn updates_slicer_cache_xml_all_selected_removes_attrs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicerCacheData>
    <slicerCacheItem n="East" s="0"/>
    <slicerCacheItem n="West" s="1"/>
  </slicerCacheData>
</slicerCache>"#;

        let selection = SlicerSelectionState {
            available_items: Vec::new(),
            selected_items: None,
        };

        let updated = slicer_cache_xml_set_selection(xml, &selection).expect("patch xml");
        let updated_str = std::str::from_utf8(&updated).expect("utf8");
        assert!(
            !updated_str.contains(" s=\""),
            "expected selection attr to be removed, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains(" selected=\""),
            "expected selection attr to be removed, got:\n{updated_str}"
        );

        let parsed = parse_slicer_cache_selection(&updated).expect("parse updated");
        assert_eq!(parsed.selected_items, None);
    }

    #[test]
    fn updates_slicer_cache_xml_index_keys() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicerCacheData>
    <slicerCacheItem x="0" s="0"/>
    <slicerCacheItem x="1" s="1"/>
  </slicerCacheData>
</slicerCache>"#;

        let selection = SlicerSelectionState {
            available_items: Vec::new(),
            selected_items: Some(HashSet::from(["0".to_string()])),
        };

        let updated = slicer_cache_xml_set_selection(xml, &selection).expect("patch xml");
        let parsed = parse_slicer_cache_selection(&updated).expect("parse updated");
        assert_eq!(
            parsed.selected_items,
            Some(HashSet::from(["0".to_string()]))
        );
    }
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
    field_index: Option<u32>,
    connected_pivot_tables: Vec<String>,
    connected_tables: Vec<String>,
}

fn resolve_slicer_cache_definition<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    cache_part: &str,
) -> Result<ResolvedSlicerCacheDefinition, XlsxError> {
    let cache_bytes = part(cache_part).ok_or_else(|| XlsxError::MissingPart(cache_part.to_string()))?;
    let parsed = parse_slicer_cache_xml(cache_bytes)?;

    // Slicer caches can connect to both pivot tables and Excel Tables (ListObjects). Pivot table
    // connections are referenced by relationship id inside the cache XML, while table connections
    // are typically represented solely by relationships of type `.../table`.
    //
    // Excel generally emits `xl/slicerCaches/_rels/slicerCache*.xml.rels`, but the part can be
    // missing or malformed in real-world files. This code is best-effort: if relationships cannot
    // be parsed we fall back to empty connection lists rather than failing slicer discovery.
    let (rel_by_id, connected_tables) = parse_slicer_cache_rels_best_effort(part, cache_part);

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
        field_index: parsed.field_index,
        connected_pivot_tables: connected_pivot_tables.into_iter().collect(),
        connected_tables,
    })
}

fn parse_slicer_cache_rels_best_effort<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    cache_part: &str,
) -> (HashMap<String, crate::openxml::Relationship>, Vec<String>) {
    const TABLE_REL_TYPE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

    let rels_name = rels_part_name(cache_part);
    let relationships = match part(&rels_name) {
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

fn resolve_timeline_cache_definition<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    resolve_relationship_target: &impl Fn(&str, &str) -> Result<Option<String>, XlsxError>,
    cache_part: &str,
) -> Result<ResolvedTimelineCacheDefinition, XlsxError> {
    let cache_bytes = part(cache_part).ok_or_else(|| XlsxError::MissingPart(cache_part.to_string()))?;
    let parsed = parse_timeline_cache_xml(cache_bytes)?;
    let connected_pivot_tables =
        resolve_relationship_targets(resolve_relationship_target, cache_part, parsed.pivot_table_rids)?;

    Ok(ResolvedTimelineCacheDefinition {
        cache_name: parsed.cache_name,
        source_name: parsed.source_name,
        base_field: parsed.base_field,
        level: parsed.level,
        connected_pivot_tables,
    })
}

fn resolve_relationship_targets(
    resolve_relationship_target: &impl Fn(&str, &str) -> Result<Option<String>, XlsxError>,
    base_part: &str,
    relationship_ids: Vec<String>,
) -> Result<Vec<String>, XlsxError> {
    let mut targets = BTreeSet::new();
    for rid in relationship_ids {
        // These relationships are optional (they connect caches to pivot tables). Be tolerant of
        // malformed `.rels` payloads and continue when we can't resolve a target.
        match resolve_relationship_target(base_part, &rid) {
            Ok(Some(target)) => {
                targets.insert(target);
            }
            Ok(None) | Err(_) => {}
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
                    for attr in start.attributes().with_checks(false) {
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
                    for attr in start.attributes().with_checks(false) {
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
    field_index: Option<u32>,
    pivot_table_rids: Vec<String>,
}

fn parse_slicer_cache_xml(xml: &[u8]) -> Result<ParsedSlicerCacheXml, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut cache_name = None;
    let mut source_name = None;
    let mut field_index = None;
    let mut pivot_table_rids = Vec::new();

    fn capture_field_index(field_index: &mut Option<u32>, key: &[u8], value: &str) {
        if field_index.is_some() {
            return;
        }
        // Excel/OOXML producers are not consistent about how they encode the filtered field.
        // Be tolerant and capture the first numeric value we see for a field-ish attribute.
        if key.eq_ignore_ascii_case(b"field")
            || key.eq_ignore_ascii_case(b"fieldId")
            || key.eq_ignore_ascii_case(b"fieldIndex")
            || key.eq_ignore_ascii_case(b"sourceField")
            || key.eq_ignore_ascii_case(b"sourceFieldId")
            || key.eq_ignore_ascii_case(b"sourceFieldIndex")
            || key.eq_ignore_ascii_case(b"pivotField")
            || key.eq_ignore_ascii_case(b"pivotFieldId")
            || key.eq_ignore_ascii_case(b"pivotFieldIndex")
            || key.eq_ignore_ascii_case(b"baseField")
            || key.eq_ignore_ascii_case(b"fld")
        {
            if let Ok(idx) = value.trim().parse::<u32>() {
                *field_index = Some(idx);
            }
        }
    }

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"slicerCache") {
                    for attr in start.attributes().with_checks(false) {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"name") {
                            cache_name = Some(value);
                        } else if key.eq_ignore_ascii_case(b"sourceName") {
                            source_name = Some(value);
                        } else {
                            capture_field_index(&mut field_index, key, &value);
                        }
                    }
                } else if tag.eq_ignore_ascii_case(b"slicerCachePivotTable") {
                    for attr in start.attributes().with_checks(false) {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            pivot_table_rids.push(attr.unescape_value()?.into_owned());
                        }
                    }
                } else {
                    // Some producers store field indices on nested elements (e.g. table/pivot slicer
                    // cache variants). Capture field-ish attributes anywhere in the slicer cache XML.
                    for attr in start.attributes().with_checks(false) {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        capture_field_index(&mut field_index, key, &value);
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
        field_index,
        pivot_table_rids,
    })
}

fn resolve_timeline_field_name<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    pivot_graph: Option<&crate::pivots::graph::XlsxPivotGraph>,
    base_field: Option<u32>,
    connected_pivot_tables: &[String],
) -> Option<String> {
    let base_field = base_field?;
    let pivot_table_part = connected_pivot_tables.first()?;
    resolve_pivot_cache_field_name(
        part,
        pivot_graph,
        pivot_table_part,
        base_field,
    )
}

fn resolve_slicer_field_name<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    pivot_graph: Option<&crate::pivots::graph::XlsxPivotGraph>,
    cache_part: Option<&str>,
    connected_pivot_tables: &[String],
    selection: &SlicerSelectionState,
) -> Option<String> {
    // First, see if the slicer cache definition includes an explicit field index.
    if let Some(cache_part) = cache_part {
        if let Ok(resolved) = resolve_slicer_cache_definition(part, cache_part) {
            if let (Some(field_index), Some(pivot_table_part)) =
                (resolved.field_index, connected_pivot_tables.first())
            {
                if let Some(name) = resolve_pivot_cache_field_name(
                    part,
                    pivot_graph,
                    pivot_table_part,
                    field_index,
                ) {
                    return Some(name);
                }
            }
        }
    }

    infer_slicer_field_name(part, pivot_graph, connected_pivot_tables, selection)
}

fn resolve_pivot_cache_parts<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    pivot_graph: Option<&crate::pivots::graph::XlsxPivotGraph>,
    pivot_table_part: &str,
) -> Option<(String, Option<String>)> {
    let pivot_table_part = pivot_table_part.trim_start_matches('/');
    if let Some(graph) = pivot_graph {
        if let Some(instance) = graph
            .pivot_tables
            .iter()
            .find(|pt| pt.pivot_table_part == pivot_table_part)
        {
            if let Some(def_part) = instance.cache_definition_part.clone() {
                return Some((def_part, instance.cache_records_part.clone()));
            }
        }
    }

    // Fallback: parse `cacheId` from the pivot table and use the canonical naming convention.
    let bytes = part(pivot_table_part)?;
    let cache_id = parse_pivot_table_cache_id(bytes)?;
    let def_guess = format!("xl/pivotCache/pivotCacheDefinition{cache_id}.xml");
    if part(&def_guess).is_none() {
        return None;
    }
    let records_guess = format!("xl/pivotCache/pivotCacheRecords{cache_id}.xml");
    let records_part = part(&records_guess).map(|_| records_guess);
    Some((def_guess, records_part))
}

fn parse_pivot_table_cache_id(xml: &[u8]) -> Option<u32> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buf).ok()?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"pivotTableDefinition") {
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.ok()?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"cacheId") {
                            let value = attr.unescape_value().ok()?.into_owned();
                            return value.parse::<u32>().ok();
                        }
                    }
                    return None;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    None
}

fn resolve_pivot_cache_field_name<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    pivot_graph: Option<&crate::pivots::graph::XlsxPivotGraph>,
    pivot_table_part: &str,
    field_index: u32,
) -> Option<String> {
    let (def_part, _) = resolve_pivot_cache_parts(part, pivot_graph, pivot_table_part)?;
    let def_bytes = part(&def_part)?;
    let def = crate::pivots::cache_definition::parse_pivot_cache_definition(def_bytes).ok()?;
    def.cache_fields
        .get(field_index as usize)
        .map(|f| f.name.clone())
}

fn infer_slicer_field_name<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    pivot_graph: Option<&crate::pivots::graph::XlsxPivotGraph>,
    connected_pivot_tables: &[String],
    selection: &SlicerSelectionState,
) -> Option<String> {
    let candidate_items: Vec<String> = if !selection.available_items.is_empty() {
        selection.available_items.clone()
    } else if let Some(selected) = &selection.selected_items {
        selected.iter().cloned().collect()
    } else {
        Vec::new()
    };

    let candidate_items: HashSet<String> = candidate_items
        .into_iter()
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .collect();
    if candidate_items.is_empty() {
        return None;
    }

    let mut match_sets = Vec::new();
    for pivot_table_part in connected_pivot_tables {
        let Some(matches) = matching_fields_for_pivot_table(
            part,
            pivot_graph,
            pivot_table_part,
            &candidate_items,
        ) else {
            continue;
        };
        if !matches.is_empty() {
            match_sets.push(matches);
        }
    }

    if match_sets.is_empty() {
        return None;
    }

    let mut intersection = match_sets[0].clone();
    for matches in match_sets.iter().skip(1) {
        intersection = intersection
            .intersection(matches)
            .cloned()
            .collect::<HashSet<_>>();
    }
    if intersection.len() == 1 {
        return intersection.into_iter().next();
    }

    let mut union = HashSet::new();
    for matches in match_sets {
        union.extend(matches);
    }
    if union.len() == 1 {
        return union.into_iter().next();
    }

    None
}

fn matching_fields_for_pivot_table<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    pivot_graph: Option<&crate::pivots::graph::XlsxPivotGraph>,
    pivot_table_part: &str,
    candidate_items: &HashSet<String>,
) -> Option<HashSet<String>> {
    let (def_part, records_part) = resolve_pivot_cache_parts(part, pivot_graph, pivot_table_part)?;
    let def_bytes = part(&def_part)?;
    let def = crate::pivots::cache_definition::parse_pivot_cache_definition(def_bytes).ok()?;
    if def.cache_fields.is_empty() {
        return None;
    }

    let shared_items = parse_pivot_cache_shared_items(def_bytes).ok().unwrap_or_default();

    let field_count = def.cache_fields.len();
    let mut matches = HashSet::new();
    let mut needs_records = vec![true; field_count];

    // If shared items are available, use them directly to match candidate values without
    // scanning cache records.
    for idx in 0..field_count {
        let Some(items) = shared_items.get(idx) else {
            continue;
        };
        if items.is_empty() {
            continue;
        }
        let values: HashSet<&str> = items.iter().map(|s| s.trim()).collect();
        if candidate_items.iter().all(|item| {
            values.contains(item.as_str())
                || item
                    .parse::<usize>()
                    .ok()
                    .is_some_and(|idx| idx < items.len())
        }) {
            matches.insert(def.cache_fields[idx].name.clone());
        }
        needs_records[idx] = false;
    }

    if needs_records.iter().all(|v| !*v) {
        return Some(matches);
    }

    let Some(records_part) = records_part else {
        // No cache records and no shared item metadata for some fields.
        return Some(matches);
    };

    let Some(records_bytes) = part(&records_part) else {
        return Some(matches);
    };

    let mut found: Vec<HashSet<String>> = (0..field_count).map(|_| HashSet::new()).collect();
    let mut done = vec![false; field_count];

    let mut reader = crate::pivots::cache_records::PivotCacheRecordsReader::new(records_bytes);
    while let Some(record) = reader.next_record() {
        for field_idx in 0..field_count {
            if !needs_records[field_idx] || done[field_idx] {
                continue;
            }
            let value = record
                .get(field_idx)
                .cloned()
                .unwrap_or(crate::pivots::cache_records::PivotCacheValue::Missing);

            let shared_for_field = shared_items.get(field_idx);
            if capture_candidate_items_from_value(
                &value,
                shared_for_field,
                candidate_items,
                &mut found[field_idx],
            ) {
                done[field_idx] = true;
            }
        }

        if needs_records
            .iter()
            .enumerate()
            .all(|(idx, needs)| !*needs || done[idx])
        {
            break;
        }
    }

    for idx in 0..field_count {
        if needs_records[idx] && found[idx].len() == candidate_items.len() {
            matches.insert(def.cache_fields[idx].name.clone());
        }
    }

    Some(matches)
}

fn capture_candidate_items_from_value(
    value: &crate::pivots::cache_records::PivotCacheValue,
    shared_items_for_field: Option<&Vec<String>>,
    candidate_items: &HashSet<String>,
    found: &mut HashSet<String>,
) -> bool {
    use crate::pivots::cache_records::PivotCacheValue;

    if found.len() == candidate_items.len() {
        return true;
    }

    match value {
        PivotCacheValue::Missing => {}
        PivotCacheValue::String(s) => {
            let v = s.trim();
            if !v.is_empty() && candidate_items.contains(v) {
                found.insert(v.to_string());
            }
        }
        PivotCacheValue::Number(n) => {
            let s = n.to_string();
            if candidate_items.contains(s.as_str()) {
                found.insert(s);
            }
        }
        PivotCacheValue::Bool(b) => {
            let raw = if *b { "1" } else { "0" };
            if candidate_items.contains(raw) {
                found.insert(raw.to_string());
            }
            let lower = if *b { "true" } else { "false" };
            if candidate_items.contains(lower) {
                found.insert(lower.to_string());
            }
            let upper = if *b { "TRUE" } else { "FALSE" };
            if candidate_items.contains(upper) {
                found.insert(upper.to_string());
            }
        }
        PivotCacheValue::Error(e) => {
            let v = e.trim();
            if !v.is_empty() && candidate_items.contains(v) {
                found.insert(v.to_string());
            }
        }
        PivotCacheValue::DateTime(dt) => {
            let v = dt.trim();
            if !v.is_empty() && candidate_items.contains(v) {
                found.insert(v.to_string());
            }
            if let Some(date) = crate::pivots::cache_records::pivot_cache_datetime_to_naive_date(v)
            {
                let iso = format!("{:04}-{:02}-{:02}", date.year(), date.month(), date.day());
                if candidate_items.contains(iso.as_str()) {
                    found.insert(iso);
                }
            }
        }
        PivotCacheValue::Index(idx) => {
            let raw = idx.to_string();
            if candidate_items.contains(raw.as_str()) {
                found.insert(raw);
            }
            if let Some(items) = shared_items_for_field.and_then(|v| v.get(*idx as usize)) {
                let v = items.trim();
                if !v.is_empty() && candidate_items.contains(v) {
                    found.insert(v.to_string());
                }
            }
        }
    }

    found.len() == candidate_items.len()
}

fn parse_pivot_cache_shared_items(xml: &[u8]) -> Result<Vec<Vec<String>>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut out: Vec<Vec<String>> = Vec::new();
    let mut current_field: Option<usize> = None;
    let mut in_shared_items = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) => {
                let name = start.name();
                let tag = local_name(name.as_ref());
                if tag.eq_ignore_ascii_case(b"cacheField") {
                    out.push(Vec::new());
                    current_field = Some(out.len() - 1);
                    in_shared_items = false;
                } else if tag.eq_ignore_ascii_case(b"sharedItems") {
                    in_shared_items = true;
                } else if in_shared_items {
                    if let Some(field_idx) = current_field {
                        if let Some(value) = parse_cache_shared_item_start(&mut reader, &start)? {
                            out[field_idx].push(value);
                        }
                    }
                }
            }
            Event::Empty(start) => {
                let name = start.name();
                let tag = local_name(name.as_ref());
                if tag.eq_ignore_ascii_case(b"cacheField") {
                    out.push(Vec::new());
                    current_field = Some(out.len() - 1);
                    in_shared_items = false;
                } else if tag.eq_ignore_ascii_case(b"sharedItems") {
                    in_shared_items = true;
                } else if in_shared_items {
                    if let Some(field_idx) = current_field {
                        if let Some(value) = parse_cache_shared_item_empty(&start)? {
                            out[field_idx].push(value);
                        }
                    }
                }
            }
            Event::End(end) => {
                let name = end.name();
                let tag = local_name(name.as_ref());
                if tag.eq_ignore_ascii_case(b"cacheField") {
                    current_field = None;
                    in_shared_items = false;
                } else if tag.eq_ignore_ascii_case(b"sharedItems") {
                    in_shared_items = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_cache_shared_item_empty(
    start: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<String>, XlsxError> {
    let name = start.name();
    let tag = local_name(name.as_ref());
    let attr_v = start
        .attributes()
        .with_checks(false)
        .filter_map(|a| a.ok())
        .find(|a| local_name(a.key.as_ref()).eq_ignore_ascii_case(b"v"))
        .and_then(|a| a.unescape_value().ok())
        .map(|v| v.into_owned());

    let value = match tag {
        b"m" => Some(String::new()),
        b"s" | b"n" | b"d" | b"e" => attr_v,
        b"b" => attr_v.map(|v| {
            if v.trim() == "1" || v.trim().eq_ignore_ascii_case("true") {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }),
        _ => None,
    };
    Ok(value)
}

fn parse_cache_shared_item_start(
    reader: &mut Reader<Cursor<&[u8]>>,
    start: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<String>, XlsxError> {
    // Shared item elements are typically self-closing, but be tolerant of wrapped/text forms.
    let name = start.name();
    let tag = local_name(name.as_ref());
    let value = parse_cache_shared_item_empty(start)?;
    if value.is_some() {
        // Skip to end of the element to keep the reader state consistent.
        let mut skip_buf = Vec::new();
        reader.read_to_end_into(start.name(), &mut skip_buf)?;
        return Ok(value);
    }

    // Fallback: read text content.
    let mut buf = Vec::new();
    let mut out = None;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => {
                if out.is_none() {
                    out = Some(e.unescape()?.into_owned());
                }
            }
            Event::End(e) if e.name() == start.name() => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    let out = out.map(|s| s.trim().to_string());
    Ok(match tag {
        b"m" => Some(String::new()),
        _ => out.filter(|s| !s.is_empty()),
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
                    for attr in start.attributes().with_checks(false) {
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
                    for attr in start.attributes().with_checks(false) {
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
                    for attr in start.attributes().with_checks(false) {
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
                    for attr in start.attributes().with_checks(false) {
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
    use crate::pivots::PivotCacheField;
    use pretty_assertions::assert_eq;
    use std::collections::HashSet;
    use formula_engine::pivot::PivotFieldRef;
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

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
            source_field: "Region".into(),
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
            source_field: "Region".into(),
            allowed: None,
         };

        assert_eq!(actual, expected);
    }

    #[test]
    fn pivot_slicer_parts_tolerates_duplicate_attributes() {
        // Duplicate attributes are not well-formed XML, but we want to be tolerant of producers
        // that emit them (QuickXML's strict attribute checks reject them).
        let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        name="Slicer1" name="Slicer1">
  <slicerCache r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</slicer>"#;

        // Minimal relationship part for `rId1` -> slicer cache definition.
        let slicer_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:slicerCache" Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

        let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
            name="Cache1" name="Cache1" sourceName="Field1"/>"#;

        let pkg = build_package(&[
            ("xl/slicers/slicer1.xml", slicer_xml),
            ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels),
            ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        ]);

        let parsed = pkg
            .pivot_slicer_parts()
            .expect("should parse slicer parts despite duplicate attributes");

        assert_eq!(parsed.slicers.len(), 1);
        assert_eq!(parsed.timelines.len(), 0);

        let slicer = &parsed.slicers[0];
        assert_eq!(slicer.part_name, "xl/slicers/slicer1.xml");
        assert_eq!(slicer.name.as_deref(), Some("Slicer1"));
        assert_eq!(
            slicer.cache_part.as_deref(),
            Some("xl/slicerCaches/slicerCache1.xml")
        );
        assert_eq!(slicer.cache_name.as_deref(), Some("Cache1"));
        assert_eq!(slicer.source_name.as_deref(), Some("Field1"));
    }

    #[test]
    fn resolves_numeric_key_via_shared_items() {
        let mut field = PivotCacheField::default();
        field.name = "Field1".to_string();
        field.shared_items = Some(vec![
            PivotCacheValue::String("East".to_string()),
            PivotCacheValue::Number(42.0),
            PivotCacheValue::Bool(true),
        ]);

        let mut def = PivotCacheDefinition::default();
        def.cache_fields.push(field);

        assert_eq!(
            resolve_slicer_item_key(&def, 0, "0"),
            Some(ScalarValue::Text("East".to_string()))
        );
        assert_eq!(
            resolve_slicer_item_key(&def, 0, "1"),
            Some(ScalarValue::from(42.0))
        );
        assert_eq!(
            resolve_slicer_item_key(&def, 0, "2"),
            Some(ScalarValue::Bool(true))
        );

        let mut resolver = shared_item_key_resolver(&def, 0);
        assert_eq!(resolver("1"), Some(ScalarValue::from(42.0)));
    }

    #[test]
    fn returns_none_for_non_numeric_keys() {
        let def = PivotCacheDefinition::default();
        assert_eq!(resolve_slicer_item_key(&def, 0, "East"), None);
    }

    #[test]
    fn datetime_shared_items_resolve_to_date() {
        let mut field = PivotCacheField::default();
        field.name = "DateField".to_string();
        field.shared_items = Some(vec![PivotCacheValue::DateTime(
            "2024-01-15T00:00:00Z".to_string(),
        )]);

        let mut def = PivotCacheDefinition::default();
        def.cache_fields.push(field);

        assert_eq!(
            resolve_slicer_item_key(&def, 0, "0"),
            Some(ScalarValue::Date(
                NaiveDate::from_ymd_opt(2024, 1, 15).unwrap()
            ))
        );
    }

    #[test]
    fn timeline_selection_to_engine_filter_enumerates_dates_from_shared_items() {
        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "OrderDate".to_string(),
                shared_items: Some(vec![
                    PivotCacheValue::DateTime("2024-01-01".to_string()),
                    PivotCacheValue::DateTime("2024-01-02".to_string()),
                    PivotCacheValue::DateTime("2024-01-03".to_string()),
                ]),
                ..Default::default()
            }],
            ..Default::default()
        };

        let selection = TimelineSelectionState {
            start: Some("2024-01-02".to_string()),
            end: Some("2024-01-03".to_string()),
        };

        let filter = timeline_selection_to_engine_filter_field_with_cache(&cache_def, 0, &selection);

        assert_eq!(
            filter.source_field,
            PivotFieldRef::CacheFieldName("OrderDate".to_string())
        );

        let mut expected = HashSet::new();
        expected.insert(PivotKeyPart::Date(NaiveDate::from_ymd_opt(2024, 1, 2).unwrap()));
        expected.insert(PivotKeyPart::Date(NaiveDate::from_ymd_opt(2024, 1, 3).unwrap()));

        assert_eq!(filter.allowed, Some(expected));
    }

    #[test]
    fn timeline_selection_to_engine_filter_handles_numeric_serials_with_best_effort_date_system() {
        // Shared items can store date values as numeric serials. Without workbook metadata we can't
        // know whether they are 1900- or 1904-based; infer the date system from the selection range
        // so we don't end up producing an empty filter due to a 4-year offset.
        let cache_def = PivotCacheDefinition {
            cache_fields: vec![PivotCacheField {
                name: "OrderDate".to_string(),
                shared_items: Some(vec![
                    PivotCacheValue::Number(1.0),
                    PivotCacheValue::Number(2.0),
                    PivotCacheValue::Number(3.0),
                ]),
                ..Default::default()
            }],
            ..Default::default()
        };

        // In the 1904 date system, serial 1 = 1904-01-02 and serial 2 = 1904-01-03.
        let selection = TimelineSelectionState {
            start: Some("1904-01-02".to_string()),
            end: Some("1904-01-03".to_string()),
        };

        let filter = timeline_selection_to_engine_filter_field_with_cache(&cache_def, 0, &selection);

        let mut expected = HashSet::new();
        expected.insert(PivotKeyPart::Date(NaiveDate::from_ymd_opt(1904, 1, 2).unwrap()));
        expected.insert(PivotKeyPart::Date(NaiveDate::from_ymd_opt(1904, 1, 3).unwrap()));

        assert_eq!(filter.allowed, Some(expected));
    }
}

#[cfg(test)]
mod slicer_selection_write_tests {
    use super::*;

    use std::collections::HashSet;
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn slicer_selection_round_trip_updates_cache_items() {
        let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        name="Slicer1">
  <slicerCache r:id="rId1"/>
</slicer>"#;

        let slicer_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:slicerCache" Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

        let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicerCacheItems>
    <slicerCacheItem n="East" s="0"/>
    <slicerCacheItem n="West" s="1"/>
    <slicerCacheItem n="North" s="0"/>
  </slicerCacheItems>
</slicerCache>"#;

        let mut pkg = build_package(&[
            ("xl/slicers/slicer1.xml", slicer_xml),
            ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels),
            ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        ]);

        let selection = SlicerSelectionState {
            available_items: Vec::new(),
            selected_items: Some(HashSet::from(["East".to_string()])),
        };

        pkg.set_slicer_selection("xl/slicerCaches/slicerCache1.xml", &selection)
            .expect("set selection");

        let parts = pkg.pivot_slicer_parts().expect("parse slicer parts");
        assert_eq!(parts.slicers.len(), 1);
        assert_eq!(
            parts.slicers[0].selection.selected_items,
            Some(HashSet::from(["East".to_string()]))
        );
    }

    #[test]
    fn slicer_selection_all_selected_clears_explicit_subset() {
        let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        name="Slicer1">
  <slicerCache r:id="rId1"/>
</slicer>"#;

        let slicer_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:slicerCache" Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

        let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicerCacheItems>
    <slicerCacheItem n="East" s="1"/>
    <slicerCacheItem n="West" s="0"/>
    <slicerCacheItem n="North" s="0"/>
  </slicerCacheItems>
</slicerCache>"#;

        let mut pkg = build_package(&[
            ("xl/slicers/slicer1.xml", slicer_xml),
            ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels),
            ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        ]);

        let selection = SlicerSelectionState {
            available_items: Vec::new(),
            selected_items: None,
        };

        pkg.set_slicer_selection("xl/slicerCaches/slicerCache1.xml", &selection)
            .expect("set selection");

        let parts = pkg.pivot_slicer_parts().expect("parse slicer parts");
        assert_eq!(parts.slicers.len(), 1);
        assert_eq!(parts.slicers[0].selection.selected_items, None);
    }
}

#[cfg(test)]
mod timeline_selection_write_tests {
    use super::*;

    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn timeline_selection_updates_existing_selection_element() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <selection startDate="2024-01-01" endDate="2024-01-31"/>
</timelineCacheDefinition>"#;

        let selection = TimelineSelectionState {
            start: Some("2024-02-01".to_string()),
            end: Some("2024-02-29".to_string()),
        };

        let updated =
            patch_timeline_selection_xml(xml, &selection, DateSystem::V1900).expect("patch xml");
        let parsed = parse_timeline_selection(&updated, DateSystem::V1900.to_engine_date_system())
            .expect("parse selection");
        assert_eq!(parsed.start.as_deref(), Some("2024-02-01"));
        assert_eq!(parsed.end.as_deref(), Some("2024-02-29"));

        let updated_str = std::str::from_utf8(&updated).expect("utf8");
        assert!(updated_str.contains("startDate=\"2024-02-01\""));
        assert!(updated_str.contains("endDate=\"2024-02-29\""));
    }

    #[test]
    fn timeline_selection_inserts_when_missing() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"/>"#;

        let selection = TimelineSelectionState {
            start: Some("2024-03-01".to_string()),
            end: Some("2024-03-31".to_string()),
        };

        let updated =
            patch_timeline_selection_xml(xml, &selection, DateSystem::V1900).expect("patch xml");
        let parsed = parse_timeline_selection(&updated, DateSystem::V1900.to_engine_date_system())
            .expect("parse selection");
        assert_eq!(parsed.start.as_deref(), Some("2024-03-01"));
        assert_eq!(parsed.end.as_deref(), Some("2024-03-31"));

        let updated_str = std::str::from_utf8(&updated).expect("utf8");
        assert!(updated_str.contains("selection"));
        assert!(updated_str.contains("startDate=\"2024-03-01\""));
        assert!(updated_str.contains("endDate=\"2024-03-31\""));
    }

    #[test]
    fn timeline_selection_inserts_with_prefix_when_root_is_prefixed() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x15:timelineCacheDefinition xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"/>"#;

        let selection = TimelineSelectionState {
            start: Some("2024-05-01".to_string()),
            end: Some("2024-05-31".to_string()),
        };

        let updated =
            patch_timeline_selection_xml(xml, &selection, DateSystem::V1900).expect("patch xml");
        let updated_str = std::str::from_utf8(&updated).expect("utf8");
        assert!(
            updated_str.contains("x15:selection"),
            "expected inserted selection to reuse root prefix: {updated_str}"
        );
    }

    #[test]
    fn timeline_selection_updates_numeric_serials_using_workbook_date_system() {
        // When a timeline persists selection endpoints as numeric serials, keep that representation
        // when patching so Excel continues to interpret it correctly.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <selection startDate="1" endDate="2"/>
</timelineCacheDefinition>"#;

        // In the 1904 system, serial 0 = 1904-01-01.
        let selection = TimelineSelectionState {
            start: Some("1904-01-04".to_string()),
            end: Some("1904-01-05".to_string()),
        };

        let updated =
            patch_timeline_selection_xml(xml, &selection, DateSystem::V1904).expect("patch xml");
        let parsed = parse_timeline_selection(&updated, DateSystem::V1904.to_engine_date_system())
            .expect("parse selection");
        assert_eq!(parsed.start.as_deref(), Some("1904-01-04"));
        assert_eq!(parsed.end.as_deref(), Some("1904-01-05"));

        let updated_str = std::str::from_utf8(&updated).expect("utf8");
        assert!(updated_str.contains("startDate=\"3\""), "{updated_str}");
        assert!(updated_str.contains("endDate=\"4\""), "{updated_str}");
    }

    #[test]
    fn set_timeline_selection_round_trips_via_pivot_slicer_parts() {
        let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          name="Timeline1" uid="{00000000-0000-0000-0000-000000000001}">
  <timelineCache r:id="rId1"/>
</timeline>"#;

        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.microsoft.com/office/2007/relationships/timelineCacheDefinition"
                Target="../timelineCaches/timelineCacheDefinition1.xml"/>
</Relationships>"#;

        let cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                         name="TimelineCache1" sourceName="Date"/>"#;

        let mut pkg = build_package(&[
            ("xl/timelines/timeline1.xml", timeline_xml),
            ("xl/timelines/_rels/timeline1.xml.rels", rels_xml),
            ("xl/timelineCaches/timelineCacheDefinition1.xml", cache_xml),
        ]);

        let selection = TimelineSelectionState {
            start: Some("2024-04-01".to_string()),
            end: Some("2024-04-30".to_string()),
        };
        pkg.set_timeline_selection("xl/timelineCaches/timelineCacheDefinition1.xml", &selection)
            .expect("set selection");

        let parts = pkg.pivot_slicer_parts().expect("parse slicer parts");
        assert_eq!(parts.timelines.len(), 1);
        assert_eq!(
            parts.timelines[0].selection.start.as_deref(),
            Some("2024-04-01")
        );
        assert_eq!(
            parts.timelines[0].selection.end.as_deref(),
            Some("2024-04-30")
        );
    }

    #[test]
    fn set_timeline_selection_patches_cache_even_if_timeline_rels_is_malformed() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

        let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          name="Timeline1" uid="{00000000-0000-0000-0000-000000000001}">
  <timelineCache r:id="rId1"/>
</timeline>"#;

        // Intentionally malformed `.rels` payload: the `<Relationship>` start tag is never closed.
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.microsoft.com/office/2007/relationships/timelineCacheDefinition"
                Target="../timelineCaches/timelineCacheDefinition1.xml">
</Relationships>"#;

        let cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <selection startDate="2024-01-01" endDate="2024-01-31"/>
</timelineCacheDefinition>"#;

        let mut pkg = build_package(&[
            ("xl/workbook.xml", workbook_xml),
            ("xl/timelines/timeline1.xml", timeline_xml),
            ("xl/timelines/_rels/timeline1.xml.rels", rels_xml),
            ("xl/timelineCaches/timelineCacheDefinition1.xml", cache_xml),
        ]);

        let selection = TimelineSelectionState {
            start: Some("2024-02-01".to_string()),
            end: Some("2024-02-29".to_string()),
        };

        pkg.set_timeline_selection("xl/timelineCaches/timelineCacheDefinition1.xml", &selection)
            .expect("set selection should be best-effort for malformed rels");

        let updated = pkg
            .part("xl/timelineCaches/timelineCacheDefinition1.xml")
            .expect("cache part exists");
        let updated_str = std::str::from_utf8(updated).expect("utf8");
        assert!(updated_str.contains("startDate=\"2024-02-01\""));
        assert!(updated_str.contains("endDate=\"2024-02-29\""));
    }

    #[test]
    fn set_timeline_selection_patches_timeline_even_if_timeline_rels_is_malformed() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

        let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          name="Timeline1" uid="{00000000-0000-0000-0000-000000000001}">
  <timelineCache r:id="rId1"/>
</timeline>"#;

        // Malformed relationship part: the `<Relationship>` start tag is never closed.
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.microsoft.com/office/2007/relationships/timelineCacheDefinition"
                Target="../timelineCaches/timelineCacheDefinition1.xml">
</Relationships>"#;

        let mut pkg = build_package(&[
            ("xl/workbook.xml", workbook_xml),
            ("xl/timelines/timeline1.xml", timeline_xml),
            ("xl/timelines/_rels/timeline1.xml.rels", rels_xml),
        ]);

        let selection = TimelineSelectionState {
            start: Some("2024-02-01".to_string()),
            end: Some("2024-02-29".to_string()),
        };

        pkg.set_timeline_selection("xl/timelines/timeline1.xml", &selection)
            .expect("set selection should patch the timeline even if rels is malformed");

        let updated = pkg.part("xl/timelines/timeline1.xml").expect("timeline part");
        let parsed =
            parse_timeline_selection(updated, DateSystem::V1900.to_engine_date_system()).unwrap();
        assert_eq!(parsed.start.as_deref(), Some("2024-02-01"));
        assert_eq!(parsed.end.as_deref(), Some("2024-02-29"));
    }
}

#[cfg(test)]
mod shared_item_resolver_tests {
    use super::*;
    use crate::pivots::{PivotCacheDefinition, PivotCacheField, PivotCacheValue};
    use pretty_assertions::assert_eq;
    use std::collections::HashSet;

    #[test]
    fn slicer_selection_resolves_x_indices_to_shared_items() {
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

        let selection = SlicerSelectionState {
            available_items: vec!["0".to_string(), "1".to_string()],
            selected_items: Some(HashSet::from(["0".to_string()])),
        };

        let filter = slicer_selection_to_row_filter_with_resolver("Region", &selection, |key| {
            key.parse::<u32>()
                .ok()
                .and_then(|idx| cache_def.resolve_shared_item(0, idx))
        });

        match filter {
            RowFilter::Slicer { field, selection } => {
                assert_eq!(field, "Region");
                assert_eq!(
                    selection,
                    SlicerSelection::Items(HashSet::from([ScalarValue::from("East")]))
                );
            }
            other => panic!("expected RowFilter::Slicer, got {other:?}"),
        }
    }
}
