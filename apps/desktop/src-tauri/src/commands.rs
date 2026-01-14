use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotFieldRef, PivotKeyPart, ShowAsType,
    SortOrder, SubtotalPosition,
};
#[cfg(feature = "desktop")]
use formula_model::charts::ChartModel as FormulaChartModel;
#[cfg(feature = "desktop")]
use formula_model::drawings::Anchor as FormulaDrawingAnchor;
use formula_model::{SheetVisibility as ModelSheetVisibility, TabColor};
use serde::{de, Deserialize, Serialize};
#[cfg(any(feature = "desktop", test))]
use serde_json::json;
use serde_json::Value as JsonValue;
#[cfg(any(feature = "desktop", test))]
use std::collections::BTreeMap;
use std::collections::HashSet;
use std::fmt;
use std::marker::PhantomData;
#[cfg(any(feature = "desktop", test))]
use url::Url;

use crate::macro_trust::MacroTrustDecision;
use crate::resource_limits::{MAX_CELL_FORMULA_BYTES, MAX_CELL_VALUE_STRING_BYTES};
#[cfg(feature = "desktop")]
use crate::storage::collab_encryption_keys::{
    CollabEncryptionKeyEntry, CollabEncryptionKeyListEntry, CollabEncryptionKeyStore,
};
#[cfg(feature = "desktop")]
use crate::storage::collab_tokens::{CollabTokenEntry, CollabTokenStore};
#[cfg(feature = "desktop")]
use crate::storage::power_query_cache_key::{PowerQueryCacheKey, PowerQueryCacheKeyStore};
#[cfg(feature = "desktop")]
use crate::storage::power_query_credentials::{
    PowerQueryCredentialEntry, PowerQueryCredentialListEntry, PowerQueryCredentialStore,
};
#[cfg(feature = "desktop")]
use crate::storage::power_query_refresh_state::PowerQueryRefreshStateStore;

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PrintCellRange {
    pub start_row: u32,
    pub end_row: u32,
    pub start_col: u32,
    pub end_col: u32,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PrintRowRange {
    pub start: u32,
    pub end: u32,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PrintColRange {
    pub start: u32,
    pub end: u32,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PrintTitles {
    pub repeat_rows: Option<PrintRowRange>,
    pub repeat_cols: Option<PrintColRange>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PageMargins {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(tag = "kind")]
pub enum PageScaling {
    #[serde(rename = "percent")]
    Percent { percent: u16 },
    #[serde(rename = "fitTo")]
    FitTo { width_pages: u16, height_pages: u16 },
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PageSetup {
    pub orientation: PageOrientation,
    pub paper_size: u16,
    pub margins: PageMargins,
    pub scaling: PageScaling,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct ManualPageBreaks {
    pub row_breaks_after: Vec<u32>,
    pub col_breaks_after: Vec<u32>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct SheetPrintSettings {
    pub sheet_name: String,
    pub print_area: Option<Vec<PrintCellRange>>,
    pub print_titles: Option<PrintTitles>,
    pub page_setup: PageSetup,
    pub manual_page_breaks: ManualPageBreaks,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CellValue {
    pub value: Option<JsonValue>,
    pub formula: Option<String>,
    pub display_value: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CellUpdate {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
    pub value: Option<JsonValue>,
    pub formula: Option<String>,
    pub display_value: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct RangeData {
    pub values: Vec<Vec<CellValue>>,
    pub start_row: usize,
    pub start_col: usize,
}

/// Excel-compatible sheet visibility values used over IPC.
///
/// This intentionally uses camelCase serialization to match the frontend contract
/// (`"veryHidden"`).
#[derive(Clone, Copy, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
}

impl From<ModelSheetVisibility> for SheetVisibility {
    fn from(value: ModelSheetVisibility) -> Self {
        match value {
            ModelSheetVisibility::Visible => SheetVisibility::Visible,
            ModelSheetVisibility::Hidden => SheetVisibility::Hidden,
            ModelSheetVisibility::VeryHidden => SheetVisibility::VeryHidden,
        }
    }
}

impl From<SheetVisibility> for ModelSheetVisibility {
    fn from(value: SheetVisibility) -> Self {
        match value {
            SheetVisibility::Visible => ModelSheetVisibility::Visible,
            SheetVisibility::Hidden => ModelSheetVisibility::Hidden,
            SheetVisibility::VeryHidden => ModelSheetVisibility::VeryHidden,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetInfo {
    pub id: String,
    pub name: String,
    pub visibility: SheetVisibility,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<TabColor>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct WorkbookInfo {
    pub path: Option<String>,
    pub origin_path: Option<String>,
    pub sheets: Vec<SheetInfo>,
}

#[derive(Clone, Copy, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum EncryptionTypeDto {
    Agile,
    Standard,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct EncryptionSummaryDto {
    pub encryption_type: EncryptionTypeDto,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_algorithm: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub key_bits: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cipher: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub key_size: Option<u32>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct DefinedNameInfo {
    pub name: String,
    pub refers_to: String,
    pub sheet_id: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct TableInfo {
    pub name: String,
    pub sheet_id: String,
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
    pub columns: Vec<String>,
}

/// JSON payload for chart drawing objects imported from an XLSX package.
///
/// This includes:
/// - The drawing anchor (`anchor`) so the frontend can place a chart placeholder on the sheet.
/// - The parsed chart `model` (when available) so the placeholder can be upgraded into a rendered
///   chart via the canvas renderer.
#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct ImportedChartObjectInfo {
    pub chart_id: String,
    /// Relationship id inside the drawing part (`rId*`).
    pub rel_id: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_object_id: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_object_name: Option<String>,
    pub anchor: FormulaDrawingAnchor,
    pub drawing_frame_xml: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub model: Option<FormulaChartModel>,
}

/// JSON payload for embedded-in-cell images imported from an XLSX package.
///
/// These correspond to Excel "Place in Cell" / RichData-backed cell images that are referenced via
/// `c/@vm` and ultimately resolve into `xl/media/*` entries.
#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct ImportedEmbeddedCellImageInfo {
    /// Worksheet part name (e.g. `xl/worksheets/sheet1.xml`).
    pub worksheet_part: String,
    /// Best-effort workbook sheet name for this worksheet part.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_name: Option<String>,
    /// 0-based row index.
    pub row: usize,
    /// 0-based column index.
    pub col: usize,
    /// Stable image id (prefers the file name, e.g. `image1.png`).
    pub image_id: String,
    /// Raw image bytes base64 encoded.
    pub bytes_base64: String,
    /// Best-effort inferred MIME type (e.g. `image/png`).
    pub mime_type: String,
    /// Optional alternative text (if present in the workbook metadata).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<String>,
}

/// JSON payload for worksheet background images imported from an XLSX package.
///
/// Excel stores tiled worksheet background images as `<picture r:id="..."/>` inside the worksheet
/// XML, with the relationship pointing at an image part in `xl/media/...`.
#[cfg(any(feature = "desktop", test))]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct ImportedSheetBackgroundImageInfo {
    pub sheet_name: String,
    /// ZIP entry name for the worksheet XML (e.g. `xl/worksheets/sheet1.xml`).
    pub worksheet_part: String,
    /// Workbook-scoped image id used by the UI image store.
    pub image_id: String,
    /// Base64-encoded raw image bytes.
    pub bytes_base64: String,
    pub mime_type: String,
}

#[cfg(any(feature = "desktop", test))]
#[derive(Clone, Debug, PartialEq)]
struct ImportedSheetBackgroundImagePayload {
    sheet_name: String,
    worksheet_part: String,
    image_id: String,
    mime_type: String,
    bytes: Vec<u8>,
}

#[cfg(any(feature = "desktop", test))]
fn sheet_background_image_payloads_from_preserved_parts(
    workbook: &crate::file_io::Workbook,
) -> Vec<ImportedSheetBackgroundImagePayload> {
    use crate::resource_limits::{
        MAX_IMPORTED_SHEET_BACKGROUND_IMAGES_TOTAL_BYTES, MAX_IMPORTED_SHEET_BACKGROUND_IMAGE_BYTES,
    };

    let Some(preserved) = workbook.preserved_drawing_parts.as_ref() else {
        return Vec::new();
    };

    // Resolve worksheet part names by sheet name so we can resolve relationship targets.
    let mut worksheet_part_by_sheet_name: std::collections::HashMap<&str, &str> =
        std::collections::HashMap::new();
    for sheet in &workbook.sheets {
        if let Some(part) = sheet.xlsx_worksheet_part.as_deref() {
            worksheet_part_by_sheet_name.insert(sheet.name.as_str(), part);
        }
    }

    let mut out = Vec::new();
    let mut total_bytes = 0usize;

    fn strip_uri_suffixes(target: &str) -> &str {
        let target = target.trim();
        let target = target.split_once('#').map(|(t, _)| t).unwrap_or(target);
        target.split_once('?').map(|(t, _)| t).unwrap_or(target)
    }

    fn normalize_part_path(path: &str) -> String {
        let mut out: Vec<&str> = Vec::new();
        for part in path.split('/') {
            match part {
                "" | "." => {}
                ".." => {
                    out.pop();
                }
                other => out.push(other),
            }
        }
        out.join("/")
    }

    fn resolve_part_target(source_part: &str, target: &str) -> String {
        let source_part: std::borrow::Cow<'_, str> = if source_part.contains('\\') {
            std::borrow::Cow::Owned(source_part.replace('\\', "/"))
        } else {
            std::borrow::Cow::Borrowed(source_part)
        };

        let target: std::borrow::Cow<'_, str> = if target.contains('\\') {
            std::borrow::Cow::Owned(target.replace('\\', "/"))
        } else {
            std::borrow::Cow::Borrowed(target)
        };

        let target = strip_uri_suffixes(target.as_ref());
        if target.is_empty() {
            return normalize_part_path(source_part.as_ref());
        }
        if let Some(target) = target.strip_prefix('/') {
            return normalize_part_path(target);
        }
        let base_dir = source_part
            .rsplit_once('/')
            .map(|(dir, _)| dir)
            .unwrap_or("");
        normalize_part_path(&format!("{base_dir}/{target}"))
    }

    for (sheet_name, picture) in preserved.sheet_pictures.iter() {
        let Some(worksheet_part) = worksheet_part_by_sheet_name.get(sheet_name.as_str()) else {
            continue;
        };

        let target_part = resolve_part_target(worksheet_part, picture.picture_rel.target.as_str());
        if target_part.trim().is_empty() {
            continue;
        }

        // Prefer filename-only ids to match other in-app image sources.
        let image_id = target_part
            .rsplit('/')
            .next()
            .unwrap_or(target_part.as_str())
            .to_string();
        if image_id.trim().is_empty() {
            continue;
        }

        let Some(bytes) = preserved.parts.get(&target_part) else {
            continue;
        };

        let byte_len = bytes.len();
        if byte_len == 0 {
            continue;
        }
        if byte_len > MAX_IMPORTED_SHEET_BACKGROUND_IMAGE_BYTES {
            continue;
        }
        if total_bytes.saturating_add(byte_len) > MAX_IMPORTED_SHEET_BACKGROUND_IMAGES_TOTAL_BYTES {
            // Best-effort: stop once we've hit the aggregate cap.
            break;
        }
        total_bytes = total_bytes.saturating_add(byte_len);

        let mime_type = mime_guess::from_path(&image_id)
            .first_or_octet_stream()
            .essence_str()
            .to_string();

        out.push(ImportedSheetBackgroundImagePayload {
            sheet_name: sheet_name.clone(),
            worksheet_part: (*worksheet_part).to_string(),
            image_id,
            mime_type,
            bytes: bytes.clone(),
        });
    }

    out
}

#[cfg(any(feature = "desktop", test))]
fn imported_sheet_background_images_from_preserved_payloads(
    payloads: Vec<ImportedSheetBackgroundImagePayload>,
) -> Vec<ImportedSheetBackgroundImageInfo> {
    use base64::{engine::general_purpose::STANDARD, Engine as _};
    payloads
        .into_iter()
        .map(|p| ImportedSheetBackgroundImageInfo {
            sheet_name: p.sheet_name,
            worksheet_part: p.worksheet_part,
            image_id: p.image_id,
            bytes_base64: STANDARD.encode(p.bytes),
            mime_type: p.mime_type,
        })
        .collect()
}

/// JSON payload for DrawingML objects + images imported from an XLSX package.
///
/// This is used to hydrate the desktop DocumentController with floating worksheet drawings
/// (pictures/shapes/chart placeholders) and their referenced image bytes so the drawings overlay
/// can render real imported workbook drawings.
#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct ImportedDrawingObjectsSheetEntry {
    pub sheet_name: String,
    pub sheet_part: String,
    pub drawing_part: String,
    pub objects: Vec<formula_model::drawings::DrawingObject>,
}

/// Image payload matching the DocumentController snapshot schema (`snapshot.images`).
#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ImportedWorkbookImageEntry {
    pub id: String,
    pub bytes_base64: String,
    pub mime_type: Option<String>,
}

#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct ImportedDrawingLayerPayload {
    pub drawings: Vec<ImportedDrawingObjectsSheetEntry>,
    pub images: Vec<ImportedWorkbookImageEntry>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct SheetUsedRange {
    pub start_row: usize,
    pub end_row: usize,
    pub start_col: usize,
    pub end_col: usize,
}

/// Per-sheet view/layout overrides imported from the source workbook.
///
/// Currently this only exposes column widths + hidden columns needed for Excel-compatible
/// `CELL("width")` semantics.
#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct SheetViewOverrides {
    /// Sparse column width overrides in Excel "character" units (0-based column index).
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub col_widths: BTreeMap<u32, f32>,
    /// 0-based indices for user-hidden columns.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hidden_cols: Vec<u32>,
}

#[cfg(any(feature = "desktop", test))]
const SHEET_FORMATTING_METADATA_KEY: &str = "formula_ui_formatting";
#[cfg(any(feature = "desktop", test))]
const SHEET_FORMATTING_SCHEMA_VERSION: i64 = 1;

#[cfg(any(feature = "desktop", test))]
fn json_within_byte_limit(value: &JsonValue, max_bytes: usize, what: &str) -> Result<(), String> {
    struct ByteLimitWriter {
        remaining: usize,
    }

    impl std::io::Write for ByteLimitWriter {
        fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
            if buf.len() > self.remaining {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::Other,
                    "payload too large",
                ));
            }
            self.remaining -= buf.len();
            Ok(buf.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    let mut writer = ByteLimitWriter {
        remaining: max_bytes,
    };
    match serde_json::to_writer(&mut writer, value) {
        Ok(()) => Ok(()),
        Err(err) if err.is_io() => Err(format!("{what} is too large (max {max_bytes} bytes)")),
        Err(_) => Err(format!("Failed to serialize {what} as JSON")),
    }
}

#[cfg(any(feature = "desktop", test))]
fn validate_sheet_formatting_metadata_size(value: &JsonValue) -> Result<(), String> {
    use crate::resource_limits::MAX_SHEET_FORMATTING_METADATA_BYTES;
    json_within_byte_limit(
        value,
        MAX_SHEET_FORMATTING_METADATA_BYTES,
        "Sheet formatting metadata",
    )
}

#[cfg(any(feature = "desktop", test))]
fn default_sheet_formatting_payload() -> JsonValue {
    let mut obj = serde_json::Map::new();
    obj.insert(
        "schemaVersion".to_string(),
        JsonValue::from(SHEET_FORMATTING_SCHEMA_VERSION),
    );
    JsonValue::Object(obj)
}

#[cfg(any(feature = "desktop", test))]
fn sheet_formatting_payload_for_ipc(sheet_id: &str, raw: Option<&JsonValue>) -> JsonValue {
    use crate::resource_limits::MAX_SHEET_FORMATTING_METADATA_BYTES;

    let Some(raw) = raw else {
        return default_sheet_formatting_payload();
    };

    if let Err(err) = json_within_byte_limit(
        raw,
        MAX_SHEET_FORMATTING_METADATA_BYTES,
        "Sheet formatting metadata",
    ) {
        eprintln!("[sheet formatting] {err} for sheet {sheet_id}; returning default formatting");
        return default_sheet_formatting_payload();
    }

    raw.clone()
}

#[cfg(feature = "desktop")]
const SHEET_VIEW_METADATA_KEY: &str = "formula_ui_sheet_view";
#[cfg(feature = "desktop")]
const SHEET_VIEW_SCHEMA_VERSION: i64 = 1;

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetRowFormatDelta {
    pub row: i64,
    pub format: JsonValue,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetColFormatDelta {
    pub col: i64,
    pub format: JsonValue,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetCellFormatDelta {
    pub row: i64,
    pub col: i64,
    pub format: JsonValue,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetFormatRunDelta {
    pub start_row: i64,
    pub end_row_exclusive: i64,
    pub format: JsonValue,
}

/// IPC-deserialized list of `SheetFormatRunDelta` with a maximum length.
///
/// This prevents a compromised webview from sending arbitrarily large `runs` arrays and forcing the
/// backend to allocate an unbounded `Vec` during deserialization.
#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(transparent)]
pub struct LimitedSheetFormatRunDeltas(pub Vec<SheetFormatRunDelta>);

impl<'de> Deserialize<'de> for LimitedSheetFormatRunDeltas {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedSheetFormatRunDeltas;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of format runs")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_SHEET_FORMATTING_RUNS_PER_COL;

                if let Some(hint) = seq.size_hint() {
                    if hint > MAX_SHEET_FORMATTING_RUNS_PER_COL {
                        return Err(de::Error::custom(format!(
                            "formatRunsByCol[].runs is too large (max {MAX_SHEET_FORMATTING_RUNS_PER_COL} runs per column)"
                        )));
                    }
                }

                let mut out = Vec::new();
                while let Some(v) = seq.next_element::<SheetFormatRunDelta>()? {
                    if out.len() >= MAX_SHEET_FORMATTING_RUNS_PER_COL {
                        return Err(de::Error::custom(format!(
                            "formatRunsByCol[].runs is too large (max {MAX_SHEET_FORMATTING_RUNS_PER_COL} runs per column)"
                        )));
                    }
                    out.push(v);
                }
                Ok(LimitedSheetFormatRunDeltas(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetFormatRunsByColDelta {
    pub col: i64,
    pub runs: LimitedSheetFormatRunDeltas,
}

/// IPC-deserialized list of `SheetRowFormatDelta` with a maximum length.
#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(transparent)]
pub struct LimitedSheetRowFormatDeltas(pub Vec<SheetRowFormatDelta>);

impl<'de> Deserialize<'de> for LimitedSheetRowFormatDeltas {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedSheetRowFormatDeltas;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of row formatting deltas")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_SHEET_FORMATTING_ROW_DELTAS;

                if let Some(hint) = seq.size_hint() {
                    if hint > MAX_SHEET_FORMATTING_ROW_DELTAS {
                        return Err(de::Error::custom(format!(
                            "rowFormats is too large (max {MAX_SHEET_FORMATTING_ROW_DELTAS} deltas)"
                        )));
                    }
                }

                let mut out = Vec::new();
                while let Some(v) = seq.next_element::<SheetRowFormatDelta>()? {
                    if out.len() >= MAX_SHEET_FORMATTING_ROW_DELTAS {
                        return Err(de::Error::custom(format!(
                            "rowFormats is too large (max {MAX_SHEET_FORMATTING_ROW_DELTAS} deltas)"
                        )));
                    }
                    out.push(v);
                }
                Ok(LimitedSheetRowFormatDeltas(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

/// IPC-deserialized list of `SheetColFormatDelta` with a maximum length.
#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(transparent)]
pub struct LimitedSheetColFormatDeltas(pub Vec<SheetColFormatDelta>);

impl<'de> Deserialize<'de> for LimitedSheetColFormatDeltas {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedSheetColFormatDeltas;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of column formatting deltas")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_SHEET_FORMATTING_COL_DELTAS;

                if let Some(hint) = seq.size_hint() {
                    if hint > MAX_SHEET_FORMATTING_COL_DELTAS {
                        return Err(de::Error::custom(format!(
                            "colFormats is too large (max {MAX_SHEET_FORMATTING_COL_DELTAS} deltas)"
                        )));
                    }
                }

                let mut out = Vec::new();
                while let Some(v) = seq.next_element::<SheetColFormatDelta>()? {
                    if out.len() >= MAX_SHEET_FORMATTING_COL_DELTAS {
                        return Err(de::Error::custom(format!(
                            "colFormats is too large (max {MAX_SHEET_FORMATTING_COL_DELTAS} deltas)"
                        )));
                    }
                    out.push(v);
                }
                Ok(LimitedSheetColFormatDeltas(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

/// IPC-deserialized list of `SheetCellFormatDelta` with a maximum length.
#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(transparent)]
pub struct LimitedSheetCellFormatDeltas(pub Vec<SheetCellFormatDelta>);

impl<'de> Deserialize<'de> for LimitedSheetCellFormatDeltas {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedSheetCellFormatDeltas;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of cell formatting deltas")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_SHEET_FORMATTING_CELL_DELTAS;

                if let Some(hint) = seq.size_hint() {
                    if hint > MAX_SHEET_FORMATTING_CELL_DELTAS {
                        return Err(de::Error::custom(format!(
                            "cellFormats is too large (max {MAX_SHEET_FORMATTING_CELL_DELTAS} deltas)"
                        )));
                    }
                }

                let mut out = Vec::new();
                while let Some(v) = seq.next_element::<SheetCellFormatDelta>()? {
                    if out.len() >= MAX_SHEET_FORMATTING_CELL_DELTAS {
                        return Err(de::Error::custom(format!(
                            "cellFormats is too large (max {MAX_SHEET_FORMATTING_CELL_DELTAS} deltas)"
                        )));
                    }
                    out.push(v);
                }
                Ok(LimitedSheetCellFormatDeltas(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

/// IPC-deserialized list of `SheetFormatRunsByColDelta` with a maximum length.
#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(transparent)]
pub struct LimitedSheetFormatRunsByColDeltas(pub Vec<SheetFormatRunsByColDelta>);

impl<'de> Deserialize<'de> for LimitedSheetFormatRunsByColDeltas {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedSheetFormatRunsByColDeltas;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of format-run-by-column deltas")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_SHEET_FORMATTING_RUN_COLS;

                if let Some(hint) = seq.size_hint() {
                    if hint > MAX_SHEET_FORMATTING_RUN_COLS {
                        return Err(de::Error::custom(format!(
                            "formatRunsByCol is too large (max {MAX_SHEET_FORMATTING_RUN_COLS} columns)"
                        )));
                    }
                }

                let mut out = Vec::new();
                while let Some(v) = seq.next_element::<SheetFormatRunsByColDelta>()? {
                    if out.len() >= MAX_SHEET_FORMATTING_RUN_COLS {
                        return Err(de::Error::custom(format!(
                            "formatRunsByCol is too large (max {MAX_SHEET_FORMATTING_RUN_COLS} columns)"
                        )));
                    }
                    out.push(v);
                }
                Ok(LimitedSheetFormatRunsByColDeltas(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ApplySheetFormattingDeltasRequest {
    pub sheet_id: String,
    /// If present: `null` clears the sheet default format; an object sets it.
    #[serde(default)]
    pub default_format: Option<Option<JsonValue>>,
    /// Row formatting deltas; `format: null` clears the override for that row.
    #[serde(default)]
    pub row_formats: Option<LimitedSheetRowFormatDeltas>,
    /// Column formatting deltas; `format: null` clears the override for that col.
    #[serde(default)]
    pub col_formats: Option<LimitedSheetColFormatDeltas>,
    /// Replace range-run formatting for the specified columns (runs are replaced wholesale).
    #[serde(default)]
    pub format_runs_by_col: Option<LimitedSheetFormatRunsByColDeltas>,
    /// Cell formatting deltas; `format: null` clears the override for that cell.
    #[serde(default)]
    pub cell_formats: Option<LimitedSheetCellFormatDeltas>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetColWidthDelta {
    pub col: i64,
    pub width: Option<f64>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetRowHeightDelta {
    pub row: i64,
    pub height: Option<f64>,
}

/// IPC-deserialized list of `SheetColWidthDelta` with a maximum length.
#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(transparent)]
pub struct LimitedSheetColWidthDeltas(pub Vec<SheetColWidthDelta>);

impl<'de> Deserialize<'de> for LimitedSheetColWidthDeltas {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedSheetColWidthDeltas;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of column width deltas")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_SHEET_VIEW_COL_WIDTH_DELTAS;

                if let Some(hint) = seq.size_hint() {
                    if hint > MAX_SHEET_VIEW_COL_WIDTH_DELTAS {
                        return Err(de::Error::custom(format!(
                            "colWidths is too large (max {MAX_SHEET_VIEW_COL_WIDTH_DELTAS} deltas)"
                        )));
                    }
                }

                let mut out = Vec::new();
                for _ in 0..MAX_SHEET_VIEW_COL_WIDTH_DELTAS {
                    match seq.next_element::<SheetColWidthDelta>()? {
                        Some(v) => out.push(v),
                        None => return Ok(LimitedSheetColWidthDeltas(out)),
                    }
                }

                if seq.next_element::<de::IgnoredAny>()?.is_some() {
                    return Err(de::Error::custom(format!(
                        "colWidths is too large (max {MAX_SHEET_VIEW_COL_WIDTH_DELTAS} deltas)"
                    )));
                }

                Ok(LimitedSheetColWidthDeltas(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

/// IPC-deserialized list of `SheetRowHeightDelta` with a maximum length.
#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(transparent)]
pub struct LimitedSheetRowHeightDeltas(pub Vec<SheetRowHeightDelta>);

impl<'de> Deserialize<'de> for LimitedSheetRowHeightDeltas {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedSheetRowHeightDeltas;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of row height deltas")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_SHEET_VIEW_ROW_HEIGHT_DELTAS;

                if let Some(hint) = seq.size_hint() {
                    if hint > MAX_SHEET_VIEW_ROW_HEIGHT_DELTAS {
                        return Err(de::Error::custom(format!(
                            "rowHeights is too large (max {MAX_SHEET_VIEW_ROW_HEIGHT_DELTAS} deltas)"
                        )));
                    }
                }

                let mut out = Vec::new();
                for _ in 0..MAX_SHEET_VIEW_ROW_HEIGHT_DELTAS {
                    match seq.next_element::<SheetRowHeightDelta>()? {
                        Some(v) => out.push(v),
                        None => return Ok(LimitedSheetRowHeightDeltas(out)),
                    }
                }

                if seq.next_element::<de::IgnoredAny>()?.is_some() {
                    return Err(de::Error::custom(format!(
                        "rowHeights is too large (max {MAX_SHEET_VIEW_ROW_HEIGHT_DELTAS} deltas)"
                    )));
                }

                Ok(LimitedSheetRowHeightDeltas(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ApplySheetViewDeltasRequest {
    pub sheet_id: String,
    /// Column width deltas; `width: null` clears the override for that col.
    #[serde(default)]
    pub col_widths: Option<LimitedSheetColWidthDeltas>,
    /// Row height deltas; `height: null` clears the override for that row.
    #[serde(default)]
    pub row_heights: Option<LimitedSheetRowHeightDeltas>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookThemePalette {
    pub dk1: String,
    pub lt1: String,
    pub dk2: String,
    pub lt2: String,
    pub accent1: String,
    pub accent2: String,
    pub accent3: String,
    pub accent4: String,
    pub accent5: String,
    pub accent6: String,
    pub hlink: String,
    pub followed_hlink: String,
}

#[cfg(any(feature = "desktop", test))]
fn rgb_hex(argb: u32) -> String {
    format!("#{:06X}", argb & 0x00FF_FFFF)
}

#[cfg(any(feature = "desktop", test))]
fn workbook_theme_palette(workbook: &crate::file_io::Workbook) -> Option<WorkbookThemePalette> {
    let palette = workbook.theme_palette.as_ref()?;
    Some(WorkbookThemePalette {
        dk1: rgb_hex(palette.dk1),
        lt1: rgb_hex(palette.lt1),
        dk2: rgb_hex(palette.dk2),
        lt2: rgb_hex(palette.lt2),
        accent1: rgb_hex(palette.accent1),
        accent2: rgb_hex(palette.accent2),
        accent3: rgb_hex(palette.accent3),
        accent4: rgb_hex(palette.accent4),
        accent5: rgb_hex(palette.accent5),
        accent6: rgb_hex(palette.accent6),
        hlink: rgb_hex(palette.hlink),
        followed_hlink: rgb_hex(palette.followed_hlink),
    })
}

/// A string wrapper used for IPC inputs that enforces a maximum byte length during deserialization.
///
/// This is defensive: a compromised webview could otherwise send arbitrarily large strings and
/// force the backend to allocate excessive memory while deserializing the command payload.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LimitedString<const MAX_BYTES: usize>(String);

impl<const MAX_BYTES: usize> LimitedString<MAX_BYTES> {
    pub fn into_inner(self) -> String {
        self.0
    }
}

impl<const MAX_BYTES: usize> AsRef<str> for LimitedString<MAX_BYTES> {
    fn as_ref(&self) -> &str {
        &self.0
    }
}

impl<const MAX_BYTES: usize> From<LimitedString<MAX_BYTES>> for String {
    fn from(value: LimitedString<MAX_BYTES>) -> Self {
        value.0
    }
}

impl<const MAX_BYTES: usize> Serialize for LimitedString<MAX_BYTES> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        serializer.serialize_str(&self.0)
    }
}

impl<'de, const MAX_BYTES: usize> Deserialize<'de> for LimitedString<MAX_BYTES> {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct LimitedStringVisitor<const MAX_BYTES: usize>;

        impl<const MAX_BYTES: usize> LimitedStringVisitor<MAX_BYTES> {
            fn validate<E>(value: &str) -> Result<(), E>
            where
                E: de::Error,
            {
                if value.len() > MAX_BYTES {
                    return Err(E::custom(format!(
                        "string is too large (max {MAX_BYTES} bytes)"
                    )));
                }
                Ok(())
            }
        }

        impl<'de, const MAX_BYTES: usize> de::Visitor<'de> for LimitedStringVisitor<MAX_BYTES> {
            type Value = LimitedString<MAX_BYTES>;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                write!(formatter, "a string (max {MAX_BYTES} bytes)")
            }

            fn visit_borrowed_str<E>(self, v: &'de str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::validate::<E>(v)?;
                Ok(LimitedString(v.to_owned()))
            }

            fn visit_str<E>(self, v: &str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::validate::<E>(v)?;
                Ok(LimitedString(v.to_owned()))
            }

            fn visit_string<E>(self, v: String) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::validate::<E>(&v)?;
                Ok(LimitedString(v))
            }
        }

        deserializer.deserialize_str(LimitedStringVisitor::<MAX_BYTES>)
    }
}

/// IPC string wrapper for script code payloads with a maximum byte length.
///
/// This uses a more user-friendly error message than `LimitedString` when the limit is exceeded
/// since it is surfaced directly to users via the Python/TypeScript execution UI.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LimitedScriptCode<const MAX_BYTES: usize>(String);

impl<const MAX_BYTES: usize> LimitedScriptCode<MAX_BYTES> {
    pub fn into_inner(self) -> String {
        self.0
    }
}

impl<const MAX_BYTES: usize> AsRef<str> for LimitedScriptCode<MAX_BYTES> {
    fn as_ref(&self) -> &str {
        &self.0
    }
}

impl<const MAX_BYTES: usize> Serialize for LimitedScriptCode<MAX_BYTES> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        serializer.serialize_str(&self.0)
    }
}

impl<'de, const MAX_BYTES: usize> Deserialize<'de> for LimitedScriptCode<MAX_BYTES> {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct ScriptCodeVisitor<const MAX_BYTES: usize>;

        impl<const MAX_BYTES: usize> ScriptCodeVisitor<MAX_BYTES> {
            fn validate<E>(value: &str) -> Result<(), E>
            where
                E: de::Error,
            {
                if value.len() > MAX_BYTES {
                    return Err(E::custom(format!(
                        "Script is too large (max {MAX_BYTES} bytes)"
                    )));
                }
                Ok(())
            }
        }

        impl<'de, const MAX_BYTES: usize> de::Visitor<'de> for ScriptCodeVisitor<MAX_BYTES> {
            type Value = LimitedScriptCode<MAX_BYTES>;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                write!(formatter, "a script string (max {MAX_BYTES} bytes)")
            }

            fn visit_borrowed_str<E>(self, v: &'de str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::validate::<E>(v)?;
                Ok(LimitedScriptCode(v.to_owned()))
            }

            fn visit_str<E>(self, v: &str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::validate::<E>(v)?;
                Ok(LimitedScriptCode(v.to_owned()))
            }

            fn visit_string<E>(self, v: String) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::validate::<E>(&v)?;
                Ok(LimitedScriptCode(v))
            }
        }

        deserializer.deserialize_str(ScriptCodeVisitor::<MAX_BYTES>)
    }
}

/// Alias used in Tauri IPC command signatures for script code.
///
/// The source-level IPC origin guardrail tests locate function bodies by the first `{` after the
/// `pub fn ...` signature. Keeping the const generic expression behind a type alias avoids
/// confusing that heuristic.
pub type IpcScriptCode = LimitedScriptCode<{ crate::ipc_limits::MAX_SCRIPT_CODE_BYTES }>;
/// IPC string wrapper for cell formula payloads with a maximum byte length.
///
/// This uses a cell-edit-specific error message so clients can distinguish formula-size failures
/// from other IPC string limit failures.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LimitedCellFormula(String);

impl LimitedCellFormula {
    pub fn into_inner(self) -> String {
        self.0
    }
}

impl AsRef<str> for LimitedCellFormula {
    fn as_ref(&self) -> &str {
        &self.0
    }
}

impl From<LimitedCellFormula> for String {
    fn from(value: LimitedCellFormula) -> Self {
        value.0
    }
}

impl Serialize for LimitedCellFormula {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        serializer.serialize_str(self.as_ref())
    }
}

impl<'de> Deserialize<'de> for LimitedCellFormula {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct FormulaVisitor;

        impl FormulaVisitor {
            fn validate<E>(value: &str) -> Result<(), E>
            where
                E: de::Error,
            {
                if value.len() > MAX_CELL_FORMULA_BYTES {
                    return Err(E::custom(format!(
                        "cell formula is too large (max {MAX_CELL_FORMULA_BYTES} bytes)"
                    )));
                }
                Ok(())
            }
        }

        impl<'de> de::Visitor<'de> for FormulaVisitor {
            type Value = LimitedCellFormula;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                write!(
                    formatter,
                    "a cell formula string (max {MAX_CELL_FORMULA_BYTES} bytes)"
                )
            }

            fn visit_borrowed_str<E>(self, v: &'de str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::validate::<E>(v)?;
                Ok(LimitedCellFormula(v.to_owned()))
            }

            fn visit_str<E>(self, v: &str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::validate::<E>(v)?;
                Ok(LimitedCellFormula(v.to_owned()))
            }

            fn visit_string<E>(self, v: String) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::validate::<E>(&v)?;
                Ok(LimitedCellFormula(v))
            }
        }

        deserializer.deserialize_str(FormulaVisitor)
    }
}

/// IPC-only cell value type that only accepts scalar JSON values.
///
/// This rejects arrays/objects immediately during deserialization to avoid allocating or
/// materializing deeply nested JSON structures.
#[derive(Clone, Debug, PartialEq)]
pub enum LimitedCellValue {
    Null,
    Bool(bool),
    Number(f64),
    String(LimitedString<MAX_CELL_VALUE_STRING_BYTES>),
}

impl LimitedCellValue {
    pub fn into_json(self) -> Option<JsonValue> {
        match self {
            LimitedCellValue::Null => None,
            LimitedCellValue::Bool(b) => Some(JsonValue::Bool(b)),
            LimitedCellValue::Number(n) => Some(JsonValue::from(n)),
            LimitedCellValue::String(s) => Some(JsonValue::String(s.into_inner())),
        }
    }
}

impl Serialize for LimitedCellValue {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        match self {
            LimitedCellValue::Null => serializer.serialize_unit(),
            LimitedCellValue::Bool(v) => serializer.serialize_bool(*v),
            LimitedCellValue::Number(v) => serializer.serialize_f64(*v),
            LimitedCellValue::String(v) => serializer.serialize_str(v.as_ref()),
        }
    }
}

impl<'de> Deserialize<'de> for LimitedCellValue {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct CellValueVisitor;

        impl<'de> de::Visitor<'de> for CellValueVisitor {
            type Value = LimitedCellValue;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("a scalar JSON value (null, boolean, number, or string)")
            }

            fn visit_unit<E>(self) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Ok(LimitedCellValue::Null)
            }

            fn visit_none<E>(self) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Ok(LimitedCellValue::Null)
            }

            fn visit_bool<E>(self, v: bool) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Ok(LimitedCellValue::Bool(v))
            }

            fn visit_i64<E>(self, v: i64) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Ok(LimitedCellValue::Number(v as f64))
            }

            fn visit_u64<E>(self, v: u64) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Ok(LimitedCellValue::Number(v as f64))
            }

            fn visit_f64<E>(self, v: f64) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Ok(LimitedCellValue::Number(v))
            }

            fn visit_borrowed_str<E>(self, v: &'de str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                if v.len() > MAX_CELL_VALUE_STRING_BYTES {
                    return Err(E::custom(format!(
                        "cell value string is too large (max {MAX_CELL_VALUE_STRING_BYTES} bytes)"
                    )));
                }
                Ok(LimitedCellValue::String(LimitedString(v.to_owned())))
            }

            fn visit_str<E>(self, v: &str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                if v.len() > MAX_CELL_VALUE_STRING_BYTES {
                    return Err(E::custom(format!(
                        "cell value string is too large (max {MAX_CELL_VALUE_STRING_BYTES} bytes)"
                    )));
                }
                Ok(LimitedCellValue::String(LimitedString(v.to_owned())))
            }

            fn visit_string<E>(self, v: String) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                if v.len() > MAX_CELL_VALUE_STRING_BYTES {
                    return Err(E::custom(format!(
                        "cell value string is too large (max {MAX_CELL_VALUE_STRING_BYTES} bytes)"
                    )));
                }
                Ok(LimitedCellValue::String(LimitedString(v)))
            }

            fn visit_seq<A>(self, _seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                Err(de::Error::custom(
                    "cell value must be a scalar (null, boolean, number, or string), not an array",
                ))
            }

            fn visit_map<A>(self, _map: A) -> Result<Self::Value, A::Error>
            where
                A: de::MapAccess<'de>,
            {
                Err(de::Error::custom(
                    "cell value must be a scalar (null, boolean, number, or string), not an object",
                ))
            }
        }

        deserializer.deserialize_any(CellValueVisitor)
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(deny_unknown_fields)]
pub struct RangeCellEdit {
    pub value: Option<LimitedCellValue>,
    pub formula: Option<LimitedCellFormula>,
}

/// IPC-deserialized matrix with size limits applied during deserialization.
///
/// This prevents a compromised webview from sending arbitrarily large 2D payloads and forcing the
/// backend to allocate excessive memory before we can apply additional validation.
#[derive(Clone, Debug, PartialEq)]
pub struct LimitedMatrix<T, const MAX_DIM: usize, const MAX_CELLS: usize>(pub Vec<Vec<T>>);

impl<'de, T, const MAX_DIM: usize, const MAX_CELLS: usize> Deserialize<'de>
    for LimitedMatrix<T, MAX_DIM, MAX_CELLS>
where
    T: Deserialize<'de>,
{
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct MatrixVisitor<T, const MAX_DIM: usize, const MAX_CELLS: usize>(PhantomData<T>);

        impl<'de, T, const MAX_DIM: usize, const MAX_CELLS: usize> de::Visitor<'de>
            for MatrixVisitor<T, MAX_DIM, MAX_CELLS>
        where
            T: Deserialize<'de>,
        {
            type Value = LimitedMatrix<T, MAX_DIM, MAX_CELLS>;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("a 2D array")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                struct RowSeed<'a, T, const MAX_DIM: usize, const MAX_CELLS: usize> {
                    total_cells: &'a mut usize,
                    marker: PhantomData<T>,
                }

                impl<'de, T, const MAX_DIM: usize, const MAX_CELLS: usize> de::DeserializeSeed<'de>
                    for RowSeed<'_, T, MAX_DIM, MAX_CELLS>
                where
                    T: Deserialize<'de>,
                {
                    type Value = Vec<T>;

                    fn deserialize<D>(self, deserializer: D) -> Result<Self::Value, D::Error>
                    where
                        D: serde::Deserializer<'de>,
                    {
                        struct RowVisitor<'a, T, const MAX_DIM: usize, const MAX_CELLS: usize> {
                            total_cells: &'a mut usize,
                            marker: PhantomData<T>,
                        }

                        impl<'de, T, const MAX_DIM: usize, const MAX_CELLS: usize> de::Visitor<'de>
                            for RowVisitor<'_, T, MAX_DIM, MAX_CELLS>
                        where
                            T: Deserialize<'de>,
                        {
                            type Value = Vec<T>;

                            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                                formatter.write_str("an array")
                            }

                            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
                            where
                                A: de::SeqAccess<'de>,
                            {
                                let mut row = Vec::new();
                                while row.len() < MAX_DIM {
                                    // Avoid deserializing an element beyond the total cell limit.
                                    // This is important even when `T` itself is bounded, since we'd
                                    // rather fail fast without allocating/validating another item.
                                    if *self.total_cells >= MAX_CELLS {
                                        if seq.next_element::<de::IgnoredAny>()?.is_some() {
                                            return Err(de::Error::custom(format!(
                                                "range values payload is too large (max {MAX_CELLS} cells)"
                                            )));
                                        }
                                        return Ok(row);
                                    }

                                    match seq.next_element::<T>()? {
                                        Some(cell) => {
                                            *self.total_cells = self.total_cells.saturating_add(1);
                                            if *self.total_cells > MAX_CELLS {
                                                return Err(de::Error::custom(format!(
                                                    "range values payload is too large (max {MAX_CELLS} cells)"
                                                )));
                                            }

                                            row.push(cell);
                                        }
                                        None => return Ok(row),
                                    }
                                }

                                // Reject extra columns without fully deserializing them.
                                if seq.next_element::<de::IgnoredAny>()?.is_some() {
                                    return Err(de::Error::custom(format!(
                                        "range values row is too large (max {MAX_DIM} columns)"
                                    )));
                                }
                                Ok(row)
                            }
                        }

                        deserializer.deserialize_seq(RowVisitor::<T, MAX_DIM, MAX_CELLS> {
                            total_cells: self.total_cells,
                            marker: PhantomData,
                        })
                    }
                }

                let mut rows = Vec::new();
                let mut total_cells = 0usize;
                while rows.len() < MAX_DIM {
                    match seq.next_element_seed(RowSeed::<T, MAX_DIM, MAX_CELLS> {
                        total_cells: &mut total_cells,
                        marker: PhantomData,
                    })? {
                        Some(row) => rows.push(row),
                        None => return Ok(LimitedMatrix(rows)),
                    }
                }

                // Reject extra rows without fully deserializing them (which could allocate large
                // values/formulas for a payload we are going to reject anyway).
                if seq.next_element::<de::IgnoredAny>()?.is_some() {
                    return Err(de::Error::custom(format!(
                        "range values payload is too large (max {MAX_DIM} rows)"
                    )));
                }
                Ok(LimitedMatrix(rows))
            }
        }

        deserializer.deserialize_seq(MatrixVisitor::<T, MAX_DIM, MAX_CELLS>(PhantomData))
    }
}

/// IPC-deserialized matrix of cell edits with size limits applied during deserialization.
///
/// This prevents a compromised webview from sending arbitrarily large `values` payloads to the
/// `set_range` command and forcing the backend to allocate excessive memory before we can apply
/// range-size checks.
pub type LimitedRangeCellEdits = LimitedMatrix<
    RangeCellEdit,
    { crate::resource_limits::MAX_RANGE_DIM },
    { crate::resource_limits::MAX_RANGE_CELLS_PER_CALL },
>;

/// IPC-deserialized vector of `f64` with a maximum length.
///
/// Used for PDF export inputs (`col_widths_points`, `row_heights_points`) to avoid unbounded IPC
/// allocations from a compromised webview.
#[derive(Clone, Debug, PartialEq)]
pub struct LimitedF64Vec(pub Vec<f64>);

impl<'de> Deserialize<'de> for LimitedF64Vec {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedF64Vec;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of numbers")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_RANGE_DIM;

                let mut out = Vec::new();
                while let Some(v) = seq.next_element::<f64>()? {
                    if out.len() >= MAX_RANGE_DIM {
                        return Err(de::Error::custom(format!(
                            "array is too large (max {MAX_RANGE_DIM} items)"
                        )));
                    }
                    out.push(v);
                }
                Ok(LimitedF64Vec(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

/// IPC-deserialized vector with a maximum length.
///
/// Used for command arguments that accept arrays from an untrusted webview to avoid allocating
/// arbitrarily large `Vec`s during deserialization.
#[derive(Clone, Debug, PartialEq)]
pub struct LimitedVec<T, const MAX_LEN: usize>(pub Vec<T>);

impl<T, const MAX_LEN: usize> LimitedVec<T, MAX_LEN> {
    pub fn into_inner(self) -> Vec<T> {
        self.0
    }
}

impl<'de, T, const MAX_LEN: usize> Deserialize<'de> for LimitedVec<T, MAX_LEN>
where
    T: Deserialize<'de>,
{
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct LimitedVecVisitor<T, const MAX_LEN: usize>(PhantomData<T>);

        impl<'de, T, const MAX_LEN: usize> de::Visitor<'de> for LimitedVecVisitor<T, MAX_LEN>
        where
            T: Deserialize<'de>,
        {
            type Value = LimitedVec<T, MAX_LEN>;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                let mut out = match seq.size_hint() {
                    Some(hint) => Vec::with_capacity(hint.min(MAX_LEN)),
                    None => Vec::new(),
                };

                for _ in 0..MAX_LEN {
                    match seq.next_element::<T>()? {
                        Some(v) => out.push(v),
                        None => return Ok(LimitedVec(out)),
                    }
                }

                if seq.next_element::<de::IgnoredAny>()?.is_some() {
                    return Err(de::Error::custom(format!(
                        "array is too large (max {MAX_LEN} items)"
                    )));
                }

                Ok(LimitedVec(out))
            }
        }

        deserializer.deserialize_seq(LimitedVecVisitor::<T, MAX_LEN>(PhantomData))
    }
}

/// IPC-specific pivot types with backend-enforced resource limits.
///
/// `formula_engine::pivot::PivotConfig` contains several unbounded collections (`Vec`, `HashSet`,
/// nested `Vec`s, etc). A compromised WebView could send a huge config payload over IPC, forcing the
/// backend to allocate large amounts of memory during deserialization (or spend significant CPU
/// validating/processing pivots). These IPC-only mirror types apply conservative size limits at
/// deserialization time and are converted to the core `PivotConfig` after validation.
pub type PivotText = LimitedString<{ crate::resource_limits::MAX_PIVOT_TEXT_BYTES }>;

/// IPC-friendly mirror of `formula_engine::pivot::PivotKeyPart` with resource limits applied.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(tag = "type", content = "value", rename_all = "camelCase")]
pub enum IpcPivotKeyPart {
    Blank,
    Number(u64),
    Date(sqlx::types::chrono::NaiveDate),
    Text(PivotText),
    Bool(bool),
}

impl From<IpcPivotKeyPart> for PivotKeyPart {
    fn from(value: IpcPivotKeyPart) -> Self {
        match value {
            IpcPivotKeyPart::Blank => PivotKeyPart::Blank,
            IpcPivotKeyPart::Number(bits) => PivotKeyPart::Number(bits),
            IpcPivotKeyPart::Date(date) => PivotKeyPart::Date(date),
            IpcPivotKeyPart::Text(text) => PivotKeyPart::Text(text.into_inner()),
            IpcPivotKeyPart::Bool(v) => PivotKeyPart::Bool(v),
        }
    }
}

/// IPC-friendly mirror of `formula_engine::pivot::PivotField` with resource limits applied.
#[derive(Clone, Debug, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct IpcPivotField {
    pub source_field: IpcPivotFieldRef,
    #[serde(default)]
    pub sort_order: SortOrder,
    #[serde(default)]
    pub manual_sort: Option<
        LimitedVec<IpcPivotKeyPart, { crate::resource_limits::MAX_PIVOT_MANUAL_SORT_ITEMS }>,
    >,
}

/// IPC-friendly mirror of `formula_engine::pivot::PivotFieldRef` with resource limits applied.
#[derive(Clone, Debug, PartialEq)]
pub enum IpcPivotFieldRef {
    /// Backward-compatible worksheet/cache field name.
    Text(PivotText),
    /// Structured Data Model column ref.
    Column { table: PivotText, column: PivotText },
    /// Structured Data Model measure ref.
    Measure { measure: PivotText },
    /// Alternate structured Data Model measure shape (matches `formula_model` schema).
    MeasureName { name: PivotText },
}

impl<'de> Deserialize<'de> for IpcPivotFieldRef {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        use serde::de::IntoDeserializer;

        struct PivotFieldRefVisitor;

        impl PivotFieldRefVisitor {
            fn parse_text<E>(value: &str) -> Result<IpcPivotFieldRef, E>
            where
                E: de::Error,
            {
                let text = PivotText::deserialize(value.into_deserializer())?;
                Ok(IpcPivotFieldRef::Text(text))
            }
        }
        impl<'de> de::Visitor<'de> for PivotFieldRefVisitor {
            type Value = IpcPivotFieldRef;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("a pivot field ref (string or object)")
            }

            fn visit_borrowed_str<E>(self, v: &'de str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::parse_text(v)
            }

            fn visit_str<E>(self, v: &str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                Self::parse_text(v)
            }

            fn visit_string<E>(self, v: String) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                let text = PivotText::deserialize(v.into_deserializer())?;
                Ok(IpcPivotFieldRef::Text(text))
            }

            fn visit_map<A>(self, mut map: A) -> Result<Self::Value, A::Error>
            where
                A: de::MapAccess<'de>,
            {
                let mut table: Option<PivotText> = None;
                let mut column: Option<PivotText> = None;
                let mut measure: Option<PivotText> = None;
                let mut name: Option<PivotText> = None;

                while let Some(key) = map.next_key::<String>()? {
                    match key.as_str() {
                        "table" => {
                            if table.is_some() {
                                return Err(de::Error::duplicate_field("table"));
                            }
                            table = Some(map.next_value()?);
                        }
                        "column" => {
                            if column.is_some() {
                                return Err(de::Error::duplicate_field("column"));
                            }
                            column = Some(map.next_value()?);
                        }
                        "measure" => {
                            if measure.is_some() {
                                return Err(de::Error::duplicate_field("measure"));
                            }
                            measure = Some(map.next_value()?);
                        }
                        "name" => {
                            if name.is_some() {
                                return Err(de::Error::duplicate_field("name"));
                            }
                            name = Some(map.next_value()?);
                        }
                        other => {
                            return Err(de::Error::unknown_field(
                                other,
                                &["table", "column", "measure", "name"],
                            ));
                        }
                    }
                }

                match (table, column, measure, name) {
                    (Some(table), Some(column), None, None) => {
                        Ok(IpcPivotFieldRef::Column { table, column })
                    }
                    (None, None, Some(measure), None) => Ok(IpcPivotFieldRef::Measure { measure }),
                    (None, None, None, Some(name)) => Ok(IpcPivotFieldRef::MeasureName { name }),
                    _ => Err(de::Error::custom(
                        "expected {table,column}, {measure}, or {name} object for pivot field ref",
                    )),
                }
            }
        }

        deserializer.deserialize_any(PivotFieldRefVisitor)
    }
}

impl From<IpcPivotFieldRef> for PivotFieldRef {
    fn from(value: IpcPivotFieldRef) -> Self {
        match value {
            IpcPivotFieldRef::Text(raw) => PivotFieldRef::from_unstructured_owned(raw.into_inner()),
            IpcPivotFieldRef::Column { table, column } => PivotFieldRef::DataModelColumn {
                table: table.into_inner(),
                column: column.into_inner(),
            },
            IpcPivotFieldRef::Measure { measure } => {
                PivotFieldRef::DataModelMeasure(measure.into_inner())
            }
            IpcPivotFieldRef::MeasureName { name } => {
                PivotFieldRef::DataModelMeasure(name.into_inner())
            }
        }
    }
}

impl From<IpcPivotField> for formula_engine::pivot::PivotField {
    fn from(value: IpcPivotField) -> Self {
        Self {
            source_field: value.source_field.into(),
            sort_order: value.sort_order,
            manual_sort: value.manual_sort.map(|v| {
                v.into_inner()
                    .into_iter()
                    .map(PivotKeyPart::from)
                    .collect::<Vec<_>>()
            }),
        }
    }
}

/// IPC-friendly mirror of `formula_engine::pivot::ValueField` with resource limits applied.
#[derive(Clone, Debug, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct IpcValueField {
    pub source_field: IpcPivotFieldRef,
    pub name: PivotText,
    pub aggregation: AggregationType,
    #[serde(default)]
    pub number_format: Option<PivotText>,
    #[serde(default)]
    pub show_as: Option<ShowAsType>,
    #[serde(default)]
    pub base_field: Option<IpcPivotFieldRef>,
    #[serde(default)]
    pub base_item: Option<PivotText>,
}

impl From<IpcValueField> for formula_engine::pivot::ValueField {
    fn from(value: IpcValueField) -> Self {
        Self {
            source_field: value.source_field.into(),
            name: value.name.into_inner(),
            aggregation: value.aggregation,
            number_format: value.number_format.map(|fmt| fmt.into_inner()),
            show_as: value.show_as,
            base_field: value.base_field.map(PivotFieldRef::from),
            base_item: value.base_item.map(|s| s.into_inner()),
        }
    }
}

/// IPC-friendly mirror of `formula_engine::pivot::FilterField` with resource limits applied.
#[derive(Clone, Debug, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct IpcFilterField {
    pub source_field: IpcPivotFieldRef,
    /// Allowed values. `None` means allow all.
    #[serde(default)]
    pub allowed: Option<
        LimitedVec<IpcPivotKeyPart, { crate::resource_limits::MAX_PIVOT_FILTER_ALLOWED_VALUES }>,
    >,
}

impl From<IpcFilterField> for formula_engine::pivot::FilterField {
    fn from(value: IpcFilterField) -> Self {
        Self {
            source_field: value.source_field.into(),
            allowed: value.allowed.map(|vals| {
                vals.into_inner()
                    .into_iter()
                    .map(PivotKeyPart::from)
                    .collect::<HashSet<_>>()
            }),
        }
    }
}

/// IPC-friendly mirror of `formula_engine::pivot::CalculatedField` with resource limits applied.
#[derive(Clone, Debug, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct IpcCalculatedField {
    pub name: PivotText,
    pub formula: PivotText,
}

impl From<IpcCalculatedField> for formula_engine::pivot::CalculatedField {
    fn from(value: IpcCalculatedField) -> Self {
        Self {
            name: value.name.into_inner(),
            formula: value.formula.into_inner(),
        }
    }
}

/// IPC-friendly mirror of `formula_engine::pivot::CalculatedItem` with resource limits applied.
#[derive(Clone, Debug, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct IpcCalculatedItem {
    pub field: PivotText,
    pub name: PivotText,
    pub formula: PivotText,
}

impl From<IpcCalculatedItem> for formula_engine::pivot::CalculatedItem {
    fn from(value: IpcCalculatedItem) -> Self {
        Self {
            field: value.field.into_inner(),
            name: value.name.into_inner(),
            formula: value.formula.into_inner(),
        }
    }
}

/// IPC-friendly mirror of `formula_engine::pivot::PivotConfig` with resource limits applied.
#[derive(Clone, Debug, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct IpcPivotConfig {
    pub row_fields: LimitedVec<IpcPivotField, { crate::resource_limits::MAX_PIVOT_FIELDS }>,
    pub column_fields: LimitedVec<IpcPivotField, { crate::resource_limits::MAX_PIVOT_FIELDS }>,
    pub value_fields: LimitedVec<IpcValueField, { crate::resource_limits::MAX_PIVOT_FIELDS }>,
    pub filter_fields: LimitedVec<IpcFilterField, { crate::resource_limits::MAX_PIVOT_FIELDS }>,
    #[serde(default)]
    pub calculated_fields: Option<
        LimitedVec<IpcCalculatedField, { crate::resource_limits::MAX_PIVOT_CALCULATED_FIELDS }>,
    >,
    #[serde(default)]
    pub calculated_items: Option<
        LimitedVec<IpcCalculatedItem, { crate::resource_limits::MAX_PIVOT_CALCULATED_ITEMS }>,
    >,
    pub layout: Layout,
    pub subtotals: SubtotalPosition,
    pub grand_totals: GrandTotals,
}

impl From<IpcPivotConfig> for PivotConfig {
    fn from(value: IpcPivotConfig) -> Self {
        Self {
            row_fields: value
                .row_fields
                .into_inner()
                .into_iter()
                .map(formula_engine::pivot::PivotField::from)
                .collect(),
            column_fields: value
                .column_fields
                .into_inner()
                .into_iter()
                .map(formula_engine::pivot::PivotField::from)
                .collect(),
            value_fields: value
                .value_fields
                .into_inner()
                .into_iter()
                .map(formula_engine::pivot::ValueField::from)
                .collect(),
            filter_fields: value
                .filter_fields
                .into_inner()
                .into_iter()
                .map(formula_engine::pivot::FilterField::from)
                .collect(),
            calculated_fields: value
                .calculated_fields
                .map(|v| {
                    v.into_inner()
                        .into_iter()
                        .map(formula_engine::pivot::CalculatedField::from)
                        .collect::<Vec<_>>()
                })
                .unwrap_or_default(),
            calculated_items: value
                .calculated_items
                .map(|v| {
                    v.into_inner()
                        .into_iter()
                        .map(formula_engine::pivot::CalculatedItem::from)
                        .collect::<Vec<_>>()
                })
                .unwrap_or_default(),
            layout: value.layout,
            subtotals: value.subtotals,
            grand_totals: value.grand_totals,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PivotCellRange {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PivotDestination {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
}

#[derive(Clone, Debug, Deserialize, PartialEq)]
pub struct CreatePivotTableRequest {
    pub name: PivotText,
    pub source_sheet_id: String,
    pub source_range: PivotCellRange,
    pub destination: PivotDestination,
    pub config: IpcPivotConfig,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CreatePivotTableResponse {
    pub pivot_id: String,
    pub updates: Vec<CellUpdate>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct RefreshPivotTableRequest {
    pub pivot_id: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PivotTableSummary {
    pub id: String,
    pub name: String,
    pub source_sheet_id: String,
    pub source_range: PivotCellRange,
    pub destination: PivotDestination,
}

#[cfg(feature = "desktop")]
use crate::file_io::read_workbook;
#[cfg(any(feature = "desktop", test))]
use crate::ipc_limits::MAX_IPC_PATH_BYTES;
#[cfg(feature = "desktop")]
use crate::ipc_limits::{
    MAX_IPC_COLLAB_ENCRYPTION_KEY_BASE64_BYTES, MAX_IPC_COLLAB_TOKEN_BYTES,
    MAX_IPC_MARKETPLACE_ID_BYTES, MAX_IPC_MARKETPLACE_QUERY_BYTES,
    MAX_IPC_MARKETPLACE_VERSION_BYTES, MAX_IPC_SECURE_STORE_KEY_BYTES, MAX_IPC_URL_BYTES,
    MAX_IPC_WORKBOOK_PASSWORD_BYTES,
};
#[cfg(feature = "desktop")]
use crate::ipc_origin;
#[cfg(feature = "desktop")]
use crate::macro_trust::SharedMacroTrustStore;
#[cfg(feature = "desktop")]
use crate::persistence::{
    autosave_db_path_for_new_workbook, autosave_db_path_for_workbook, WorkbookPersistenceLocation,
};
#[cfg(feature = "desktop")]
use crate::state::SharedAppState;
#[cfg(any(feature = "desktop", test))]
use crate::state::{AppState, AppStateError, CellUpdateData};
#[cfg(any(feature = "desktop", test))]
use crate::{file_io::Workbook, macro_trust::compute_macro_fingerprint};
#[cfg(any(feature = "desktop", test))]
use std::path::PathBuf;
#[cfg(feature = "desktop")]
use tauri::{Emitter, State};
#[cfg(feature = "desktop")]
use tauri_plugin_shell::ShellExt;

#[cfg(feature = "desktop")]
fn app_error(err: AppStateError) -> String {
    err.to_string()
}

#[cfg(any(feature = "desktop", test))]
fn coerce_save_path_to_xlsx(path: &str) -> String {
    let mut buf = PathBuf::from(path);
    let Some(ext) = buf.extension().and_then(|s| s.to_str()) else {
        return path.to_string();
    };

    // We can only write XLSX/XLSM/XLSB bytes. If the workbook was opened from a non-workbook
    // source (CSV/Parquet/etc) or a legacy format that we import into the workbook model (XLS),
    // saving without "Save As" would otherwise write an XLSX file to a non-XLSX filename.
    if ext.eq_ignore_ascii_case("xlsx")
        || ext.eq_ignore_ascii_case("xlsm")
        || ext.eq_ignore_ascii_case("xltx")
        || ext.eq_ignore_ascii_case("xltm")
        || ext.eq_ignore_ascii_case("xlam")
        || ext.eq_ignore_ascii_case("xlsb")
    {
        return path.to_string();
    }

    buf.set_extension("xlsx");
    buf.to_string_lossy().to_string()
}

#[cfg(any(feature = "desktop", test))]
fn wants_origin_bytes_for_save_path(path: &str) -> bool {
    std::path::Path::new(path)
        .extension()
        .and_then(|s| s.to_str())
        .is_some_and(|ext| {
            ext.eq_ignore_ascii_case("xlsx")
                || ext.eq_ignore_ascii_case("xlsm")
                || ext.eq_ignore_ascii_case("xltx")
                || ext.eq_ignore_ascii_case("xltm")
                || ext.eq_ignore_ascii_case("xlam")
        })
}

#[cfg(feature = "desktop")]
fn cell_value_from_state(
    state: &AppState,
    sheet_id: &str,
    row: usize,
    col: usize,
) -> Result<CellValue, String> {
    let cell = state.get_cell(sheet_id, row, col).map_err(app_error)?;
    Ok(CellValue {
        value: cell.value.as_json(),
        formula: cell.formula,
        display_value: cell.display_value,
    })
}

#[cfg(any(feature = "desktop", test))]
fn cell_update_from_state(update: CellUpdateData) -> CellUpdate {
    CellUpdate {
        sheet_id: update.sheet_id,
        row: update.row,
        col: update.col,
        value: update.value.as_json(),
        formula: update.formula,
        display_value: update.display_value,
    }
}

#[cfg(any(feature = "desktop", test))]
fn ole_stream_exists<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    if ole.open_stream(name).is_ok() {
        return true;
    }
    let with_leading_slash = format!("/{name}");
    ole.open_stream(&with_leading_slash).is_ok()
}

#[cfg(any(feature = "desktop", test))]
fn parse_cryptoapi_cipher(alg_id: u32) -> Option<String> {
    // See `ALG_ID` values (e.g. `CALG_AES_128`) in MS-OFFCRYPTO / CryptoAPI.
    let name = match alg_id {
        0x0000_6801 => "RC4",
        0x0000_6603 => "3DES",
        0x0000_660E => "AES-128",
        0x0000_660F => "AES-192",
        0x0000_6610 => "AES-256",
        _ => return None,
    };
    Some(name.to_string())
}

#[cfg(any(feature = "desktop", test))]
fn parse_cryptoapi_hash(alg_id_hash: u32) -> Option<String> {
    let name = match alg_id_hash {
        0x0000_8002 => "MD4",
        0x0000_8003 => "MD5",
        0x0000_8004 => "SHA1",
        0x0000_800C => "SHA256",
        0x0000_800D => "SHA384",
        0x0000_800E => "SHA512",
        _ => return None,
    };
    Some(name.to_string())
}

#[cfg(any(feature = "desktop", test))]
fn is_nul_heavy(bytes: &[u8]) -> bool {
    if bytes.is_empty() {
        return false;
    }
    let zeros = bytes.iter().filter(|&&b| b == 0).count();
    zeros > bytes.len() / 8
}

#[cfg(any(feature = "desktop", test))]
fn strip_utf8_bom(bytes: &[u8]) -> &[u8] {
    bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes)
}

#[cfg(any(feature = "desktop", test))]
fn trim_trailing_nul_bytes(mut bytes: &[u8]) -> &[u8] {
    while let Some((&last, rest)) = bytes.split_last() {
        if last == 0 {
            bytes = rest;
        } else {
            break;
        }
    }
    bytes
}

#[cfg(any(feature = "desktop", test))]
fn looks_like_utf16(bytes: &[u8]) -> bool {
    if bytes.starts_with(&[0xFF, 0xFE]) || bytes.starts_with(&[0xFE, 0xFF]) {
        return true;
    }
    if bytes.len() >= 2
        && ((bytes[0] == b'<' && bytes[1] == 0) || (bytes[0] == 0 && bytes[1] == b'<'))
    {
        return true;
    }
    if !is_nul_heavy(bytes) {
        return false;
    }
    // NUL-heavy payloads are only treated as UTF-16 if the NUL bytes are biased toward a single
    // byte position (as you'd see in UTF-16LE/BE ASCII text).
    let mut zeros_even = 0usize;
    let mut zeros_odd = 0usize;
    for (idx, &b) in bytes.iter().enumerate().take(4096) {
        if b != 0 {
            continue;
        }
        if idx % 2 == 0 {
            zeros_even += 1;
        } else {
            zeros_odd += 1;
        }
    }
    (zeros_even > zeros_odd * 3) || (zeros_odd > zeros_even * 3)
}

#[cfg(any(feature = "desktop", test))]
fn decode_utf16_xml(mut bytes: &[u8]) -> Result<String, String> {
    // Trim trailing UTF-16 NUL terminators / padding.
    while bytes.len() >= 2 {
        let n = bytes.len();
        if bytes[n - 2] == 0 && bytes[n - 1] == 0 {
            bytes = &bytes[..n - 2];
        } else {
            break;
        }
    }

    // Determine endianness:
    // - Prefer BOM if present.
    // - Otherwise, guess based on which byte position has more NUL bytes.
    //   For UTF-16 ASCII text, LE encodes NULs in odd bytes and BE in even bytes.
    let mut big_endian = None;
    if bytes.starts_with(&[0xFF, 0xFE]) {
        bytes = &bytes[2..];
        big_endian = Some(false);
    } else if bytes.starts_with(&[0xFE, 0xFF]) {
        bytes = &bytes[2..];
        big_endian = Some(true);
    } else {
        let mut zeros_even = 0usize;
        let mut zeros_odd = 0usize;
        for (idx, &b) in bytes.iter().enumerate().take(4096) {
            if b != 0 {
                continue;
            }
            if idx % 2 == 0 {
                zeros_even += 1;
            } else {
                zeros_odd += 1;
            }
        }
        if zeros_even > zeros_odd {
            big_endian = Some(true);
        } else if zeros_odd > zeros_even {
            big_endian = Some(false);
        }
    }
    let big_endian = big_endian.unwrap_or(false);

    bytes = &bytes[..bytes.len().saturating_sub(bytes.len() % 2)];
    let mut units = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        units.push(if big_endian {
            u16::from_be_bytes([chunk[0], chunk[1]])
        } else {
            u16::from_le_bytes([chunk[0], chunk[1]])
        });
    }

    let mut xml = String::from_utf16(&units)
        .map_err(|_| "agile EncryptionInfo descriptor is not valid UTF-16".to_string())?;
    if let Some(stripped) = xml.strip_prefix('\u{FEFF}') {
        xml = stripped.to_string();
    }
    while xml.ends_with('\0') {
        xml.pop();
    }
    Ok(xml)
}

#[cfg(any(feature = "desktop", test))]
fn decode_agile_encryption_xml(xml_bytes: &[u8]) -> Result<String, String> {
    // Some producers store the EncryptionInfo XML as UTF-16 (sometimes with a BOM).
    //
    // Important: do **not** trim single trailing NUL bytes before UTF-16 decoding; ASCII UTF-16
    // encodings end in `0x00` for the final code unit.
    if looks_like_utf16(xml_bytes) {
        return decode_utf16_xml(xml_bytes);
    }

    // Primary: treat the remainder as UTF-8 XML (trim UTF-8 BOM, trim trailing NUL padding).
    let utf8_bytes = trim_trailing_nul_bytes(xml_bytes);
    let utf8_bytes = strip_utf8_bom(utf8_bytes);
    if let Ok(xml) = std::str::from_utf8(utf8_bytes) {
        return Ok(xml.strip_prefix('\u{FEFF}').unwrap_or(xml).to_string());
    }

    Err("agile EncryptionInfo descriptor is not valid UTF-8".to_string())
}

#[cfg(any(feature = "desktop", test))]
fn parse_ooxml_encryption_info(bytes: &[u8]) -> Result<EncryptionSummaryDto, String> {
    if bytes.len() < 8 {
        return Err("EncryptionInfo stream is too short".to_string());
    }
    let major = u16::from_le_bytes([bytes[0], bytes[1]]);
    let minor = u16::from_le_bytes([bytes[2], bytes[3]]);

    match (major, minor) {
        (4, 4) => {
            // Agile: version (4 bytes) + flags (4 bytes) + XML encryption descriptor.
            let xml = decode_agile_encryption_xml(&bytes[8..])?;
            let doc = roxmltree::Document::parse(&xml)
                .map_err(|e| format!("agile EncryptionInfo descriptor is not valid XML: {e}"))?;

            let mut hash_algorithm = None;
            let mut spin_count = None;
            let mut key_bits = None;

            for node in doc.descendants().filter(|n| n.is_element()) {
                match node.tag_name().name() {
                    // Password encryption parameters.
                    "encryptedKey" => {
                        if spin_count.is_none() {
                            spin_count = node
                                .attribute("spinCount")
                                .and_then(|raw| raw.parse::<u32>().ok());
                        }
                        if hash_algorithm.is_none() {
                            hash_algorithm =
                                node.attribute("hashAlgorithm").map(|s| s.to_string());
                        }
                        if key_bits.is_none() {
                            key_bits = node
                                .attribute("keyBits")
                                .and_then(|raw| raw.parse::<u32>().ok());
                        }
                    }
                    // Document encryption parameters (also includes `hashAlgorithm` / `keyBits`).
                    "keyData" => {
                        if hash_algorithm.is_none() {
                            hash_algorithm =
                                node.attribute("hashAlgorithm").map(|s| s.to_string());
                        }
                        if key_bits.is_none() {
                            key_bits = node
                                .attribute("keyBits")
                                .and_then(|raw| raw.parse::<u32>().ok());
                        }
                    }
                    _ => {}
                }
            }

            Ok(EncryptionSummaryDto {
                encryption_type: EncryptionTypeDto::Agile,
                hash_algorithm,
                spin_count,
                key_bits,
                cipher: None,
                key_size: None,
            })
        }
        // Standard (CryptoAPI) encryption.
        //
        // MS-OFFCRYPTO identifies Standard encryption via `versionMinor == 2`, but real-world
        // files vary `versionMajor` (commonly 3/4; 2 is also seen).
        (2 | 3 | 4, 2) => {
            const MIN_STANDARD_HEADER_SIZE: usize = 8 * 4;
            const MAX_STANDARD_HEADER_SIZE: usize = 1024 * 1024;

            if bytes.len() < 12 {
                return Err("EncryptionInfo stream is too short for standard encryption".to_string());
            }
            let header_size = u32::from_le_bytes([bytes[8], bytes[9], bytes[10], bytes[11]]) as usize;
            if header_size < MIN_STANDARD_HEADER_SIZE || header_size > MAX_STANDARD_HEADER_SIZE {
                return Err("EncryptionInfo standard header size is out of bounds".to_string());
            }
            let header_start = 12usize;
            let header_end = header_start.saturating_add(header_size);
            if bytes.len() < header_end {
                return Err("EncryptionInfo stream is truncated (header size exceeds stream length)".to_string());
            }
            let header = &bytes[header_start..header_end];
            let alg_id = u32::from_le_bytes([header[8], header[9], header[10], header[11]]);
            let alg_id_hash = u32::from_le_bytes([header[12], header[13], header[14], header[15]]);
            let key_size = u32::from_le_bytes([header[16], header[17], header[18], header[19]]);

            Ok(EncryptionSummaryDto {
                encryption_type: EncryptionTypeDto::Standard,
                hash_algorithm: parse_cryptoapi_hash(alg_id_hash),
                spin_count: None,
                key_bits: None,
                cipher: parse_cryptoapi_cipher(alg_id),
                key_size: (key_size != 0).then_some(key_size),
            })
        }
        _ => Err(format!(
            "Unsupported workbook encryption version {major}.{minor} (expected 4.4 agile or *.2 standard)"
        )),
    }
}

/// Inspect an encrypted OOXML workbook (password-protected `.xlsx`) and extract a summary of the
/// encryption parameters for password prompts.
///
/// Returns:
/// - `Ok(None)` if the workbook does not appear to be an encrypted OOXML container.
/// - `Ok(Some(...))` if the workbook is encrypted and its `EncryptionInfo` stream can be parsed.
/// - `Err(...)` on I/O failures or unsupported/invalid encryption info formats.
#[cfg(any(feature = "desktop", test))]
#[cfg_attr(feature = "desktop", tauri::command)]
pub fn inspect_workbook_encryption(
    path: LimitedString<MAX_IPC_PATH_BYTES>,
) -> Result<Option<EncryptionSummaryDto>, String> {
    use std::io::{Read, Seek};

    const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    const MAX_ENCRYPTION_INFO_BYTES: u64 = 1024 * 1024;

    let path = path.into_inner();
    let allowed_roots = crate::fs_scope::desktop_allowed_roots().map_err(|e| e.to_string())?;
    let resolved =
        crate::fs_scope::canonicalize_in_allowed_roots(std::path::Path::new(&path), &allowed_roots)
            .map_err(|e| e.to_string())?;

    let mut file = std::fs::File::open(&resolved).map_err(|e| e.to_string())?;
    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|e| e.to_string())?;
    if n < OLE_MAGIC.len() || header != OLE_MAGIC {
        return Ok(None);
    }
    file.rewind().map_err(|e| e.to_string())?;

    let mut ole = cfb::CompoundFile::open(file).map_err(|e| e.to_string())?;
    if !(ole_stream_exists(&mut ole, "EncryptionInfo")
        && ole_stream_exists(&mut ole, "EncryptedPackage"))
    {
        return Ok(None);
    }

    let stream = ole
        .open_stream("EncryptionInfo")
        .or_else(|_| ole.open_stream("/EncryptionInfo"))
        .map_err(|e| e.to_string())?;
    let mut bytes = Vec::new();
    stream
        .take(MAX_ENCRYPTION_INFO_BYTES + 1)
        .read_to_end(&mut bytes)
        .map_err(|e| e.to_string())?;
    if bytes.len() as u64 > MAX_ENCRYPTION_INFO_BYTES {
        return Err(format!(
            "EncryptionInfo stream is too large (max {MAX_ENCRYPTION_INFO_BYTES} bytes)"
        ));
    }
    Ok(Some(parse_ooxml_encryption_info(&bytes)?))
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn open_workbook(
    window: tauri::WebviewWindow,
    path: LimitedString<MAX_IPC_PATH_BYTES>,
    password: Option<LimitedString<MAX_IPC_WORKBOOK_PASSWORD_BYTES>>,
    state: State<'_, SharedAppState>,
) -> Result<WorkbookInfo, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "workbook opening", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "workbook opening", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "workbook opening", ipc_origin::Verb::Is)?;

    let path = path.into_inner();
    let password = password.map(|p| p.into_inner());
    let allowed_roots = crate::fs_scope::desktop_allowed_roots().map_err(|e| e.to_string())?;
    let resolved =
        crate::fs_scope::canonicalize_in_allowed_roots(std::path::Path::new(&path), &allowed_roots)
            .map_err(|e| e.to_string())?;
    let resolved_str = resolved.to_string_lossy().to_string();

    let workbook = read_workbook(resolved, password)
        .await
        .map_err(|e| e.to_string())?;
    let location = autosave_db_path_for_workbook(&resolved_str)
        .map(WorkbookPersistenceLocation::OnDisk)
        .unwrap_or(WorkbookPersistenceLocation::InMemory);

    let shared = state.inner().clone();
    let info = tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        state
            .load_workbook_persistent(workbook, location)
            .map_err(app_error)
    })
    .await
    .map_err(|e| e.to_string())?;
    let info = info?;

    Ok(WorkbookInfo {
        path: info.path,
        origin_path: info.origin_path,
        sheets: info
            .sheets
            .into_iter()
            .map(|s| SheetInfo {
                id: s.id,
                name: s.name,
                visibility: s.visibility.into(),
                tab_color: s.tab_color,
            })
            .collect(),
    })
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn new_workbook(state: State<'_, SharedAppState>) -> Result<WorkbookInfo, String> {
    let shared = state.inner().clone();
    let location = autosave_db_path_for_new_workbook()
        .map(WorkbookPersistenceLocation::OnDisk)
        .unwrap_or(WorkbookPersistenceLocation::InMemory);
    let info = tauri::async_runtime::spawn_blocking(move || {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = shared.lock().unwrap();
        state
            .load_workbook_persistent(workbook, location)
            .map_err(app_error)
    })
    .await
    .map_err(|e| e.to_string())?;
    let info = info?;

    Ok(WorkbookInfo {
        path: info.path,
        origin_path: info.origin_path,
        sheets: info
            .sheets
            .into_iter()
            .map(|s| SheetInfo {
                id: s.id,
                name: s.name,
                visibility: s.visibility.into(),
                tab_color: s.tab_color,
            })
            .collect(),
    })
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn add_sheet(
    name: String,
    sheet_id: Option<String>,
    id: Option<String>,
    after_sheet_id: Option<String>,
    index: Option<usize>,
    state: State<'_, SharedAppState>,
) -> Result<SheetInfo, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let sheet_id = sheet_id.or(id);
        let sheet = state
            .add_sheet(name, sheet_id, after_sheet_id, index)
            .map_err(app_error)?;
        Ok::<_, String>(SheetInfo {
            id: sheet.id,
            name: sheet.name,
            visibility: sheet.visibility.into(),
            tab_color: sheet.tab_color,
        })
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn add_sheet_with_id(
    sheet_id: String,
    name: String,
    after_sheet_id: Option<String>,
    index: Option<usize>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        state
            .add_sheet_with_id(sheet_id, name, after_sheet_id, index)
            .map_err(app_error)?;
        Ok::<_, String>(())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn reorder_sheets(
    sheet_ids: LimitedVec<
        LimitedString<{ crate::ipc_limits::MAX_SHEET_ID_BYTES }>,
        { crate::ipc_limits::MAX_REORDER_SHEET_IDS },
    >,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let sheet_ids = sheet_ids
        .into_inner()
        .into_iter()
        .map(|id| id.into_inner())
        .collect();
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        state.reorder_sheets(sheet_ids).map_err(app_error)?;
        Ok::<_, String>(())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn rename_sheet(
    sheet_id: String,
    name: String,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        state.rename_sheet(&sheet_id, name).map_err(app_error)?;
        Ok::<_, String>(())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn set_sheet_visibility(
    sheet_id: String,
    visibility: SheetVisibility,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    // Note: the desktop UI only exposes "visible" / "hidden", but we still accept `veryHidden`
    // from the webview so workbook state reconciliation (applyState/restore) and automation can
    // round-trip Excel-compatible visibility metadata. Backend invariants (e.g. "cannot hide the
    // last visible sheet") are enforced in `AppState::set_sheet_visibility`.
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        set_sheet_visibility_core(&mut state, &sheet_id, visibility).map_err(app_error)?;
        Ok::<_, String>(())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(any(feature = "desktop", test))]
fn set_sheet_visibility_core(
    state: &mut AppState,
    sheet_id: &str,
    visibility: SheetVisibility,
) -> Result<(), AppStateError> {
    state.set_sheet_visibility(sheet_id, visibility.into())
}

#[cfg(any(feature = "desktop", test))]
fn normalize_tab_color_rgb(raw: &str) -> Result<String, String> {
    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return Err("tab color rgb cannot be empty".to_string());
    }
    let hex = trimmed.strip_prefix('#').unwrap_or(trimmed);
    if hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
        return Ok(format!("FF{}", hex.to_ascii_uppercase()));
    }
    if hex.len() == 8 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
        return Ok(hex.to_ascii_uppercase());
    }
    Err("tab color rgb must be 6-digit (RRGGBB) or 8-digit (AARRGGBB) hex".to_string())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn set_sheet_tab_color(
    sheet_id: String,
    tab_color: Option<formula_model::TabColor>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let tab_color = match tab_color {
            None => None,
            Some(mut color) => {
                if let Some(rgb) = color.rgb.as_deref() {
                    let trimmed = rgb.trim();
                    if trimmed.is_empty() {
                        color.rgb = None;
                    } else {
                        color.rgb = Some(normalize_tab_color_rgb(trimmed)?);
                    }
                }

                // Treat an all-empty payload as clearing the tab color.
                if color.rgb.is_none()
                    && color.theme.is_none()
                    && color.indexed.is_none()
                    && color.tint.is_none()
                    && color.auto.is_none()
                {
                    None
                } else {
                    Some(color)
                }
            }
        };
        state
            .set_sheet_tab_color(&sheet_id, tab_color)
            .map_err(app_error)?;
        Ok::<_, String>(())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn move_sheet(
    sheet_id: String,
    to_index: usize,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        state.move_sheet(&sheet_id, to_index).map_err(app_error)?;
        Ok::<_, String>(())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn delete_sheet(
    sheet_id: String,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        state.delete_sheet(&sheet_id).map_err(app_error)?;
        Ok::<_, String>(())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(any(feature = "desktop", test))]
fn read_text_file_blocking(path: &std::path::Path) -> Result<String, String> {
    use std::io::Read;

    let metadata = std::fs::metadata(path).map_err(|e| e.to_string())?;
    if !metadata.is_file() {
        return Err("Path is not a regular file".to_string());
    }
    crate::ipc_file_limits::validate_full_read_size(metadata.len()).map_err(|e| e.to_string())?;

    let file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut buf = Vec::with_capacity(metadata.len() as usize);
    file.take(crate::ipc_file_limits::MAX_READ_FULL_BYTES + 1)
        .read_to_end(&mut buf)
        .map_err(|e| e.to_string())?;
    crate::ipc_file_limits::validate_full_read_size(buf.len() as u64).map_err(|e| e.to_string())?;

    String::from_utf8(buf).map_err(|e| e.to_string())
}

/// Read a local text file on behalf of the frontend.
///
/// This exists so the desktop webview can power-query local sources (CSV/JSON) without
/// depending on the optional Tauri FS plugin.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn read_text_file(
    window: tauri::WebviewWindow,
    path: LimitedString<MAX_IPC_PATH_BYTES>,
) -> Result<String, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "filesystem access", ipc_origin::Verb::Is)?;

    let path = path.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let allowed_roots = crate::fs_scope::desktop_allowed_roots().map_err(|e| e.to_string())?;
        let resolved = crate::fs_scope::canonicalize_in_allowed_roots(
            std::path::Path::new(&path),
            &allowed_roots,
        )
        .map_err(|e| e.to_string())?;

        read_text_file_blocking(&resolved)
    })
    .await
    .map_err(|e| e.to_string())?
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct FileStat {
    pub mtime_ms: u64,
    pub size_bytes: u64,
}

/// Stat a local file and return its modification time and size.
///
/// This is used by Power Query's cache validation logic to decide whether cached results can be
/// reused when reading local sources (CSV/JSON/Parquet).
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn stat_file(
    window: tauri::WebviewWindow,
    path: LimitedString<MAX_IPC_PATH_BYTES>,
) -> Result<FileStat, String> {
    use std::time::UNIX_EPOCH;

    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "filesystem access", ipc_origin::Verb::Is)?;

    let path = path.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let allowed_roots = crate::fs_scope::desktop_allowed_roots().map_err(|e| e.to_string())?;
        let resolved = crate::fs_scope::canonicalize_in_allowed_roots(
            std::path::Path::new(&path),
            &allowed_roots,
        )
        .map_err(|e| e.to_string())?;

        let metadata = std::fs::metadata(&resolved).map_err(|e| e.to_string())?;
        let modified = metadata.modified().map_err(|e| e.to_string())?;
        let duration = modified
            .duration_since(UNIX_EPOCH)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(FileStat {
            mtime_ms: duration.as_millis() as u64,
            size_bytes: metadata.len(),
        })
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Read a local file and return its contents as base64.
///
/// The frontend decodes this into a `Uint8Array` for Parquet sources.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn read_binary_file(
    window: tauri::WebviewWindow,
    path: LimitedString<MAX_IPC_PATH_BYTES>,
) -> Result<String, String> {
    use base64::{engine::general_purpose::STANDARD, Engine as _};

    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "filesystem access", ipc_origin::Verb::Is)?;

    let path = path.into_inner();
    let bytes = tauri::async_runtime::spawn_blocking(move || {
        let allowed_roots = crate::fs_scope::desktop_allowed_roots().map_err(|e| e.to_string())?;
        let resolved = crate::fs_scope::canonicalize_in_allowed_roots(
            std::path::Path::new(&path),
            &allowed_roots,
        )
        .map_err(|e| e.to_string())?;

        read_binary_file_blocking(&resolved)
    })
    .await
    .map_err(|e| e.to_string())??;

    Ok(STANDARD.encode(bytes))
}

#[cfg(any(feature = "desktop", test))]
fn read_binary_file_blocking(path: &std::path::Path) -> Result<Vec<u8>, String> {
    use std::io::Read;

    let metadata = std::fs::metadata(path).map_err(|e| e.to_string())?;
    if !metadata.is_file() {
        return Err("Path is not a regular file".to_string());
    }
    crate::ipc_file_limits::validate_full_read_size(metadata.len()).map_err(|e| e.to_string())?;

    let file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut buf = Vec::with_capacity(metadata.len() as usize);
    file.take(crate::ipc_file_limits::MAX_READ_FULL_BYTES + 1)
        .read_to_end(&mut buf)
        .map_err(|e| e.to_string())?;
    crate::ipc_file_limits::validate_full_read_size(buf.len() as u64).map_err(|e| e.to_string())?;

    Ok(buf)
}

#[cfg(any(feature = "desktop", test))]
fn read_binary_file_range_blocking(
    path: &std::path::Path,
    offset: u64,
    len: usize,
) -> Result<Vec<u8>, String> {
    use std::io::{Read, Seek, SeekFrom};

    let metadata = std::fs::metadata(path).map_err(|e| e.to_string())?;
    if !metadata.is_file() {
        return Err("Path is not a regular file".to_string());
    }

    let mut file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    file.seek(SeekFrom::Start(offset))
        .map_err(|e| e.to_string())?;

    // Pre-allocate based on the expected read size to avoid wasting memory when callers request
    // ranges past EOF (which should return an empty buffer).
    let cap = metadata.len().saturating_sub(offset).min(len as u64) as usize;
    let mut buf = Vec::with_capacity(cap);
    file.take(len as u64)
        .read_to_end(&mut buf)
        .map_err(|e| e.to_string())?;

    Ok(buf)
}

/// Read a byte range from a local file and return the contents as base64.
///
/// This enables streaming reads for large local sources (e.g. CSV/Parquet) without first
/// materializing the full file into memory.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn read_binary_file_range(
    window: tauri::WebviewWindow,
    path: LimitedString<MAX_IPC_PATH_BYTES>,
    offset: u64,
    length: u64,
) -> Result<String, String> {
    use base64::{engine::general_purpose::STANDARD, Engine as _};

    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "filesystem access", ipc_origin::Verb::Is)?;

    let len =
        crate::ipc_file_limits::validate_read_range_length(length).map_err(|e| e.to_string())?;
    if len == 0 {
        return Ok(String::new());
    }

    let path = path.into_inner();
    let bytes = tauri::async_runtime::spawn_blocking(move || {
        let allowed_roots = crate::fs_scope::desktop_allowed_roots().map_err(|e| e.to_string())?;
        let resolved = crate::fs_scope::canonicalize_in_allowed_roots(
            std::path::Path::new(&path),
            &allowed_roots,
        )
        .map_err(|e| e.to_string())?;

        read_binary_file_range_blocking(&resolved, offset, len)
    })
    .await
    .map_err(|e| e.to_string())??;

    Ok(STANDARD.encode(bytes))
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ListDirEntry {
    pub path: String,
    pub name: String,
    pub size: u64,
    pub mtime_ms: u64,
}

/// Core implementation of `list_dir` (shared by the Tauri command and unit tests).
///
/// This intentionally enforces backend-side resource limits to prevent unbounded memory usage
/// when listing very large directories.
#[cfg(any(feature = "desktop", test))]
fn list_dir_blocking(path: &str, recursive: bool) -> Result<Vec<ListDirEntry>, String> {
    use std::path::{Path, PathBuf};

    fn visit(
        dir: &Path,
        recursive: bool,
        depth: usize,
        out: &mut Vec<ListDirEntry>,
        allowed_roots: &[PathBuf],
        seen: &mut usize,
    ) -> Result<(), String> {
        let canonical_dir = crate::fs_scope::canonicalize_in_allowed_roots(dir, allowed_roots)
            .map_err(|e| e.to_string())?;
        if depth > crate::resource_limits::MAX_LIST_DIR_DEPTH {
            return Err(format!(
                "Directory listing exceeded depth limit (max {} levels)",
                crate::resource_limits::MAX_LIST_DIR_DEPTH
            ));
        }

        let entries = std::fs::read_dir(&canonical_dir).map_err(|e| e.to_string())?;
        for entry in entries {
            let entry = entry.map_err(|e| e.to_string())?;
            if *seen >= crate::resource_limits::MAX_LIST_DIR_ENTRIES {
                return Err(format!(
                    "Directory listing exceeded limit (max {} entries)",
                    crate::resource_limits::MAX_LIST_DIR_ENTRIES
                ));
            }
            *seen += 1;

            let entry_path = entry.path();
            let file_type = entry.file_type().map_err(|e| e.to_string())?;
            let resolved_path =
                match crate::fs_scope::canonicalize_in_allowed_roots(&entry_path, allowed_roots) {
                    Ok(path) => path,
                    Err(_) => continue,
                };
            let metadata = std::fs::metadata(&resolved_path).map_err(|e| e.to_string())?;

            if metadata.is_dir() {
                // Never follow symlinked directories, to avoid cycles and unexpected traversal
                // outside the requested subtree.
                if recursive && !file_type.is_symlink() {
                    visit(&entry_path, recursive, depth + 1, out, allowed_roots, seen)?;
                }
                continue;
            }

            if !metadata.is_file() {
                continue;
            }

            let modified = metadata.modified().map_err(|e| e.to_string())?;
            let duration = modified
                .duration_since(std::time::UNIX_EPOCH)
                .map_err(|e| e.to_string())?;

            let name = entry.file_name().to_str().unwrap_or_default().to_string();

            out.push(ListDirEntry {
                path: entry_path.to_string_lossy().to_string(),
                name,
                size: metadata.len(),
                mtime_ms: duration.as_millis() as u64,
            });
        }
        Ok(())
    }

    let allowed_roots = crate::fs_scope::desktop_allowed_roots().map_err(|e| e.to_string())?;
    let root = PathBuf::from(path);
    let mut out = Vec::new();
    let mut seen = 0usize;
    visit(&root, recursive, 0, &mut out, &allowed_roots, &mut seen)?;
    Ok(out)
}

/// List files in a directory (optionally recursively) and return basic metadata.
///
/// This supports Power Query-style `Folder.Files` / `Folder.Contents` sources in the webview
/// without depending on the optional Tauri FS plugin.
///
/// Resource limits:
/// - The directory traversal is capped at `MAX_LIST_DIR_ENTRIES` (see `resource_limits.rs`).
/// - Recursive traversal is capped at `MAX_LIST_DIR_DEPTH`.
/// - Symlinked directories are not followed.
///
/// If a limit is reached, the command returns an error instead of a truncated result.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn list_dir(
    window: tauri::WebviewWindow,
    path: LimitedString<MAX_IPC_PATH_BYTES>,
    recursive: Option<bool>,
) -> Result<Vec<ListDirEntry>, String> {
    let recursive = recursive.unwrap_or(false);

    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "filesystem access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "filesystem access", ipc_origin::Verb::Is)?;

    let path = path.into_inner();
    tauri::async_runtime::spawn_blocking(move || list_dir_blocking(&path, recursive))
        .await
        .map_err(|e| e.to_string())?
}

/// Power Query: retrieve (or create) the AES-256-GCM key used to encrypt cached
/// query results at rest.
///
/// The key material is stored in the OS keychain so cached results remain
/// decryptable across app restarts.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_cache_key_get_or_create(
    window: tauri::WebviewWindow,
) -> Result<PowerQueryCacheKey, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "power query cache key access",
        ipc_origin::Verb::Is,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "power query cache key access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(
        &window,
        "power query cache key access",
        ipc_origin::Verb::Is,
    )?;

    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCacheKeyStore::open_default();
        store.get_or_create().map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Collaboration: load a persisted sync token by key.
///
/// Tokens are stored encrypted-at-rest using an OS-keychain-backed keyring.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn collab_token_get(
    window: tauri::WebviewWindow,
    token_key: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
) -> Result<Option<CollabTokenEntry>, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "collaboration tokens",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "collaboration tokens", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "collaboration tokens", ipc_origin::Verb::Are)?;

    let token_key = token_key.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = CollabTokenStore::open_default().map_err(|e| e.to_string())?;
        store.get(&token_key).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CollabTokenEntryIpc {
    pub token: LimitedString<MAX_IPC_COLLAB_TOKEN_BYTES>,
    #[serde(default)]
    pub expires_at_ms: Option<i64>,
}

#[cfg(feature = "desktop")]
impl From<CollabTokenEntryIpc> for CollabTokenEntry {
    fn from(value: CollabTokenEntryIpc) -> Self {
        CollabTokenEntry {
            token: value.token.into_inner(),
            expires_at_ms: value.expires_at_ms,
        }
    }
}

/// Collaboration: persist a sync token entry for a key.
///
/// IMPORTANT: token strings must never be logged.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn collab_token_set(
    window: tauri::WebviewWindow,
    token_key: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
    entry: CollabTokenEntryIpc,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "collaboration tokens",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "collaboration tokens", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "collaboration tokens", ipc_origin::Verb::Are)?;

    let token_key = token_key.into_inner();
    let entry: CollabTokenEntry = entry.into();
    tauri::async_runtime::spawn_blocking(move || {
        let store = CollabTokenStore::open_default().map_err(|e| e.to_string())?;
        store.set(&token_key, entry).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Collaboration: delete any persisted sync token entry for a key.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn collab_token_delete(
    window: tauri::WebviewWindow,
    token_key: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "collaboration tokens",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "collaboration tokens", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "collaboration tokens", ipc_origin::Verb::Are)?;

    let token_key = token_key.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = CollabTokenStore::open_default().map_err(|e| e.to_string())?;
        store.delete(&token_key).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: retrieve a persisted credential entry by scope key.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_credential_get(
    window: tauri::WebviewWindow,
    scope_key: LimitedString<{ crate::power_query_validation::MAX_CREDENTIAL_SCOPE_KEY_LEN }>,
) -> Result<Option<PowerQueryCredentialEntry>, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "power query credentials",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "power query credentials", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "power query credentials", ipc_origin::Verb::Are)?;

    let scope_key = scope_key.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCredentialStore::open_default().map_err(|e| e.to_string())?;
        store.get(&scope_key).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: persist a credential entry for a scope key.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_credential_set(
    window: tauri::WebviewWindow,
    scope_key: LimitedString<{ crate::power_query_validation::MAX_CREDENTIAL_SCOPE_KEY_LEN }>,
    secret: JsonValue,
) -> Result<PowerQueryCredentialEntry, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "power query credentials",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "power query credentials", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "power query credentials", ipc_origin::Verb::Are)?;

    let scope_key = scope_key.into_inner();
    crate::power_query_validation::validate_power_query_credential_payload(&scope_key, &secret)
        .map_err(|e| e.to_string())?;
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCredentialStore::open_default().map_err(|e| e.to_string())?;
        store.set(&scope_key, secret).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: delete any persisted credential entry for a scope key.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_credential_delete(
    window: tauri::WebviewWindow,
    scope_key: LimitedString<{ crate::power_query_validation::MAX_CREDENTIAL_SCOPE_KEY_LEN }>,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "power query credentials",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "power query credentials", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "power query credentials", ipc_origin::Verb::Are)?;

    let scope_key = scope_key.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCredentialStore::open_default().map_err(|e| e.to_string())?;
        store.delete(&scope_key).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: list persisted credential scope keys and IDs (for debugging).
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_credential_list(
    window: tauri::WebviewWindow,
) -> Result<Vec<PowerQueryCredentialListEntry>, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "power query credentials",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "power query credentials", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "power query credentials", ipc_origin::Verb::Are)?;

    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCredentialStore::open_default().map_err(|e| e.to_string())?;
        store.list().map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Collab E2E cell encryption: retrieve a persisted encryption key for (docId, keyId).
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn collab_encryption_key_get(
    window: tauri::WebviewWindow,
    doc_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
    key_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
) -> Result<Option<CollabEncryptionKeyEntry>, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "collab encryption keys",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "collab encryption keys", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "collab encryption keys", ipc_origin::Verb::Are)?;

    let doc_id = doc_id.into_inner();
    let key_id = key_id.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = CollabEncryptionKeyStore::open_default().map_err(|e| e.to_string())?;
        store.get(&doc_id, &key_id).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Collab E2E cell encryption: persist an encryption key for (docId, keyId).
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn collab_encryption_key_set(
    window: tauri::WebviewWindow,
    doc_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
    key_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
    key_bytes_base64: LimitedString<MAX_IPC_COLLAB_ENCRYPTION_KEY_BASE64_BYTES>,
) -> Result<CollabEncryptionKeyListEntry, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "collab encryption keys",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "collab encryption keys", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "collab encryption keys", ipc_origin::Verb::Are)?;

    let doc_id = doc_id.into_inner();
    let key_id = key_id.into_inner();
    let key_bytes_base64 = key_bytes_base64.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = CollabEncryptionKeyStore::open_default().map_err(|e| e.to_string())?;
        store
            .set(&doc_id, &key_id, &key_bytes_base64)
            .map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Collab E2E cell encryption: delete a persisted encryption key for (docId, keyId).
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn collab_encryption_key_delete(
    window: tauri::WebviewWindow,
    doc_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
    key_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "collab encryption keys",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "collab encryption keys", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "collab encryption keys", ipc_origin::Verb::Are)?;

    let doc_id = doc_id.into_inner();
    let key_id = key_id.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = CollabEncryptionKeyStore::open_default().map_err(|e| e.to_string())?;
        store.delete(&doc_id, &key_id).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Collab E2E cell encryption: list persisted encryption keys for a document.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn collab_encryption_key_list(
    window: tauri::WebviewWindow,
    doc_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
) -> Result<Vec<CollabEncryptionKeyListEntry>, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "collab encryption keys",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "collab encryption keys", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "collab encryption keys", ipc_origin::Verb::Are)?;

    let doc_id = doc_id.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = CollabEncryptionKeyStore::open_default().map_err(|e| e.to_string())?;
        store.list(&doc_id).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: load the persisted refresh scheduling state for a workbook.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_refresh_state_get(
    window: tauri::WebviewWindow,
    workbook_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
) -> Result<Option<JsonValue>, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "power query refresh state",
        ipc_origin::Verb::Is,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "power query refresh state", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "power query refresh state", ipc_origin::Verb::Is)?;

    let workbook_id = workbook_id.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryRefreshStateStore::open_default().map_err(|e| e.to_string())?;
        store.load(&workbook_id).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: persist refresh scheduling state for a workbook.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_refresh_state_set(
    window: tauri::WebviewWindow,
    workbook_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
    state: JsonValue,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "power query refresh state",
        ipc_origin::Verb::Is,
    )?;
    ipc_origin::ensure_trusted_origin(&url, "power query refresh state", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "power query refresh state", ipc_origin::Verb::Is)?;

    crate::power_query_validation::validate_power_query_refresh_state_payload(&state)
        .map_err(|e| e.to_string())?;
    let workbook_id = workbook_id.into_inner();
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryRefreshStateStore::open_default().map_err(|e| e.to_string())?;
        store.save(&workbook_id, state).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: get the workbook-backed query definition payload (if present).
///
/// This returns the raw XML part contents stored at `xl/formula/power-query.xml` (UTF-8) or `null`
/// when no workbook is loaded / the part is absent.
#[cfg(feature = "desktop")]
#[tauri::command]
pub fn power_query_state_get(
    window: tauri::WebviewWindow,
    state: State<'_, SharedAppState>,
) -> Result<Option<String>, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "power query state", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "power query state", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "power query state", ipc_origin::Verb::Is)?;

    let state = state.inner().lock().unwrap();
    let Ok(workbook) = state.get_workbook() else {
        return Ok(None);
    };
    let Some(bytes) = workbook.power_query_xml.as_deref() else {
        return Ok(None);
    };
    Ok(std::str::from_utf8(bytes).ok().map(|s| s.to_string()))
}

/// Power Query: set (or clear) the workbook-backed query definition payload in memory.
///
/// This updates the active workbook's `xl/formula/power-query.xml` part content but does not
/// persist it to disk until the workbook is saved.
#[cfg(feature = "desktop")]
#[tauri::command]
pub fn power_query_state_set(
    window: tauri::WebviewWindow,
    xml: Option<LimitedString<{ crate::power_query_validation::MAX_POWER_QUERY_XML_BYTES }>>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "power query state", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "power query state", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "power query state", ipc_origin::Verb::Is)?;

    if let Some(xml) = xml.as_ref() {
        crate::power_query_validation::validate_power_query_xml_payload(xml.as_ref())
            .map_err(|e| e.to_string())?;
    }

    let mut state = state.inner().lock().unwrap();
    let Ok(workbook) = state.get_workbook_mut() else {
        return Ok(());
    };
    workbook.power_query_xml = xml.map(|xml| xml.into_inner().into_bytes());
    Ok(())
}

/// Execute a SQL query against a local database connection.
///
/// Used by the desktop Power Query engine (`source.type === "database"`).
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn sql_query(
    window: tauri::WebviewWindow,
    connection: JsonValue,
    sql: LimitedString<{ crate::ipc_limits::MAX_SQL_QUERY_TEXT_BYTES }>,
    params: Option<LimitedVec<JsonValue, { crate::ipc_limits::MAX_SQL_QUERY_PARAMS }>>,
    credentials: Option<JsonValue>,
) -> Result<crate::sql::SqlQueryResult, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "SQL queries", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_trusted_origin(&url, "SQL queries", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "SQL queries", ipc_origin::Verb::Are)?;

    let sql = sql.into_inner();
    let params = params.map(|p| p.into_inner()).unwrap_or_default();
    crate::sql::sql_query(connection, sql, params, credentials)
        .await
        .map_err(|e| e.to_string())
}

/// Describe a SQL query (columns/types) without returning data rows.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn sql_get_schema(
    window: tauri::WebviewWindow,
    connection: JsonValue,
    sql: LimitedString<{ crate::ipc_limits::MAX_SQL_QUERY_TEXT_BYTES }>,
    credentials: Option<JsonValue>,
) -> Result<crate::sql::SqlSchemaResult, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "SQL queries", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_trusted_origin(&url, "SQL queries", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "SQL queries", ipc_origin::Verb::Are)?;

    let sql = sql.into_inner();
    crate::sql::sql_get_schema(connection, sql, credentials)
        .await
        .map_err(|e| e.to_string())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_workbook_theme_palette(
    state: State<'_, SharedAppState>,
) -> Result<Option<WorkbookThemePalette>, String> {
    let state = state.inner().lock().unwrap();
    let workbook = state.get_workbook().map_err(app_error)?;
    Ok(workbook_theme_palette(workbook))
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn list_defined_names(
    state: State<'_, SharedAppState>,
) -> Result<Vec<DefinedNameInfo>, String> {
    let state = state.inner().lock().unwrap();
    let workbook = state.get_workbook().map_err(app_error)?;

    let names = workbook
        .defined_names
        .iter()
        .filter(|n| !n.hidden)
        .filter(|n| !n.name.trim().is_empty())
        .filter(|n| !n.name.to_ascii_lowercase().starts_with("_xlnm."))
        .map(|n| DefinedNameInfo {
            name: n.name.clone(),
            refers_to: n.refers_to.clone(),
            sheet_id: n.sheet_id.clone(),
        })
        .collect();

    Ok(names)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn list_tables(state: State<'_, SharedAppState>) -> Result<Vec<TableInfo>, String> {
    let state = state.inner().lock().unwrap();
    let workbook = state.get_workbook().map_err(app_error)?;

    let tables = workbook
        .tables
        .iter()
        .filter(|t| !t.name.trim().is_empty())
        .filter(|t| !t.columns.is_empty())
        .map(|t| TableInfo {
            name: t.name.clone(),
            sheet_id: t.sheet_id.clone(),
            start_row: t.start_row,
            start_col: t.start_col,
            end_row: t.end_row,
            end_col: t.end_col,
            columns: t.columns.clone(),
        })
        .collect();

    Ok(tables)
}

#[cfg(any(feature = "desktop", test))]
fn worksheet_background_picture_rel_id(worksheet_xml: &[u8]) -> Option<String> {
    // Background images in SpreadsheetML use a simple `<picture r:id="rId1"/>` element under
    // `<worksheet>`. Avoid pulling in an additional XML dependency in the Tauri crate by using a
    // small start-tag scanner that:
    // - Locates the first `<picture ...>` start tag
    // - Extracts the first attribute whose local name is `id` (e.g. `r:id`)
    //
    // This is intentionally best-effort and tolerant of prefixes/whitespace.
    let bytes = worksheet_xml;
    let mut i = 0usize;
    while i < bytes.len() {
        if bytes[i] != b'<' {
            i += 1;
            continue;
        }
        if i + 1 >= bytes.len() {
            break;
        }

        // Skip closing tags, processing instructions, and declarations/comments.
        match bytes[i + 1] {
            b'/' | b'?' | b'!' => {
                i += 2;
                continue;
            }
            _ => {}
        }

        let mut j = i + 1;
        while j < bytes.len() && bytes[j].is_ascii_whitespace() {
            j += 1;
        }
        let name_start = j;
        while j < bytes.len()
            && !bytes[j].is_ascii_whitespace()
            && bytes[j] != b'>'
            && bytes[j] != b'/'
        {
            j += 1;
        }
        if name_start == j {
            i += 1;
            continue;
        }

        let tag_name = &bytes[name_start..j];
        if !formula_xlsx::openxml::local_name(tag_name).eq_ignore_ascii_case(b"picture") {
            i = j;
            continue;
        }

        // Parse attributes until the end of the start tag.
        let mut k = j;
        while k < bytes.len() {
            while k < bytes.len() && bytes[k].is_ascii_whitespace() {
                k += 1;
            }
            if k >= bytes.len() {
                break;
            }

            match bytes[k] {
                b'>' => break,
                b'/' => {
                    // Handle `/>` and tolerate `/ >`.
                    k += 1;
                    continue;
                }
                _ => {}
            }

            let attr_start = k;
            while k < bytes.len()
                && !bytes[k].is_ascii_whitespace()
                && bytes[k] != b'='
                && bytes[k] != b'>'
                && bytes[k] != b'/'
            {
                k += 1;
            }
            if attr_start == k {
                k += 1;
                continue;
            }
            let attr_name = &bytes[attr_start..k];

            while k < bytes.len() && bytes[k].is_ascii_whitespace() {
                k += 1;
            }
            if k >= bytes.len() || bytes[k] != b'=' {
                continue;
            }
            k += 1;
            while k < bytes.len() && bytes[k].is_ascii_whitespace() {
                k += 1;
            }
            if k >= bytes.len() {
                break;
            }

            let quote = bytes[k];
            let (value_start, value_end, next_k) = if quote == b'"' || quote == b'\'' {
                let value_start = k + 1;
                let mut end = value_start;
                while end < bytes.len() && bytes[end] != quote {
                    end += 1;
                }
                let next_k = if end < bytes.len() { end + 1 } else { end };
                (value_start, end, next_k)
            } else {
                let value_start = k;
                let mut end = value_start;
                while end < bytes.len()
                    && !bytes[end].is_ascii_whitespace()
                    && bytes[end] != b'>'
                    && bytes[end] != b'/'
                {
                    end += 1;
                }
                (value_start, end, end)
            };

            if formula_xlsx::openxml::local_name(attr_name).eq_ignore_ascii_case(b"id") {
                if let Ok(v) = std::str::from_utf8(&bytes[value_start..value_end]) {
                    let trimmed = v.trim();
                    if !trimmed.is_empty() {
                        return Some(trimmed.to_string());
                    }
                }
            }

            k = next_k;
        }

        // Stop after the first `<picture>` element. If it has no relationship id, treat it as
        // absent (best-effort).
        return None;
    }
    None
}

#[cfg(any(feature = "desktop", test))]
fn extract_imported_sheet_background_images_from_xlsx_bytes(
    xlsx_bytes: &[u8],
) -> Vec<ImportedSheetBackgroundImageInfo> {
    use base64::{engine::general_purpose::STANDARD, Engine as _};
    use std::collections::HashMap;

    use crate::resource_limits::{
        MAX_IMPORTED_SHEET_BACKGROUND_IMAGE_BYTES, MAX_IMPORTED_SHEET_BACKGROUND_IMAGES_TOTAL_BYTES,
    };

    let pkg = match formula_xlsx::XlsxPackage::from_bytes(xlsx_bytes) {
        Ok(pkg) => pkg,
        Err(_) => return Vec::new(),
    };

    let worksheets = match pkg.worksheet_parts() {
        Ok(parts) => parts,
        Err(_) => match formula_xlsx::worksheet_parts_from_reader(std::io::Cursor::new(xlsx_bytes)) {
            Ok(parts) => parts,
            Err(_) => return Vec::new(),
        },
    };

    // Cache base64 payloads for shared image parts so multiple worksheets that reference the same
    // background image don't repeatedly re-encode bytes.
    let mut cached_by_target: HashMap<String, (String, String, usize)> = HashMap::new();
    let mut total_bytes = 0usize;
    let mut out = Vec::new();
    for sheet in worksheets {
        let worksheet_part = sheet.worksheet_part.clone();
        let Some(sheet_xml) = pkg.part(&worksheet_part) else {
            continue;
        };

        let Some(rel_id) = worksheet_background_picture_rel_id(sheet_xml) else {
            continue;
        };

        let Ok(Some(target_part)) = formula_xlsx::openxml::resolve_relationship_target(
            &pkg,
            &worksheet_part,
            &rel_id,
        ) else {
            continue;
        };

        // Prefer a filename-only image id so it can be shared with other in-app image sources
        // (drawing images, embedded cell images, etc).
        let image_id = target_part
            .rsplit('/')
            .next()
            .unwrap_or(target_part.as_str())
            .to_string();
        if image_id.trim().is_empty() {
            continue;
        }

        // Reuse cached base64 if we've already loaded this target part.
        if let Some((bytes_base64, mime_type, byte_len)) =
            cached_by_target.get(&target_part).cloned()
        {
            if total_bytes.saturating_add(byte_len) > MAX_IMPORTED_SHEET_BACKGROUND_IMAGES_TOTAL_BYTES
            {
                continue;
            }
            total_bytes = total_bytes.saturating_add(byte_len);
            out.push(ImportedSheetBackgroundImageInfo {
                sheet_name: sheet.name.clone(),
                worksheet_part,
                image_id,
                bytes_base64,
                mime_type,
            });
            continue;
        }

        let Some(image_bytes) = pkg.part(&target_part) else {
            continue;
        };

        let byte_len = image_bytes.len();
        if byte_len == 0 {
            continue;
        }
        if byte_len > MAX_IMPORTED_SHEET_BACKGROUND_IMAGE_BYTES {
            continue;
        }
        if total_bytes.saturating_add(byte_len) > MAX_IMPORTED_SHEET_BACKGROUND_IMAGES_TOTAL_BYTES {
            continue;
        }
        total_bytes = total_bytes.saturating_add(byte_len);

        let bytes_base64 = STANDARD.encode(image_bytes);
        let mime_type = mime_guess::from_path(&image_id)
            .first_or_octet_stream()
            .essence_str()
            .to_string();

        cached_by_target.insert(
            target_part.clone(),
            (bytes_base64.clone(), mime_type.clone(), byte_len),
        );

        out.push(ImportedSheetBackgroundImageInfo {
            sheet_name: sheet.name,
            worksheet_part,
            image_id,
            bytes_base64,
            mime_type,
        });
    }

    out
}

/// Extract worksheet background images from the opened XLSX package (when available).
///
/// Background images are stored as `<worksheet><picture r:id="..."/></worksheet>` and reference
/// an image part (e.g. `xl/media/image1.png`) via the worksheet relationships file.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn list_imported_sheet_background_images(
    window: tauri::WebviewWindow,
    state: State<'_, SharedAppState>,
) -> Result<Vec<ImportedSheetBackgroundImageInfo>, String> {
    ipc_origin::ensure_main_window_and_stable_origin(
        &window,
        "imported sheet background image extraction",
        ipc_origin::Verb::Are,
    )?;

    let (origin_bytes, preserved_payloads) = {
        let state = state.inner().lock().unwrap();
        let Ok(workbook) = state.get_workbook() else {
            return Ok(Vec::new());
        };
        let origin_bytes = workbook.origin_xlsx_bytes.clone();
        let preserved_payloads = if origin_bytes.is_none() {
            sheet_background_image_payloads_from_preserved_parts(&workbook)
        } else {
            Vec::new()
        };
        (origin_bytes, preserved_payloads)
    };

    // Fallback: when we don't retain the full origin XLSX bytes (large workbooks or explicitly
    // dropped origin bytes), use preserved drawing parts to still surface worksheet backgrounds.
    if origin_bytes.is_none() {
        if preserved_payloads.is_empty() {
            return Ok(Vec::new());
        }

        return tauri::async_runtime::spawn_blocking(move || {
            Ok::<_, String>(imported_sheet_background_images_from_preserved_payloads(
                preserved_payloads,
            ))
        })
        .await
        .map_err(|e| e.to_string())?;
    }

    let origin_bytes = origin_bytes.expect("checked origin_bytes above");

    tauri::async_runtime::spawn_blocking(move || {
        // Best-effort: background image parsing should never prevent workbook interactions.
        Ok::<_, String>(extract_imported_sheet_background_images_from_xlsx_bytes(
            origin_bytes.as_ref(),
        ))
    })
    .await
    .map_err(|e| e.to_string())?
}
/// Extract chart drawing objects (anchors + optional parsed models) from the opened XLSX package.
///
/// The frontend uses this to populate the drawings layer (`drawingsBySheet`) so imported
/// `xdr:graphicFrame` chart objects show up in the drawing overlay. Any objects that include a
/// parsed `model` can then be rendered via `ChartRendererAdapter`.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn list_imported_chart_objects(
    window: tauri::WebviewWindow,
    state: State<'_, SharedAppState>,
) -> Result<Vec<ImportedChartObjectInfo>, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "imported chart object extraction",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(
        &url,
        "imported chart object extraction",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_stable_origin(
        &window,
        "imported chart object extraction",
        ipc_origin::Verb::Are,
    )?;

    let origin_bytes = {
        let state = state.inner().lock().unwrap();
        let Ok(workbook) = state.get_workbook() else {
            return Ok(Vec::new());
        };
        workbook.origin_xlsx_bytes.clone()
    };

    let Some(origin_bytes) = origin_bytes else {
        return Ok(Vec::new());
    };

    tauri::async_runtime::spawn_blocking(move || {
        // Best-effort: chart parsing should never prevent workbook interactions.
        let pkg = match formula_xlsx::XlsxPackage::from_bytes(origin_bytes.as_ref()) {
            Ok(pkg) => pkg,
            Err(_) => return Ok::<_, String>(Vec::new()),
        };

        let chart_objects = match pkg.extract_chart_objects() {
            Ok(objs) => objs,
            Err(_) => return Ok::<_, String>(Vec::new()),
        };

        let mut out = Vec::new();
        for obj in chart_objects {
            let chart_id = match (obj.sheet_name.as_deref(), obj.drawing_object_id) {
                (Some(sheet_name), Some(object_id)) => format!("{sheet_name}:{object_id}"),
                _ => {
                    // Fallback: when we cannot resolve the sheet/object id, fall back to a stable
                    // id based on the drawing part + relationship id.
                    format!("{}:{}", obj.drawing_part, obj.drawing_rel_id)
                }
            };

            out.push(ImportedChartObjectInfo {
                chart_id,
                rel_id: obj.drawing_rel_id,
                sheet_name: obj.sheet_name,
                drawing_object_id: obj.drawing_object_id,
                drawing_object_name: obj.drawing_object_name,
                anchor: obj.anchor,
                drawing_frame_xml: obj.drawing_frame_xml,
                model: obj.model,
            });
        }

        Ok::<_, String>(out)
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Extract embedded-in-cell images from the opened XLSX package (when available).
///
/// These correspond to Excel "Place in Cell" images (RichData `vm=` references) and are separate
/// from DrawingML images.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn list_imported_embedded_cell_images(
    window: tauri::WebviewWindow,
    state: State<'_, SharedAppState>,
) -> Result<Vec<ImportedEmbeddedCellImageInfo>, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(
        window.label(),
        "imported embedded cell image extraction",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_trusted_origin(
        &url,
        "imported embedded cell image extraction",
        ipc_origin::Verb::Are,
    )?;
    ipc_origin::ensure_stable_origin(
        &window,
        "imported embedded cell image extraction",
        ipc_origin::Verb::Are,
    )?;

    let origin_bytes = {
        let state = state.inner().lock().unwrap();
        let Ok(workbook) = state.get_workbook() else {
            return Ok(Vec::new());
        };
        workbook.origin_xlsx_bytes.clone()
    };

    let Some(origin_bytes) = origin_bytes else {
        return Ok(Vec::new());
    };

    tauri::async_runtime::spawn_blocking(move || {
        use base64::{engine::general_purpose::STANDARD, Engine as _};

        fn infer_mime_type(image_id: &str) -> String {
            let ext = image_id
                .rsplit('.')
                .next()
                .unwrap_or_default()
                .trim()
                .to_ascii_lowercase();
            match ext.as_str() {
                "png" => "image/png",
                "jpg" | "jpeg" => "image/jpeg",
                "gif" => "image/gif",
                "bmp" => "image/bmp",
                "webp" => "image/webp",
                "svg" => "image/svg+xml",
                "tif" | "tiff" => "image/tiff",
                _ => "application/octet-stream",
            }
            .to_string()
        }

        // Best-effort: embedded image parsing should never prevent workbook interactions.
        let pkg = match formula_xlsx::XlsxPackage::from_bytes(origin_bytes.as_ref()) {
            Ok(pkg) => pkg,
            Err(_) => return Ok::<_, String>(Vec::new()),
        };

        let images = match pkg.extract_embedded_cell_images() {
            Ok(images) => images,
            Err(_) => return Ok::<_, String>(Vec::new()),
        };

        let sheet_name_by_part: std::collections::HashMap<String, String> = pkg
            .worksheet_parts()
            .ok()
            .map(|parts| {
                parts
                    .into_iter()
                    .map(|info| (info.worksheet_part, info.name))
                    .collect()
            })
            .unwrap_or_default();

        // Keep individual images and the overall IPC payload bounded. We cap per-image bytes based on
        // the same limit used for other binary IPC payloads (`read_binary_file_range`), and then cap
        // the *aggregate* bytes so workbooks containing many small images cannot force an oversized
        // base64 response.
        //
        // Note: base64 expands payloads by ~33%, so the true serialized JSON payload will be larger
        // than this raw-byte sum.
        let max_image_bytes = crate::ipc_file_limits::MAX_READ_RANGE_BYTES as usize;
        let max_total_image_bytes = crate::resource_limits::MAX_MARKETPLACE_PACKAGE_BYTES;
        let mut total_image_bytes: usize = 0;

        // Iterate in a deterministic order so any truncation (due to caps) is stable across runs.
        let mut images: Vec<_> = images.into_iter().collect();
        images.sort_by(|((a_part, a_cell), a_img), ((b_part, b_cell), b_img)| {
            a_part
                .cmp(b_part)
                .then_with(|| a_cell.row.cmp(&b_cell.row))
                .then_with(|| a_cell.col.cmp(&b_cell.col))
                .then_with(|| a_img.image_part.cmp(&b_img.image_part))
        });

        let mut out = Vec::new();
        for ((worksheet_part, cell_ref), image) in images {
            let image_len = image.image_bytes.len();
            if image_len > max_image_bytes {
                // Skip oversized payloads to keep IPC memory usage bounded.
                continue;
            }
            if total_image_bytes.saturating_add(image_len) > max_total_image_bytes {
                // Best-effort: stop once we've hit the aggregate cap.
                break;
            }
            total_image_bytes = total_image_bytes.saturating_add(image_len);

            let image_part = image.image_part;
            let image_id = image_part
                .strip_prefix("xl/media/")
                .or_else(|| image_part.strip_prefix("/xl/media/"))
                .unwrap_or(&image_part)
                .to_string();

            out.push(ImportedEmbeddedCellImageInfo {
                worksheet_part: worksheet_part.clone(),
                sheet_name: sheet_name_by_part.get(&worksheet_part).cloned(),
                row: cell_ref.row as usize,
                col: cell_ref.col as usize,
                image_id: image_id.clone(),
                bytes_base64: STANDARD.encode(&image.image_bytes),
                mime_type: infer_mime_type(&image_id),
                alt_text: image.alt_text,
            });
        }

        Ok::<_, String>(out)
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Extract worksheet DrawingML objects (floating images/shapes/chart placeholders) and their
/// referenced image bytes from the opened XLSX package (when available).
///
/// Best-effort: any failure returns an empty payload so workbook open is never blocked by
/// DrawingML parsing.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn list_imported_drawing_objects(
    window: tauri::WebviewWindow,
    state: State<'_, SharedAppState>,
) -> Result<ImportedDrawingLayerPayload, String> {
    ipc_origin::ensure_main_window_and_stable_origin(
        &window,
        "imported drawing object extraction",
        ipc_origin::Verb::Are,
    )?;

    let origin_bytes = {
        let state = state.inner().lock().unwrap();
        let Ok(workbook) = state.get_workbook() else {
            return Ok(ImportedDrawingLayerPayload {
                drawings: Vec::new(),
                images: Vec::new(),
            });
        };
        workbook.origin_xlsx_bytes.clone()
    };

    let Some(origin_bytes) = origin_bytes else {
        return Ok(ImportedDrawingLayerPayload {
            drawings: Vec::new(),
            images: Vec::new(),
        });
    };

    tauri::async_runtime::spawn_blocking(move || {
        // Best-effort: DrawingML parsing should never prevent workbook interactions.
        let pkg = match formula_xlsx::XlsxPackage::from_bytes(origin_bytes.as_ref()) {
            Ok(pkg) => pkg,
            Err(_) => {
                return Ok::<_, String>(ImportedDrawingLayerPayload {
                    drawings: Vec::new(),
                    images: Vec::new(),
                })
            }
        };

        let extracted = match pkg.extract_drawing_objects() {
            Ok(objs) => objs,
            Err(_) => {
                return Ok::<_, String>(ImportedDrawingLayerPayload {
                    drawings: Vec::new(),
                    images: Vec::new(),
                })
            }
        };

        // Safety caps to keep IPC payload sizes bounded (large images can be embedded in XLSX).
        //
        // - Individual image cap: skip any single image larger than 20MiB.
        // - Total cap: stop adding images once we hit 50MiB total.
        const MAX_SINGLE_IMAGE_BYTES: usize = 20 * 1024 * 1024;
        const MAX_TOTAL_IMAGE_BYTES: usize = 50 * 1024 * 1024;

        use base64::engine::general_purpose;
        use base64::Engine as _;

        fn infer_mime_type(id: &str, bytes: &[u8]) -> Option<String> {
            let ext = id.rsplit_once('.').map(|(_, ext)| ext).unwrap_or("");
            let by_ext = formula_xlsx::drawings::content_type_for_extension(ext);
            if by_ext != "application/octet-stream" {
                return Some(by_ext.to_string());
            }
            infer::get(bytes).map(|kind| kind.mime_type().to_string())
        }

        // Collect unique image ids referenced by DrawingObjectKind::Image.
        let mut images_by_id: BTreeMap<String, ImportedWorkbookImageEntry> = BTreeMap::new();
        let mut total_bytes = 0usize;

        for entry in &extracted {
            for obj in &entry.objects {
                let formula_model::drawings::DrawingObjectKind::Image { image_id } = &obj.kind
                else {
                    continue;
                };

                let id = image_id.as_str().to_string();
                if images_by_id.contains_key(&id) {
                    continue;
                }

                // Prefer canonical `xl/media/<id>`; fall back to the raw id as a part name.
                let media_part = format!("xl/media/{id}");
                let bytes = pkg
                    .part(&media_part)
                    .or_else(|| pkg.part(&id))
                    .map(|b| b.to_vec());

                let Some(bytes) = bytes else {
                    // Missing image part: keep the drawing object but omit the image store entry.
                    continue;
                };

                if bytes.len() > MAX_SINGLE_IMAGE_BYTES {
                    continue;
                }
                if total_bytes.saturating_add(bytes.len()) > MAX_TOTAL_IMAGE_BYTES {
                    break;
                }

                let bytes_base64 = general_purpose::STANDARD.encode(&bytes);
                let mime_type = infer_mime_type(&id, &bytes);

                images_by_id.insert(
                    id.clone(),
                    ImportedWorkbookImageEntry {
                        id,
                        bytes_base64,
                        mime_type,
                    },
                );
                total_bytes = total_bytes.saturating_add(bytes.len());
            }

            if total_bytes >= MAX_TOTAL_IMAGE_BYTES {
                break;
            }
        }

        let drawings: Vec<ImportedDrawingObjectsSheetEntry> = extracted
            .into_iter()
            .map(|entry| ImportedDrawingObjectsSheetEntry {
                sheet_name: entry.sheet_name,
                sheet_part: entry.sheet_part,
                drawing_part: entry.drawing_part,
                objects: entry.objects,
            })
            .collect();

        let images: Vec<ImportedWorkbookImageEntry> = images_by_id.into_values().collect();

        Ok::<_, String>(ImportedDrawingLayerPayload { drawings, images })
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn save_workbook(
    window: tauri::WebviewWindow,
    path: Option<LimitedString<MAX_IPC_PATH_BYTES>>,
    password: Option<LimitedString<MAX_IPC_WORKBOOK_PASSWORD_BYTES>>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "workbook saving", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "workbook saving", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "workbook saving", ipc_origin::Verb::Is)?;

    let path = path.map(|p| p.into_inner());
    let password = password.map(|p| p.into_inner());
    let (save_path, workbook, storage, memory, workbook_id, autosave) = {
        let state = state.inner().lock().unwrap();
        let workbook = state.get_workbook().map_err(app_error)?.clone();
        let storage = state
            .persistent_storage()
            .ok_or_else(|| "no persistent storage available".to_string())?;
        let memory = state
            .persistent_memory_manager()
            .ok_or_else(|| "no memory manager available".to_string())?;
        let workbook_id = state
            .persistent_workbook_id()
            .ok_or_else(|| "no persistent workbook id available".to_string())?;
        let autosave = state.autosave_manager();
        let save_path = path
            .or_else(|| workbook.path.clone())
            .ok_or_else(|| "no save path provided".to_string())?;
        (save_path, workbook, storage, memory, workbook_id, autosave)
    };

    let save_path = coerce_save_path_to_xlsx(&save_path);

    let wants_origin_bytes = wants_origin_bytes_for_save_path(&save_path);
    if let Some(autosave) = autosave.as_ref() {
        autosave.flush().await.map_err(|e| e.to_string())?;
    }

    // Always flush the paging cache before exporting to ensure changes are
    // applied even if the autosave task has exited unexpectedly.
    memory.flush_dirty_pages().map_err(|e| e.to_string())?;

    let save_path_clone = save_path.clone();
    let password = password.clone();
    let (validated_save_path, written_bytes) = tauri::async_runtime::spawn_blocking(move || {
        let allowed_roots = crate::fs_scope::desktop_allowed_roots().map_err(|e| e.to_string())?;
        let resolved_path = crate::fs_scope::resolve_save_path_in_allowed_roots(
            std::path::Path::new(&save_path_clone),
            &allowed_roots,
        )
        .map_err(|e| e.to_string())?;
        let validated_save_path = resolved_path.to_string_lossy().to_string();

        let ext = resolved_path
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or_default();

        if let Some(password) = password.as_deref() {
            if password.trim().is_empty() {
                return Err("INVALID_PASSWORD: password must not be empty".to_string());
            }
        }

        // XLSB saves must go through the `formula-xlsb` round-trip writer. The storage export
        // path only knows how to generate XLSX.
        if ext.eq_ignore_ascii_case("xlsb") {
            if password.is_some() {
                return Err(
                    "ENCRYPTION_UNSUPPORTED: saving encrypted .xlsb workbooks is not supported yet; Save As .xlsx instead"
                        .to_string(),
                );
            }
            crate::file_io::write_xlsb_to_disk_blocking(&resolved_path, &workbook)
                .map_err(|e| e.to_string())?;
            return Ok::<_, String>((
                validated_save_path,
                std::sync::Arc::<[u8]>::from(Vec::new()),
            ));
        }

        if let Some(password) = password.as_deref() {
            let zip_bytes = if workbook.origin_xlsx_bytes.is_some() {
                crate::file_io::build_xlsx_bytes_blocking(&resolved_path, &workbook)
                    .map_err(|e| e.to_string())?
            } else {
                crate::persistence::build_xlsx_from_storage(
                    &storage,
                    workbook_id,
                    &workbook,
                    &resolved_path,
                )
                .map_err(|e| e.to_string())?
            };

            let ole_bytes = crate::file_io::encrypt_package_to_ole_bytes(zip_bytes.as_ref(), password)
                .map_err(|e| e.to_string())?;

            crate::atomic_write::write_file_atomic(&resolved_path, &ole_bytes)
                .map_err(|e| e.to_string())?;

            return Ok::<_, String>((validated_save_path, zip_bytes));
        }

        // Prefer the existing patch-based save path when we have the original XLSX bytes.
        // This preserves unknown parts (theme, comments, conditional formatting, etc.) by
        // rewriting only the modified worksheet XML.
        //
        // Fall back to the storage->model export path for non-XLSX origins (csv/xls) and
        // for new workbooks without an `origin_xlsx_bytes` baseline.
        let bytes = if workbook.origin_xlsx_bytes.is_some() {
            crate::file_io::write_xlsx_blocking(&resolved_path, &workbook)
                .map_err(|e| e.to_string())?
        } else {
            crate::persistence::write_xlsx_from_storage(
                &storage,
                workbook_id,
                &workbook,
                &resolved_path,
            )
            .map_err(|e| e.to_string())?
        };
        Ok::<_, String>((validated_save_path, bytes))
    })
    .await
    .map_err(|e| e.to_string())??;

    {
        let mut state = state.inner().lock().unwrap();
        state
            .mark_saved(
                Some(validated_save_path),
                wants_origin_bytes.then_some(written_bytes),
            )
            .map_err(app_error)?;
    }

    Ok(())
}

/// Mark the in-memory workbook state as saved (clears the dirty flag) without writing a file.
///
/// This is useful when the frontend returns to the last-saved state via undo/redo and wants the
/// close prompt to match `DocumentController.isDirty`.
#[cfg(feature = "desktop")]
#[tauri::command]
pub fn mark_saved(state: State<'_, SharedAppState>) -> Result<(), String> {
    let mut state = state.inner().lock().unwrap();
    state.mark_saved(None, None).map_err(app_error)?;
    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_cell(
    sheet_id: String,
    row: usize,
    col: usize,
    state: State<'_, SharedAppState>,
) -> Result<CellValue, String> {
    let state = state.inner().lock().unwrap();
    cell_value_from_state(&state, &sheet_id, row, col)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_precedents(
    sheet_id: String,
    row: usize,
    col: usize,
    transitive: Option<bool>,
    state: State<'_, SharedAppState>,
) -> Result<Vec<String>, String> {
    let state = state.inner().lock().unwrap();
    state
        .get_precedents(&sheet_id, row, col, transitive.unwrap_or(false))
        .map_err(app_error)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_dependents(
    sheet_id: String,
    row: usize,
    col: usize,
    transitive: Option<bool>,
    state: State<'_, SharedAppState>,
) -> Result<Vec<String>, String> {
    let state = state.inner().lock().unwrap();
    state
        .get_dependents(&sheet_id, row, col, transitive.unwrap_or(false))
        .map_err(app_error)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn set_cell(
    sheet_id: String,
    row: usize,
    col: usize,
    value: Option<LimitedCellValue>,
    formula: Option<LimitedCellFormula>,
    state: State<'_, SharedAppState>,
) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state
            .set_cell(
                &sheet_id,
                row,
                col,
                value.and_then(LimitedCellValue::into_json),
                formula.map(Into::into),
            )
            .map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_range(
    sheet_id: String,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    state: State<'_, SharedAppState>,
) -> Result<RangeData, String> {
    let state = state.inner().lock().unwrap();
    let cells = state
        .get_range(&sheet_id, start_row, start_col, end_row, end_col)
        .map_err(app_error)?;
    let values = cells
        .into_iter()
        .map(|row| {
            row.into_iter()
                .map(|cell| CellValue {
                    value: cell.value.as_json(),
                    formula: cell.formula,
                    display_value: cell.display_value,
                })
                .collect::<Vec<_>>()
        })
        .collect();

    Ok(RangeData {
        values,
        start_row,
        start_col,
    })
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_sheet_used_range(
    sheet_id: String,
    state: State<'_, SharedAppState>,
) -> Result<Option<SheetUsedRange>, String> {
    let state = state.inner().lock().unwrap();
    let workbook = state.get_workbook().map_err(app_error)?;
    let sheet = workbook
        .sheet(&sheet_id)
        .ok_or_else(|| app_error(AppStateError::UnknownSheet(sheet_id)))?;

    if let Some(table) = &sheet.columnar {
        let rows = table.row_count();
        let cols = table.column_count();
        if rows == 0 || cols == 0 {
            return Ok(None);
        }
        return Ok(Some(SheetUsedRange {
            start_row: 0,
            end_row: rows.saturating_sub(1),
            start_col: 0,
            end_col: cols.saturating_sub(1),
        }));
    }

    let mut min_row = usize::MAX;
    let mut min_col = usize::MAX;
    let mut max_row = 0usize;
    let mut max_col = 0usize;
    let mut has_any = false;

    for ((row, col), cell) in sheet.cells_iter() {
        // Ignore format-only cells (the UI considers used range based on value/formula).
        if cell.formula.is_none() && cell.input_value.is_none() {
            continue;
        }
        has_any = true;
        min_row = min_row.min(row);
        min_col = min_col.min(col);
        max_row = max_row.max(row);
        max_col = max_col.max(col);
    }

    if !has_any {
        return Ok(None);
    }

    Ok(Some(SheetUsedRange {
        start_row: min_row,
        end_row: max_row,
        start_col: min_col,
        end_col: max_col,
    }))
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_sheet_view_overrides(
    sheet_id: String,
    state: State<'_, SharedAppState>,
) -> Result<SheetViewOverrides, String> {
    let state = state.inner().lock().unwrap();
    let workbook = state.get_workbook().map_err(app_error)?;
    let sheet = workbook
        .sheet(&sheet_id)
        .ok_or_else(|| app_error(AppStateError::UnknownSheet(sheet_id)))?;

    let mut col_widths = BTreeMap::<u32, f32>::new();
    for (col, width) in &sheet.col_widths {
        if let Ok(col_u32) = u32::try_from(*col) {
            col_widths.insert(col_u32, *width);
        }
    }

    let mut hidden_cols: Vec<u32> = sheet
        .hidden_cols
        .iter()
        .filter_map(|c| u32::try_from(*c).ok())
        .collect();
    hidden_cols.sort_unstable();

    Ok(SheetViewOverrides {
        col_widths,
        hidden_cols,
    })
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_sheet_formatting(
    sheet_id: String,
    state: State<'_, SharedAppState>,
) -> Result<JsonValue, String> {
    let state = state.inner().lock().unwrap();
    let sheet_uuid = state.persistent_sheet_uuid(&sheet_id).map_err(app_error)?;
    let Some(storage) = state.persistent_storage() else {
        return Err(app_error(AppStateError::Persistence(
            "workbook is not backed by persistent storage".to_string(),
        )));
    };

    let meta = storage
        .get_sheet_meta(sheet_uuid)
        .map_err(|e| e.to_string())?;
    let raw = meta
        .metadata
        .as_ref()
        .and_then(|metadata| metadata.get(SHEET_FORMATTING_METADATA_KEY));
    Ok(sheet_formatting_payload_for_ipc(&sheet_id, raw))
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_sheet_view_state(
    sheet_id: String,
    state: State<'_, SharedAppState>,
) -> Result<JsonValue, String> {
    let state = state.inner().lock().unwrap();
    let sheet_uuid = state.persistent_sheet_uuid(&sheet_id).map_err(app_error)?;
    let Some(storage) = state.persistent_storage() else {
        return Err(app_error(AppStateError::Persistence(
            "workbook is not backed by persistent storage".to_string(),
        )));
    };

    let meta = storage
        .get_sheet_meta(sheet_uuid)
        .map_err(|e| e.to_string())?;
    let Some(metadata) = meta.metadata else {
        return Ok(json!({ "schemaVersion": SHEET_VIEW_SCHEMA_VERSION }));
    };

    Ok(metadata
        .get(SHEET_VIEW_METADATA_KEY)
        .cloned()
        .unwrap_or_else(|| json!({ "schemaVersion": SHEET_VIEW_SCHEMA_VERSION })))
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_sheet_imported_col_properties(
    sheet_id: String,
    state: State<'_, SharedAppState>,
) -> Result<JsonValue, String> {
    let state = state.inner().lock().unwrap();
    let sheet_uuid = state.persistent_sheet_uuid(&sheet_id).map_err(app_error)?;
    let Some(storage) = state.persistent_storage() else {
        return Err(app_error(AppStateError::Persistence(
            "workbook is not backed by persistent storage".to_string(),
        )));
    };

    let sheet = match storage
        .get_sheet_model_worksheet(sheet_uuid)
        .map_err(|e| e.to_string())?
    {
        Some(sheet) => sheet,
        None => {
            return Ok(json!({
                "schemaVersion": 1,
                "colProperties": {}
            }))
        }
    };

    let mut props_out = serde_json::Map::new();
    for (col, props) in sheet.col_properties.iter() {
        let mut entry = serde_json::Map::new();
        if let Some(width) = props.width {
            if width.is_finite() && width > 0.0 {
                entry.insert("width".to_string(), json!(width));
            }
        }
        if props.hidden {
            entry.insert("hidden".to_string(), json!(true));
        }
        if entry.is_empty() {
            continue;
        }
        props_out.insert(col.to_string(), JsonValue::Object(entry));
    }

    Ok(json!({
        "schemaVersion": 1,
        "colProperties": props_out
    }))
}

#[cfg(any(feature = "desktop", test))]
pub(crate) fn apply_sheet_formatting_deltas_inner(
    state: &mut AppState,
    payload: ApplySheetFormattingDeltasRequest,
) -> Result<(), AppStateError> {
    #[derive(Clone, Debug, PartialEq)]
    struct FormatRun {
        start_row: i64,
        end_row_exclusive: i64,
        format: JsonValue,
    }

    #[derive(Clone, Debug, Default)]
    struct FormattingState {
        default_format: JsonValue,
        row_formats: BTreeMap<i64, JsonValue>,
        col_formats: BTreeMap<i64, JsonValue>,
        format_runs_by_col: BTreeMap<i64, Vec<FormatRun>>,
        cell_formats: BTreeMap<(i64, i64), JsonValue>,
    }

    fn parse_non_negative_i64(raw: Option<&JsonValue>) -> Option<i64> {
        let v = raw?;
        let n = v
            .as_i64()
            .or_else(|| v.as_u64().and_then(|u| i64::try_from(u).ok()))?;
        (n >= 0).then_some(n)
    }

    fn parse_formatting_state(raw: Option<&JsonValue>) -> FormattingState {
        let mut out = FormattingState {
            default_format: JsonValue::Null,
            ..Default::default()
        };

        let Some(raw) = raw else {
            return out;
        };
        let Some(obj) = raw.as_object() else {
            return out;
        };

        out.default_format = obj.get("defaultFormat").cloned().unwrap_or(JsonValue::Null);

        // Row formats.
        if let Some(rows) = obj.get("rowFormats").and_then(|v| v.as_array()) {
            for entry in rows {
                let Some(row) =
                    parse_non_negative_i64(entry.get("row").or_else(|| entry.get("index")))
                else {
                    continue;
                };
                let format = entry.get("format").cloned().unwrap_or(JsonValue::Null);
                if !format.is_null() {
                    out.row_formats.insert(row, format);
                }
            }
        }

        // Col formats.
        if let Some(cols) = obj.get("colFormats").and_then(|v| v.as_array()) {
            for entry in cols {
                let Some(col) =
                    parse_non_negative_i64(entry.get("col").or_else(|| entry.get("index")))
                else {
                    continue;
                };
                let format = entry.get("format").cloned().unwrap_or(JsonValue::Null);
                if !format.is_null() {
                    out.col_formats.insert(col, format);
                }
            }
        }

        // Range-run formats.
        if let Some(cols) = obj.get("formatRunsByCol").and_then(|v| v.as_array()) {
            for entry in cols {
                let Some(col) =
                    parse_non_negative_i64(entry.get("col").or_else(|| entry.get("index")))
                else {
                    continue;
                };
                let Some(runs) = entry.get("runs").and_then(|v| v.as_array()) else {
                    continue;
                };
                let mut parsed: Vec<FormatRun> = Vec::new();
                for run in runs {
                    let Some(start_row) = parse_non_negative_i64(run.get("startRow")) else {
                        continue;
                    };
                    let Some(end_row_exclusive) = run
                        .get("endRowExclusive")
                        .and_then(|v| v.as_i64())
                        .or_else(|| {
                            run.get("endRow")
                                .and_then(|v| v.as_i64())
                                .map(|end| end.saturating_add(1))
                        })
                    else {
                        continue;
                    };
                    if end_row_exclusive <= start_row {
                        continue;
                    }
                    let format = run.get("format").cloned().unwrap_or(JsonValue::Null);
                    if format.is_null() {
                        continue;
                    }
                    parsed.push(FormatRun {
                        start_row,
                        end_row_exclusive,
                        format,
                    });
                }
                parsed.sort_by(|a, b| {
                    if a.start_row == b.start_row {
                        a.end_row_exclusive.cmp(&b.end_row_exclusive)
                    } else {
                        a.start_row.cmp(&b.start_row)
                    }
                });
                if !parsed.is_empty() {
                    out.format_runs_by_col.insert(col, parsed);
                }
            }
        }

        // Cell formats.
        if let Some(cells) = obj.get("cellFormats").and_then(|v| v.as_array()) {
            for entry in cells {
                let Some(row) = parse_non_negative_i64(entry.get("row")) else {
                    continue;
                };
                let Some(col) = parse_non_negative_i64(entry.get("col")) else {
                    continue;
                };
                let format = entry.get("format").cloned().unwrap_or(JsonValue::Null);
                if !format.is_null() {
                    out.cell_formats.insert((row, col), format);
                }
            }
        }

        out
    }

    fn serialize_formatting_state(state: FormattingState) -> JsonValue {
        let row_formats = state
            .row_formats
            .into_iter()
            .map(|(row, format)| json!({ "row": row, "format": format }))
            .collect::<Vec<_>>();
        let col_formats = state
            .col_formats
            .into_iter()
            .map(|(col, format)| json!({ "col": col, "format": format }))
            .collect::<Vec<_>>();

        let mut out = serde_json::Map::new();
        out.insert(
            "schemaVersion".to_string(),
            json!(SHEET_FORMATTING_SCHEMA_VERSION),
        );
        out.insert("defaultFormat".to_string(), state.default_format);
        out.insert("rowFormats".to_string(), JsonValue::Array(row_formats));
        out.insert("colFormats".to_string(), JsonValue::Array(col_formats));

        if !state.format_runs_by_col.is_empty() {
            let cols = state
                .format_runs_by_col
                .into_iter()
                .map(|(col, runs)| {
                    let runs = runs
                        .into_iter()
                        .map(|run| {
                            json!({
                                "startRow": run.start_row,
                                "endRowExclusive": run.end_row_exclusive,
                                "format": run.format,
                            })
                        })
                        .collect::<Vec<_>>();
                    json!({ "col": col, "runs": runs })
                })
                .collect::<Vec<_>>();
            out.insert("formatRunsByCol".to_string(), JsonValue::Array(cols));
        }

        if !state.cell_formats.is_empty() {
            let cells = state
                .cell_formats
                .into_iter()
                .map(|((row, col), format)| json!({ "row": row, "col": col, "format": format }))
                .collect::<Vec<_>>();
            out.insert("cellFormats".to_string(), JsonValue::Array(cells));
        }

        JsonValue::Object(out)
    }

    let sheet_uuid = state.persistent_sheet_uuid(&payload.sheet_id)?;
    let Some(storage) = state.persistent_storage() else {
        return Err(AppStateError::Persistence(
            "workbook is not backed by persistent storage".to_string(),
        ));
    };

    let sheet_meta = storage
        .get_sheet_meta(sheet_uuid)
        .map_err(|e| AppStateError::Persistence(e.to_string()))?;
    let mut metadata_root = match sheet_meta.metadata {
        Some(JsonValue::Object(map)) => map,
        _ => serde_json::Map::new(),
    };
    let current_formatting = metadata_root.get(SHEET_FORMATTING_METADATA_KEY).cloned();
    let mut formatting_state = parse_formatting_state(current_formatting.as_ref());

    // Apply deltas.
    if let Some(default_format) = payload.default_format.as_ref() {
        formatting_state.default_format = default_format.clone().unwrap_or(JsonValue::Null);
    }
    if let Some(LimitedSheetRowFormatDeltas(deltas)) = payload.row_formats.as_ref() {
        for delta in deltas {
            if delta.row < 0 {
                continue;
            }
            if delta.format.is_null() {
                formatting_state.row_formats.remove(&delta.row);
            } else {
                formatting_state
                    .row_formats
                    .insert(delta.row, delta.format.clone());
            }
        }
    }
    if let Some(LimitedSheetColFormatDeltas(deltas)) = payload.col_formats.as_ref() {
        for delta in deltas {
            if delta.col < 0 {
                continue;
            }
            if delta.format.is_null() {
                formatting_state.col_formats.remove(&delta.col);
            } else {
                formatting_state
                    .col_formats
                    .insert(delta.col, delta.format.clone());
            }
        }
    }
    if let Some(LimitedSheetFormatRunsByColDeltas(deltas)) = payload.format_runs_by_col.as_ref() {
        for delta in deltas {
            if delta.col < 0 {
                continue;
            }
            let mut runs = delta
                .runs
                .0
                .iter()
                .filter(|r| {
                    r.start_row >= 0 && r.end_row_exclusive > r.start_row && !r.format.is_null()
                })
                .map(|r| FormatRun {
                    start_row: r.start_row,
                    end_row_exclusive: r.end_row_exclusive,
                    format: r.format.clone(),
                })
                .collect::<Vec<_>>();
            runs.sort_by(|a, b| {
                if a.start_row == b.start_row {
                    a.end_row_exclusive.cmp(&b.end_row_exclusive)
                } else {
                    a.start_row.cmp(&b.start_row)
                }
            });
            if runs.is_empty() {
                formatting_state.format_runs_by_col.remove(&delta.col);
            } else {
                formatting_state.format_runs_by_col.insert(delta.col, runs);
            }
        }
    }
    if let Some(LimitedSheetCellFormatDeltas(deltas)) = payload.cell_formats.as_ref() {
        for delta in deltas {
            if delta.row < 0 || delta.col < 0 {
                continue;
            }
            if delta.format.is_null() {
                formatting_state
                    .cell_formats
                    .remove(&(delta.row, delta.col));
            } else {
                formatting_state
                    .cell_formats
                    .insert((delta.row, delta.col), delta.format.clone());
            }
        }
    }

    let next_formatting = serialize_formatting_state(formatting_state);
    validate_sheet_formatting_metadata_size(&next_formatting)
        .map_err(AppStateError::Persistence)?;

    metadata_root.insert(SHEET_FORMATTING_METADATA_KEY.to_string(), next_formatting);
    storage
        .set_sheet_metadata(sheet_uuid, Some(JsonValue::Object(metadata_root)))
        .map_err(|e| AppStateError::Persistence(e.to_string()))?;

    // Apply the same deltas to the in-memory engine so style-aware functions (e.g. `CELL()`)
    // reflect formatting edits immediately, without requiring a full engine rebuild.
    state.apply_sheet_formatting_deltas_to_engine(&payload)?;

    // Formatting changes affect persistence/export behavior; force full regeneration on save.
    state.mark_dirty();
    let workbook = state.get_workbook_mut()?;
    workbook.origin_xlsx_bytes = None;

    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn apply_sheet_view_deltas(
    payload: ApplySheetViewDeltasRequest,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    #[derive(Clone, Debug, Default)]
    struct ViewState {
        col_widths: BTreeMap<i64, f64>,
        row_heights: BTreeMap<i64, f64>,
    }

    fn parse_non_negative_i64(raw: Option<&JsonValue>) -> Option<i64> {
        let v = raw?;
        let n = v
            .as_i64()
            .or_else(|| v.as_u64().and_then(|u| i64::try_from(u).ok()))?;
        (n >= 0).then_some(n)
    }

    fn parse_positive_f64(raw: Option<&JsonValue>) -> Option<f64> {
        let v = raw?;
        let n = v.as_f64().or_else(|| v.as_i64().map(|i| i as f64))?;
        (n.is_finite() && n > 0.0).then_some(n)
    }

    fn parse_axis_map(
        raw: Option<&JsonValue>,
        index_keys: &[&str],
        size_keys: &[&str],
    ) -> BTreeMap<i64, f64> {
        let mut out = BTreeMap::new();
        let Some(raw) = raw else {
            return out;
        };

        // Preferred shape: `{ "0": 42, "1": 100 }`.
        if let Some(obj) = raw.as_object() {
            for (key, value) in obj {
                let idx = key.parse::<i64>().ok().filter(|n| *n >= 0);
                let Some(idx) = idx else {
                    continue;
                };
                let Some(size) = parse_positive_f64(Some(value)) else {
                    continue;
                };
                out.insert(idx, size);
            }
            return out;
        }

        // Legacy / alternate shapes: arrays of tuples or objects.
        if let Some(arr) = raw.as_array() {
            for entry in arr {
                if let Some(tuple) = entry.as_array() {
                    if tuple.len() < 2 {
                        continue;
                    }
                    let idx = parse_non_negative_i64(tuple.get(0));
                    let size = parse_positive_f64(tuple.get(1));
                    if let (Some(idx), Some(size)) = (idx, size) {
                        out.insert(idx, size);
                    }
                    continue;
                }

                let Some(obj) = entry.as_object() else {
                    continue;
                };

                let mut idx = None;
                for key in index_keys {
                    idx = idx.or_else(|| parse_non_negative_i64(obj.get(*key)));
                }
                let Some(idx) = idx else {
                    continue;
                };

                let mut size = None;
                for key in size_keys {
                    size = size.or_else(|| parse_positive_f64(obj.get(*key)));
                }
                let Some(size) = size else {
                    continue;
                };
                out.insert(idx, size);
            }
        }

        out
    }

    fn parse_view_state(raw: Option<&JsonValue>) -> ViewState {
        let mut out = ViewState::default();
        let Some(raw) = raw else {
            return out;
        };
        let Some(obj) = raw.as_object() else {
            return out;
        };

        out.col_widths =
            parse_axis_map(obj.get("colWidths"), &["col", "index"], &["width", "size"]);
        out.row_heights = parse_axis_map(
            obj.get("rowHeights"),
            &["row", "index"],
            &["height", "size"],
        );
        out
    }

    fn serialize_axis_map(map: BTreeMap<i64, f64>) -> JsonValue {
        let mut out = serde_json::Map::new();
        for (idx, size) in map {
            if idx < 0 || !size.is_finite() || size <= 0.0 {
                continue;
            }
            out.insert(idx.to_string(), json!(size));
        }
        JsonValue::Object(out)
    }

    fn serialize_view_state(state: ViewState) -> JsonValue {
        let mut out = serde_json::Map::new();
        out.insert(
            "schemaVersion".to_string(),
            json!(SHEET_VIEW_SCHEMA_VERSION),
        );
        if !state.col_widths.is_empty() {
            out.insert(
                "colWidths".to_string(),
                serialize_axis_map(state.col_widths),
            );
        }
        if !state.row_heights.is_empty() {
            out.insert(
                "rowHeights".to_string(),
                serialize_axis_map(state.row_heights),
            );
        }
        JsonValue::Object(out)
    }

    let mut state = state.inner().lock().unwrap();
    let sheet_uuid = state
        .persistent_sheet_uuid(&payload.sheet_id)
        .map_err(app_error)?;
    let Some(storage) = state.persistent_storage() else {
        return Err(app_error(AppStateError::Persistence(
            "workbook is not backed by persistent storage".to_string(),
        )));
    };

    let meta = storage
        .get_sheet_meta(sheet_uuid)
        .map_err(|e| e.to_string())?;
    let mut metadata_root = match meta.metadata {
        Some(JsonValue::Object(map)) => map,
        _ => serde_json::Map::new(),
    };

    let mut view_state = parse_view_state(metadata_root.get(SHEET_VIEW_METADATA_KEY));

    if let Some(LimitedSheetColWidthDeltas(deltas)) = payload.col_widths {
        for delta in deltas {
            if delta.col < 0 {
                continue;
            }
            match delta.width {
                Some(width) if width.is_finite() && width > 0.0 => {
                    view_state.col_widths.insert(delta.col, width);
                }
                _ => {
                    view_state.col_widths.remove(&delta.col);
                }
            }
        }
    }

    if let Some(LimitedSheetRowHeightDeltas(deltas)) = payload.row_heights {
        for delta in deltas {
            if delta.row < 0 {
                continue;
            }
            match delta.height {
                Some(height) if height.is_finite() && height > 0.0 => {
                    view_state.row_heights.insert(delta.row, height);
                }
                _ => {
                    view_state.row_heights.remove(&delta.row);
                }
            }
        }
    }

    metadata_root.insert(
        SHEET_VIEW_METADATA_KEY.to_string(),
        serialize_view_state(view_state),
    );
    storage
        .set_sheet_metadata(sheet_uuid, Some(JsonValue::Object(metadata_root)))
        .map_err(|e| e.to_string())?;

    // View state changes should contribute to the "dirty" close/save flow.
    state.mark_dirty();

    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn apply_sheet_formatting_deltas(
    payload: ApplySheetFormattingDeltasRequest,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let mut state = state.inner().lock().unwrap();
    apply_sheet_formatting_deltas_inner(&mut state, payload).map_err(app_error)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn set_range(
    sheet_id: String,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    values: LimitedRangeCellEdits,
    state: State<'_, SharedAppState>,
) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let normalized = values
            .0
            .into_iter()
            .map(|row| {
                row.into_iter()
                    .map(|c| {
                        (
                            c.value.and_then(LimitedCellValue::into_json),
                            c.formula.map(Into::into),
                        )
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>();
        let updates = state
            .set_range(
                &sheet_id, start_row, start_col, end_row, end_col, normalized,
            )
            .map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn create_pivot_table(
    request: CreatePivotTableRequest,
    state: State<'_, SharedAppState>,
) -> Result<CreatePivotTableResponse, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let (pivot_id, updates) = state
            .create_pivot_table(
                request.name.into_inner(),
                request.source_sheet_id,
                crate::state::CellRect {
                    start_row: request.source_range.start_row,
                    start_col: request.source_range.start_col,
                    end_row: request.source_range.end_row,
                    end_col: request.source_range.end_col,
                },
                crate::state::PivotDestination {
                    sheet_id: request.destination.sheet_id,
                    row: request.destination.row,
                    col: request.destination.col,
                },
                request.config.into(),
            )
            .map_err(app_error)?;

        Ok::<_, String>(CreatePivotTableResponse {
            pivot_id,
            updates: updates.into_iter().map(cell_update_from_state).collect(),
        })
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn refresh_pivot_table(
    request: RefreshPivotTableRequest,
    state: State<'_, SharedAppState>,
) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state
            .refresh_pivot_table(&request.pivot_id)
            .map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn refresh_all_pivots(
    state: State<'_, SharedAppState>,
) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state.refresh_all_pivots().map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn list_pivot_tables(
    state: State<'_, SharedAppState>,
) -> Result<Vec<PivotTableSummary>, String> {
    let state = state.inner().lock().unwrap();
    Ok(state
        .list_pivot_tables()
        .into_iter()
        .map(|pivot| PivotTableSummary {
            id: pivot.id,
            name: pivot.name,
            source_sheet_id: pivot.source_sheet_id,
            source_range: PivotCellRange {
                start_row: pivot.source_range.start_row,
                start_col: pivot.source_range.start_col,
                end_row: pivot.source_range.end_row,
                end_col: pivot.source_range.end_col,
            },
            destination: PivotDestination {
                sheet_id: pivot.destination.sheet_id,
                row: pivot.destination.row,
                col: pivot.destination.col,
            },
        })
        .collect())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn recalculate(state: State<'_, SharedAppState>) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state.recalculate_all().map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn undo(state: State<'_, SharedAppState>) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state.undo().map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn redo(state: State<'_, SharedAppState>) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state.redo().map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
fn to_core_page_setup(setup: &PageSetup) -> formula_xlsx::print::PageSetup {
    use formula_xlsx::print as core;

    let orientation = match setup.orientation {
        PageOrientation::Portrait => core::Orientation::Portrait,
        PageOrientation::Landscape => core::Orientation::Landscape,
    };

    let margins = core::PageMargins {
        left: setup.margins.left,
        right: setup.margins.right,
        top: setup.margins.top,
        bottom: setup.margins.bottom,
        header: setup.margins.header,
        footer: setup.margins.footer,
    };

    let scaling = match setup.scaling {
        PageScaling::Percent { percent } => core::Scaling::Percent(percent),
        PageScaling::FitTo {
            width_pages,
            height_pages,
        } => core::Scaling::FitTo {
            width: width_pages,
            height: height_pages,
        },
    };

    core::PageSetup {
        orientation,
        paper_size: core::PaperSize {
            code: setup.paper_size,
        },
        margins,
        scaling,
    }
}

#[cfg(feature = "desktop")]
fn from_core_page_setup(setup: &formula_xlsx::print::PageSetup) -> PageSetup {
    let orientation = match setup.orientation {
        formula_xlsx::print::Orientation::Portrait => PageOrientation::Portrait,
        formula_xlsx::print::Orientation::Landscape => PageOrientation::Landscape,
    };

    let scaling = match setup.scaling {
        formula_xlsx::print::Scaling::Percent(percent) => PageScaling::Percent { percent },
        formula_xlsx::print::Scaling::FitTo { width, height } => PageScaling::FitTo {
            width_pages: width,
            height_pages: height,
        },
    };

    PageSetup {
        orientation,
        paper_size: setup.paper_size.code,
        margins: PageMargins {
            left: setup.margins.left,
            right: setup.margins.right,
            top: setup.margins.top,
            bottom: setup.margins.bottom,
            header: setup.margins.header,
            footer: setup.margins.footer,
        },
        scaling,
    }
}

#[cfg(feature = "desktop")]
fn from_core_sheet_print_settings(
    settings: &formula_xlsx::print::SheetPrintSettings,
) -> SheetPrintSettings {
    SheetPrintSettings {
        sheet_name: settings.sheet_name.clone(),
        print_area: settings.print_area.as_ref().map(|ranges| {
            ranges
                .iter()
                .map(|r| PrintCellRange {
                    start_row: r.start_row,
                    end_row: r.end_row,
                    start_col: r.start_col,
                    end_col: r.end_col,
                })
                .collect()
        }),
        print_titles: settings.print_titles.as_ref().map(|t| PrintTitles {
            repeat_rows: t.repeat_rows.map(|r| PrintRowRange {
                start: r.start,
                end: r.end,
            }),
            repeat_cols: t.repeat_cols.map(|r| PrintColRange {
                start: r.start,
                end: r.end,
            }),
        }),
        page_setup: from_core_page_setup(&settings.page_setup),
        manual_page_breaks: ManualPageBreaks {
            row_breaks_after: settings
                .manual_page_breaks
                .row_breaks_after
                .iter()
                .copied()
                .collect(),
            col_breaks_after: settings
                .manual_page_breaks
                .col_breaks_after
                .iter()
                .copied()
                .collect(),
        },
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_sheet_print_settings(
    sheet_id: String,
    state: State<'_, SharedAppState>,
) -> Result<SheetPrintSettings, String> {
    let state = state.inner().lock().unwrap();
    let settings = state.sheet_print_settings(&sheet_id).map_err(app_error)?;
    Ok(from_core_sheet_print_settings(&settings))
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn set_sheet_page_setup(
    sheet_id: String,
    page_setup: PageSetup,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let mut state = state.inner().lock().unwrap();
    state
        .set_sheet_page_setup(&sheet_id, to_core_page_setup(&page_setup))
        .map_err(app_error)?;
    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn set_sheet_print_area(
    sheet_id: String,
    print_area: Option<LimitedVec<PrintCellRange, { crate::ipc_limits::MAX_PRINT_AREA_RANGES }>>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let print_area = print_area.map(|ranges| {
        ranges
            .into_inner()
            .into_iter()
            .map(|r| formula_xlsx::print::CellRange {
                start_row: r.start_row,
                end_row: r.end_row,
                start_col: r.start_col,
                end_col: r.end_col,
            })
            .collect()
    });

    let mut state = state.inner().lock().unwrap();
    state
        .set_sheet_print_area(&sheet_id, print_area)
        .map_err(app_error)?;
    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn export_sheet_range_pdf(
    sheet_id: String,
    range: PrintCellRange,
    col_widths_points: Option<LimitedF64Vec>,
    row_heights_points: Option<LimitedF64Vec>,
    state: State<'_, SharedAppState>,
) -> Result<String, String> {
    use crate::resource_limits::{MAX_PDF_BYTES, MAX_PDF_CELLS_PER_CALL, MAX_RANGE_DIM};
    use base64::{engine::general_purpose::STANDARD, Engine as _};

    let state_guard = state.inner().lock().unwrap();
    let workbook = state_guard.get_workbook().map_err(app_error)?;
    let sheet = workbook
        .sheet(&sheet_id)
        .ok_or_else(|| app_error(AppStateError::UnknownSheet(sheet_id.clone())))?;

    let settings = state_guard
        .sheet_print_settings(&sheet_id)
        .map_err(app_error)?;

    let print_area = formula_xlsx::print::CellRange {
        start_row: range.start_row,
        end_row: range.end_row,
        start_col: range.start_col,
        end_col: range.end_col,
    }
    .normalized();

    if print_area.start_row == 0
        || print_area.start_col == 0
        || print_area.end_row == 0
        || print_area.end_col == 0
    {
        return Err("invalid print range: rows/cols must be 1-based".to_string());
    }

    let row_count = (print_area.end_row as u64).saturating_sub(print_area.start_row as u64) + 1;
    let col_count = (print_area.end_col as u64).saturating_sub(print_area.start_col as u64) + 1;

    // Fail fast before PDF generation to avoid CPU/memory DoS from untrusted webview input.
    let row_count_usize = row_count as usize;
    let col_count_usize = col_count as usize;
    if row_count_usize > MAX_RANGE_DIM || col_count_usize > MAX_RANGE_DIM {
        return Err(app_error(AppStateError::RangeDimensionTooLarge {
            rows: row_count_usize,
            cols: col_count_usize,
            limit: MAX_RANGE_DIM,
        }));
    }
    let cell_count = (row_count as u128) * (col_count as u128);
    if cell_count > MAX_PDF_CELLS_PER_CALL as u128 {
        return Err(app_error(AppStateError::RangeTooLarge {
            rows: row_count_usize,
            cols: col_count_usize,
            limit: MAX_PDF_CELLS_PER_CALL,
        }));
    }

    let col_widths = col_widths_points.map(|v| v.0).unwrap_or_default();
    let row_heights = row_heights_points.map(|v| v.0).unwrap_or_default();

    let pdf_bytes = formula_xlsx::print::export_range_to_pdf_bytes(
        &sheet.name,
        print_area,
        &col_widths,
        &row_heights,
        &settings.page_setup,
        &settings.manual_page_breaks,
        |row, col| {
            let value = workbook.cell_value(&sheet_id, (row - 1) as usize, (col - 1) as usize);
            let text = value.display();
            if text.is_empty() {
                None
            } else {
                Some(text)
            }
        },
    )
    .map_err(|e| e.to_string())?;

    if pdf_bytes.len() > MAX_PDF_BYTES {
        return Err(format!(
            "generated PDF is too large: {} bytes (limit {})",
            pdf_bytes.len(),
            MAX_PDF_BYTES
        ));
    }

    Ok(STANDARD.encode(pdf_bytes))
}

pub use crate::macros::{MacroInfo, MacroPermission, MacroPermissionRequest};

/// IPC-deserialized vector of macro permissions with a conservative maximum length.
///
/// This prevents a compromised webview from sending huge permission arrays and forcing unbounded
/// allocations during JSON deserialization.
#[derive(Clone, Debug, PartialEq)]
pub struct LimitedMacroPermissions(pub Vec<MacroPermission>);

impl<'de> Deserialize<'de> for LimitedMacroPermissions {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedMacroPermissions;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of macro permissions")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_MACRO_PERMISSION_ENTRIES;

                let mut out = match seq.size_hint() {
                    Some(hint) => Vec::with_capacity(hint.min(MAX_MACRO_PERMISSION_ENTRIES)),
                    None => Vec::new(),
                };

                for _ in 0..MAX_MACRO_PERMISSION_ENTRIES {
                    match seq.next_element::<MacroPermission>()? {
                        Some(permission) => out.push(permission),
                        None => return Ok(LimitedMacroPermissions(out)),
                    }
                }

                // Detect any additional entries without allocating them.
                if seq.next_element::<de::IgnoredAny>()?.is_some() {
                    return Err(de::Error::custom(format!(
                        "macro permissions list is too large (max {MAX_MACRO_PERMISSION_ENTRIES} entries)"
                    )));
                }

                Ok(LimitedMacroPermissions(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

#[derive(Clone, Copy, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum MacroSignatureStatus {
    Unsigned,
    SignedVerified,
    SignedInvalid,
    SignedParseError,
    SignedUnverified,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroSignatureInfo {
    pub status: MacroSignatureStatus,
    pub signer_subject: Option<String>,
    /// Raw signature blob, base64 encoded. May be omitted in the future if it grows large.
    pub signature_base64: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroSecurityStatus {
    pub has_macros: bool,
    pub origin_path: Option<String>,
    pub workbook_fingerprint: Option<String>,
    pub signature: Option<MacroSignatureInfo>,
    pub trust: MacroTrustDecision,
}

#[derive(Clone, Copy, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum MacroBlockedReason {
    NotTrusted,
    SignatureRequired,
}

/// Evaluate whether Trust Center policy allows macro execution.
///
/// Note: This is intentionally a pure function so it can be unit-tested without
/// requiring the full Tauri "desktop" feature.
pub fn evaluate_macro_trust(
    trust: MacroTrustDecision,
    signature_status: MacroSignatureStatus,
) -> Result<(), MacroBlockedReason> {
    match trust {
        MacroTrustDecision::TrustedAlways | MacroTrustDecision::TrustedOnce => Ok(()),
        MacroTrustDecision::Blocked => Err(MacroBlockedReason::NotTrusted),
        MacroTrustDecision::TrustedSignedOnly => match signature_status {
            MacroSignatureStatus::SignedVerified => Ok(()),
            _ => Err(MacroBlockedReason::SignatureRequired),
        },
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroBlockedError {
    pub reason: MacroBlockedReason,
    pub status: MacroSecurityStatus,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroError {
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub code: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub blocked: Option<MacroBlockedError>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroRunResult {
    pub ok: bool,
    pub output: Vec<String>,
    pub updates: Vec<CellUpdate>,
    pub error: Option<MacroError>,
    pub permission_request: Option<MacroPermissionRequest>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum PythonFilesystemPermission {
    None,
    Read,
    Readwrite,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum PythonNetworkPermission {
    None,
    Allowlist,
    Full,
}

/// IPC-deserialized string with a maximum byte length for Python network allowlist entries.
#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
pub struct LimitedPythonNetworkAllowlistEntry(pub String);

impl<'de> Deserialize<'de> for LimitedPythonNetworkAllowlistEntry {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct StringVisitor;

        impl<'de> de::Visitor<'de> for StringVisitor {
            type Value = LimitedPythonNetworkAllowlistEntry;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("a string")
            }

            fn visit_borrowed_str<E>(self, v: &'de str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                use crate::resource_limits::MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES;
                if v.len() > MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES {
                    return Err(E::custom(format!(
                        "python network allowlist entry is too large (max {MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES} bytes)"
                    )));
                }
                Ok(LimitedPythonNetworkAllowlistEntry(v.to_string()))
            }

            fn visit_str<E>(self, v: &str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                use crate::resource_limits::MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES;
                if v.len() > MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES {
                    return Err(E::custom(format!(
                        "python network allowlist entry is too large (max {MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES} bytes)"
                    )));
                }
                Ok(LimitedPythonNetworkAllowlistEntry(v.to_string()))
            }

            fn visit_string<E>(self, v: String) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                use crate::resource_limits::MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES;
                if v.len() > MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES {
                    return Err(E::custom(format!(
                        "python network allowlist entry is too large (max {MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES} bytes)"
                    )));
                }
                Ok(LimitedPythonNetworkAllowlistEntry(v))
            }
        }

        deserializer.deserialize_str(StringVisitor)
    }
}

/// IPC-deserialized vector of Python network allowlist entries with a maximum length.
#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
pub struct LimitedPythonNetworkAllowlist(pub Vec<LimitedPythonNetworkAllowlistEntry>);

impl<'de> Deserialize<'de> for LimitedPythonNetworkAllowlist {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct VecVisitor;

        impl<'de> de::Visitor<'de> for VecVisitor {
            type Value = LimitedPythonNetworkAllowlist;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("an array of strings")
            }

            fn visit_seq<A>(self, mut seq: A) -> Result<Self::Value, A::Error>
            where
                A: de::SeqAccess<'de>,
            {
                use crate::resource_limits::MAX_PYTHON_NETWORK_ALLOWLIST_ENTRIES;

                let mut out = match seq.size_hint() {
                    Some(hint) => {
                        Vec::with_capacity(hint.min(MAX_PYTHON_NETWORK_ALLOWLIST_ENTRIES))
                    }
                    None => Vec::new(),
                };

                for _ in 0..MAX_PYTHON_NETWORK_ALLOWLIST_ENTRIES {
                    match seq.next_element::<LimitedPythonNetworkAllowlistEntry>()? {
                        Some(entry) => out.push(entry),
                        None => return Ok(LimitedPythonNetworkAllowlist(out)),
                    }
                }

                // Detect any additional entries without allocating them.
                if seq.next_element::<de::IgnoredAny>()?.is_some() {
                    return Err(de::Error::custom(format!(
                        "python network allowlist is too large (max {MAX_PYTHON_NETWORK_ALLOWLIST_ENTRIES} entries)"
                    )));
                }

                Ok(LimitedPythonNetworkAllowlist(out))
            }
        }

        deserializer.deserialize_seq(VecVisitor)
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PythonPermissions {
    pub filesystem: PythonFilesystemPermission,
    pub network: PythonNetworkPermission,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub network_allowlist: Option<LimitedPythonNetworkAllowlist>,
}

impl Default for PythonPermissions {
    fn default() -> Self {
        Self {
            filesystem: PythonFilesystemPermission::None,
            network: PythonNetworkPermission::None,
            network_allowlist: None,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PythonSelection {
    pub sheet_id: String,
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PythonRunContext {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub active_sheet_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub selection: Option<PythonSelection>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PythonError {
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stack: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PythonRunResult {
    pub ok: bool,
    pub stdout: String,
    pub stderr: String,
    pub updates: Vec<CellUpdate>,
    pub error: Option<PythonError>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct TypeScriptRunResult {
    pub ok: bool,
    pub updates: Vec<CellUpdate>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroSelectionRect {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum MigrationTarget {
    Python,
    TypeScript,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct MigrationValidationMismatch {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
    pub vba: CellValue,
    pub script: CellValue,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct MigrationValidationReport {
    pub ok: bool,
    pub macro_id: String,
    pub target: MigrationTarget,
    pub mismatches: Vec<MigrationValidationMismatch>,
    pub vba: MacroRunResult,
    pub python: Option<PythonRunResult>,
    pub typescript: Option<TypeScriptRunResult>,
    pub error: Option<String>,
}

#[cfg(any(feature = "desktop", test))]
fn workbook_identity_for_trust(workbook: &Workbook, workbook_id: Option<&str>) -> String {
    workbook
        .origin_path
        .as_deref()
        .or(workbook.path.as_deref())
        .or(workbook_id)
        .unwrap_or("untitled")
        .to_string()
}

#[cfg(any(feature = "desktop", test))]
fn compute_workbook_fingerprint(
    workbook: &mut Workbook,
    workbook_id: Option<&str>,
) -> Option<String> {
    if workbook.vba_project_bin.is_none() {
        return None;
    }
    if let Some(fp) = workbook.macro_fingerprint.as_deref() {
        return Some(fp.to_string());
    }
    let id = workbook_identity_for_trust(workbook, workbook_id);
    let vba = workbook
        .vba_project_bin
        .as_deref()
        .expect("checked is_some above");
    let fp = compute_macro_fingerprint(&id, vba);
    workbook.macro_fingerprint = Some(fp.clone());
    Some(fp)
}

#[cfg(any(feature = "desktop", test))]
fn build_macro_security_status(
    workbook: &mut Workbook,
    workbook_id: Option<&str>,
    trust_store: &crate::macro_trust::MacroTrustStore,
) -> Result<MacroSecurityStatus, String> {
    use base64::Engine as _;

    let has_macros = workbook.vba_project_bin.is_some();
    let fingerprint = compute_workbook_fingerprint(workbook, workbook_id);

    let signature = if let Some(vba_bin) = workbook.vba_project_bin.as_deref() {
        // Signature parsing is best-effort: failures should not prevent macro listing or
        // execution (trust decisions are still enforced by the fingerprint).
        // Prefer the dedicated `xl/vbaProjectSignature.bin` part when present, because some XLSM
        // producers store the `\x05DigitalSignature*` streams outside of `vbaProject.bin`.
        //
        // Note: `origin_xlsx_bytes` may be dropped during some workbook edits (forcing regeneration
        // on save). In that case, rely on the in-memory `Workbook::vba_project_signature_bin`
        // instead of re-reading the original package.
        let mut sig_part_fallback: Option<Vec<u8>> = None;
        if workbook.vba_project_signature_bin.is_none() {
            sig_part_fallback = workbook.origin_xlsx_bytes.as_deref().and_then(|origin| {
                formula_xlsx::read_part_from_reader_limited(
                    std::io::Cursor::new(origin),
                    "xl/vbaProjectSignature.bin",
                    crate::resource_limits::MAX_VBA_PROJECT_SIGNATURE_BIN_BYTES as u64,
                )
                .ok()
                .flatten()
            });
        }
        let sig_part = workbook
            .vba_project_signature_bin
            .as_deref()
            .or(sig_part_fallback.as_deref());

        // Match `formula_xlsx::XlsxPackage::verify_vba_digital_signature` behavior:
        // - Prefer the signature-part signature when it cryptographically verifies.
        // - Otherwise, fall back to an embedded signature inside `vbaProject.bin`.
        // - If neither verifies, return the best-effort signature info (parse errors included).
        let mut signature_part_result: Option<formula_vba::VbaDigitalSignature> = None;
        if let Some(sig_part) = sig_part {
            match formula_vba::verify_vba_digital_signature_with_project(vba_bin, sig_part) {
                Ok(Some(sig)) => signature_part_result = Some(sig),
                Ok(None) => {}
                Err(_) => {
                    // Not an OLE container: fall back to verifying the part bytes as a raw PKCS#7/CMS
                    // signature blob.
                    let (verification, signer_subject) =
                        formula_vba::verify_vba_signature_blob(sig_part);
                    signature_part_result = Some(formula_vba::VbaDigitalSignature {
                        stream_path: "xl/vbaProjectSignature.bin".to_string(),
                        stream_kind: formula_vba::VbaSignatureStreamKind::Unknown,
                        signer_subject,
                        signature: sig_part.to_vec(),
                        verification,
                        binding: formula_vba::VbaSignatureBinding::Unknown,
                    });
                }
            }
        }

        if let Some(sig) = signature_part_result.as_mut() {
            if sig.verification == formula_vba::VbaSignatureVerification::SignedVerified
                && sig.binding == formula_vba::VbaSignatureBinding::Unknown
            {
                sig.binding = match formula_vba::verify_vba_project_signature_binding(
                    vba_bin,
                    &sig.signature,
                ) {
                    Ok(binding) => match binding {
                        formula_vba::VbaProjectBindingVerification::BoundVerified(_) => {
                            formula_vba::VbaSignatureBinding::Bound
                        }
                        formula_vba::VbaProjectBindingVerification::BoundMismatch(_) => {
                            formula_vba::VbaSignatureBinding::NotBound
                        }
                        formula_vba::VbaProjectBindingVerification::BoundUnknown(_) => {
                            formula_vba::VbaSignatureBinding::Unknown
                        }
                    },
                    Err(_) => formula_vba::VbaSignatureBinding::Unknown,
                };
            }
        }

        let embedded = formula_vba::verify_vba_digital_signature(vba_bin)
            .ok()
            .flatten();

        let parsed = if signature_part_result.as_ref().is_some_and(|sig| {
            sig.verification == formula_vba::VbaSignatureVerification::SignedVerified
        }) {
            signature_part_result
        } else if embedded.as_ref().is_some_and(|sig| {
            sig.verification == formula_vba::VbaSignatureVerification::SignedVerified
        }) {
            embedded
        } else {
            signature_part_result.or(embedded)
        };

        Some(match parsed {
            Some(sig) => MacroSignatureInfo {
                status: match sig.verification {
                    formula_vba::VbaSignatureVerification::SignedVerified => match sig.binding {
                        formula_vba::VbaSignatureBinding::Bound => {
                            MacroSignatureStatus::SignedVerified
                        }
                        formula_vba::VbaSignatureBinding::NotBound => {
                            // Signature blob is cryptographically valid, but not bound to the VBA
                            // project contents.
                            MacroSignatureStatus::SignedInvalid
                        }
                        formula_vba::VbaSignatureBinding::Unknown => {
                            // We couldn't verify the MS-OVBA Contents Hash binding. Treat it as
                            // unverified so `TrustedSignedOnly` continues to behave conservatively.
                            MacroSignatureStatus::SignedUnverified
                        }
                    },
                    formula_vba::VbaSignatureVerification::SignedInvalid => {
                        MacroSignatureStatus::SignedInvalid
                    }
                    formula_vba::VbaSignatureVerification::SignedParseError => {
                        MacroSignatureStatus::SignedParseError
                    }
                    formula_vba::VbaSignatureVerification::SignedButUnverified => {
                        MacroSignatureStatus::SignedUnverified
                    }
                },
                signer_subject: sig.signer_subject,
                signature_base64: Some(
                    base64::engine::general_purpose::STANDARD.encode(sig.signature),
                ),
            },
            None => MacroSignatureInfo {
                status: MacroSignatureStatus::Unsigned,
                signer_subject: None,
                signature_base64: None,
            },
        })
    } else {
        None
    };

    let trust = fingerprint
        .as_deref()
        .map(|fp| trust_store.trust_state(fp))
        .unwrap_or(MacroTrustDecision::Blocked);

    Ok(MacroSecurityStatus {
        has_macros,
        origin_path: workbook.origin_path.clone(),
        workbook_fingerprint: fingerprint,
        signature,
        trust,
    })
}

#[cfg(feature = "desktop")]
fn enforce_macro_trust(
    workbook: &mut Workbook,
    workbook_id: Option<&str>,
    trust_store: &crate::macro_trust::MacroTrustStore,
) -> Result<Option<MacroBlockedError>, String> {
    let status = build_macro_security_status(workbook, workbook_id, trust_store)?;
    if !status.has_macros {
        return Ok(None);
    }

    let signature_status = status
        .signature
        .as_ref()
        .map(|s| s.status)
        .unwrap_or(MacroSignatureStatus::Unsigned);

    match evaluate_macro_trust(status.trust.clone(), signature_status) {
        Ok(()) => Ok(None),
        Err(reason) => Ok(Some(MacroBlockedError { reason, status })),
    }
}

#[cfg(feature = "desktop")]
fn macro_blocked_result(blocked: MacroBlockedError) -> MacroRunResult {
    MacroRunResult {
        ok: false,
        output: Vec::new(),
        updates: Vec::new(),
        error: Some(MacroError {
            message: "Macros are blocked by Trust Center policy.".to_string(),
            code: Some("macro_blocked".to_string()),
            blocked: Some(blocked),
        }),
        permission_request: None,
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct VbaReferenceSummary {
    pub name: Option<String>,
    pub guid: Option<String>,
    pub major: Option<u16>,
    pub minor: Option<u16>,
    pub path: Option<String>,
    pub raw: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct VbaModuleSummary {
    pub name: String,
    pub module_type: String,
    pub code: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct VbaProjectSummary {
    pub name: Option<String>,
    pub constants: Option<String>,
    pub references: Vec<VbaReferenceSummary>,
    pub modules: Vec<VbaModuleSummary>,
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn get_macro_security_status(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroSecurityStatus, String> {
    ipc_origin::ensure_main_window(window.label(), "macro trust", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro trust", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro trust", ipc_origin::Verb::Is)?;

    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let workbook_id = workbook_id.as_ref().map(|id| id.as_ref());
        let mut state = shared.lock().unwrap();
        let mut trust_store = trust_shared.lock().unwrap();
        trust_store.ensure_loaded();
        let workbook = state.get_workbook_mut().map_err(app_error)?;
        build_macro_security_status(workbook, workbook_id, &trust_store)
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn set_macro_trust(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    decision: MacroTrustDecision,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroSecurityStatus, String> {
    ipc_origin::ensure_main_window(window.label(), "macro trust", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro trust", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro trust", ipc_origin::Verb::Is)?;

    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let workbook_id = workbook_id.as_ref().map(|id| id.as_ref());
        let mut state = shared.lock().unwrap();
        let mut trust_store = trust_shared.lock().unwrap();

        let workbook = state.get_workbook_mut().map_err(app_error)?;
        let Some(fingerprint) = compute_workbook_fingerprint(workbook, workbook_id) else {
            return Err("workbook has no macros to trust".to_string());
        };

        trust_store
            .set_trust(fingerprint, decision)
            .map_err(|e| e.to_string())?;

        build_macro_security_status(workbook, workbook_id, &trust_store)
    })
    .await
    .map_err(|e| e.to_string())?
}
#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_vba_project(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    state: State<'_, SharedAppState>,
) -> Result<Option<VbaProjectSummary>, String> {
    ipc_origin::ensure_main_window(window.label(), "macro execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro execution", ipc_origin::Verb::Is)?;

    let _ = workbook_id;
    let mut state = state.inner().lock().unwrap();
    let Some(project) = state.vba_project().map_err(|e| e.to_string())? else {
        return Ok(None);
    };
    Ok(Some(VbaProjectSummary {
        name: project.name,
        constants: project.constants,
        references: project
            .references
            .into_iter()
            .map(|r| VbaReferenceSummary {
                name: r.name,
                guid: r.guid,
                major: r.major,
                minor: r.minor,
                path: r.path,
                raw: r.raw,
            })
            .collect(),
        modules: project
            .modules
            .into_iter()
            .map(|m| VbaModuleSummary {
                name: m.name,
                module_type: format!("{:?}", m.module_type),
                code: m.code,
            })
            .collect(),
    }))
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn list_macros(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    state: State<'_, SharedAppState>,
) -> Result<Vec<MacroInfo>, String> {
    ipc_origin::ensure_main_window(window.label(), "macro execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro execution", ipc_origin::Verb::Is)?;

    let _ = workbook_id;

    let mut state = state.inner().lock().unwrap();
    state.list_macros().map_err(|e| e.to_string())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn set_macro_ui_context(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    sheet_id: LimitedString<{ crate::ipc_limits::MAX_SHEET_ID_BYTES }>,
    active_row: usize,
    active_col: usize,
    selection: Option<MacroSelectionRect>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    ipc_origin::ensure_main_window(window.label(), "macro execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro execution", ipc_origin::Verb::Is)?;

    let _ = workbook_id;
    let mut state = state.inner().lock().unwrap();
    let selection = selection.map(|rect| crate::state::CellRect {
        start_row: rect.start_row,
        start_col: rect.start_col,
        end_row: rect.end_row,
        end_col: rect.end_col,
    });
    state
        .set_macro_ui_context(sheet_id.as_ref(), active_row, active_col, selection)
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn run_macro(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    macro_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
    permissions: Option<LimitedMacroPermissions>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    ipc_origin::ensure_main_window(window.label(), "macro execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro execution", ipc_origin::Verb::Is)?;

    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let mut trust_store = trust_shared.lock().unwrap();
            trust_store.ensure_loaded();
            let workbook_id = workbook_id.as_ref().map(|id| id.as_ref());
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }

        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.map(|p| p.0).unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .run_macro(macro_id.as_ref(), options)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn run_python_script(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    code: IpcScriptCode,
    permissions: Option<PythonPermissions>,
    timeout_ms: Option<u64>,
    max_memory_bytes: Option<u64>,
    context: Option<PythonRunContext>,
    state: State<'_, SharedAppState>,
) -> Result<PythonRunResult, String> {
    ipc_origin::ensure_main_window(window.label(), "python execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "python execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "python execution", ipc_origin::Verb::Is)?;

    let _ = workbook_id;

    let code = code.into_inner();
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        crate::python::run_python_script(
            &mut state,
            &code,
            permissions,
            timeout_ms,
            max_memory_bytes,
            context,
        )
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_string_literal_prefix(input: &str) -> Result<(String, &str), String> {
    let trimmed = input.trim_start();
    let mut chars = trimmed.char_indices();
    let Some((_, quote)) = chars.next() else {
        return Err("expected string literal".to_string());
    };
    if quote != '"' && quote != '\'' {
        return Err("expected string literal".to_string());
    }

    let mut out = String::new();
    let mut escape = false;
    for (idx, ch) in chars {
        if escape {
            let translated = match ch {
                'n' => '\n',
                'r' => '\r',
                't' => '\t',
                '\\' => '\\',
                '"' => '"',
                '\'' => '\'',
                other => other,
            };
            out.push(translated);
            escape = false;
            continue;
        }

        if ch == '\\' {
            escape = true;
            continue;
        }

        if ch == quote {
            let remainder = &trimmed[idx + ch.len_utf8()..];
            return Ok((out, remainder));
        }

        out.push(ch);
    }

    Err("unterminated string literal".to_string())
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_string_literal(expr: &str) -> Result<String, String> {
    let trimmed = expr.trim().trim_end_matches(';').trim();
    let (value, remainder) = parse_typescript_string_literal_prefix(trimmed)?;
    if !remainder.trim().is_empty() {
        return Err(format!(
            "unexpected trailing tokens after string literal: {remainder}"
        ));
    }
    Ok(value)
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_value_expr(expr: &str) -> Result<Option<JsonValue>, String> {
    let trimmed = expr.trim().trim_end_matches(';').trim();
    if trimmed.eq_ignore_ascii_case("null") || trimmed.eq_ignore_ascii_case("undefined") {
        return Ok(None);
    }
    if trimmed.eq_ignore_ascii_case("true") {
        return Ok(Some(JsonValue::Bool(true)));
    }
    if trimmed.eq_ignore_ascii_case("false") {
        return Ok(Some(JsonValue::Bool(false)));
    }

    if trimmed.starts_with('"') || trimmed.starts_with('\'') {
        let value = parse_typescript_string_literal(trimmed)?;
        return Ok(Some(JsonValue::String(value)));
    }

    if let Ok(int_value) = trimmed.parse::<i64>() {
        return Ok(Some(JsonValue::from(int_value)));
    }

    if let Ok(float_value) = trimmed.parse::<f64>() {
        let num = serde_json::Number::from_f64(float_value)
            .ok_or_else(|| format!("invalid numeric literal: {trimmed}"))?;
        return Ok(Some(JsonValue::Number(num)));
    }

    Err(format!("unsupported TypeScript literal: {trimmed}"))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_call_args(input: &str) -> Result<(String, &str), String> {
    let trimmed = input.trim_start();
    let mut in_string: Option<char> = None;
    let mut escape = false;
    let mut paren_depth: i32 = 0;
    let mut bracket_depth: i32 = 0;
    let mut brace_depth: i32 = 0;

    for (idx, ch) in trimmed.char_indices() {
        if let Some(quote) = in_string {
            if escape {
                escape = false;
                continue;
            }
            if ch == '\\' {
                escape = true;
                continue;
            }
            if ch == quote {
                in_string = None;
            }
            continue;
        }

        match ch {
            '"' | '\'' => in_string = Some(ch),
            '(' => paren_depth += 1,
            ')' => {
                if paren_depth == 0 && bracket_depth == 0 && brace_depth == 0 {
                    let args = trimmed[..idx].to_string();
                    let rest = &trimmed[idx + ch.len_utf8()..];
                    return Ok((args, rest));
                }
                paren_depth -= 1;
            }
            '[' => bracket_depth += 1,
            ']' => bracket_depth = (bracket_depth - 1).max(0),
            '{' => brace_depth += 1,
            '}' => brace_depth = (brace_depth - 1).max(0),
            _ => {}
        }
    }

    Err("unterminated function call".to_string())
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_cell_selector(expr: &str) -> Result<(usize, usize), String> {
    let (start_row, start_col, _end_row, _end_col) = parse_typescript_range_selector(expr)?;
    Ok((start_row, start_col))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_range_selector(expr: &str) -> Result<(usize, usize, usize, usize), String> {
    if let Some(idx) = expr.rfind(".getRange(") {
        let after = &expr[idx + ".getRange(".len()..];
        let (addr, remainder) = parse_typescript_string_literal_prefix(after)?;
        let remainder = remainder.trim_start();
        if !remainder.starts_with(')') {
            return Err(format!("unsupported getRange expression: {expr}"));
        }
        return parse_typescript_a1_range(&addr);
    }

    if let Some(idx) = expr.rfind(".range(") {
        let after = &expr[idx + ".range(".len()..];
        let (addr, remainder) = parse_typescript_string_literal_prefix(after)?;
        let remainder = remainder.trim_start();
        if !remainder.starts_with(')') {
            return Err(format!("unsupported range expression: {expr}"));
        }
        return parse_typescript_a1_range(&addr);
    }

    if let Some(idx) = expr.rfind(".getCell(") {
        let after = &expr[idx + ".getCell(".len()..];
        let (args, _remainder) = parse_typescript_call_args(after)?;
        let mut parts = args.split(',');
        let row_str = parts.next().unwrap_or("").trim();
        let col_str = parts.next().unwrap_or("").trim();
        if parts.next().is_some() {
            return Err(format!("unsupported getCell expression: {expr}"));
        }
        let row = row_str
            .parse::<usize>()
            .map_err(|_| format!("invalid row in getCell(): {row_str:?}"))?;
        let col = col_str
            .parse::<usize>()
            .map_err(|_| format!("invalid col in getCell(): {col_str:?}"))?;
        return Ok((row, col, row, col));
    }

    if let Some(idx) = expr.rfind(".cell(") {
        let after = &expr[idx + ".cell(".len()..];
        let (args, _remainder) = parse_typescript_call_args(after)?;
        let mut parts = args.split(',');
        let row_str = parts.next().unwrap_or("").trim();
        let col_str = parts.next().unwrap_or("").trim();
        if parts.next().is_some() {
            return Err(format!("unsupported cell expression: {expr}"));
        }
        let row_1 = match row_str.parse::<usize>() {
            Ok(v) if v > 0 => v - 1,
            _ => return Err(format!("invalid row in cell(): {row_str:?}")),
        };
        let col_1 = match col_str.parse::<usize>() {
            Ok(v) if v > 0 => v - 1,
            _ => return Err(format!("invalid col in cell(): {col_str:?}")),
        };
        return Ok((row_1, col_1, row_1, col_1));
    }

    Err(format!("unsupported cell selector: {expr}"))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_a1_range(addr: &str) -> Result<(usize, usize, usize, usize), String> {
    let addr = addr.trim();
    let mut parts = addr.split(':');
    let start_raw = parts.next().unwrap_or("").trim();
    let end_raw = parts.next().unwrap_or("").trim();
    if parts.next().is_some() {
        return Err(format!("invalid A1 range: {addr:?}"));
    }

    let start = formula_engine::eval::parse_a1(start_raw)
        .map_err(|e| format!("invalid A1 address {start_raw:?}: {e}"))?;
    let end = if end_raw.is_empty() {
        start
    } else {
        formula_engine::eval::parse_a1(end_raw)
            .map_err(|e| format!("invalid A1 address {end_raw:?}: {e}"))?
    };

    let start_row = std::cmp::min(start.row, end.row) as usize;
    let end_row = std::cmp::max(start.row, end.row) as usize;
    let start_col = std::cmp::min(start.col, end.col) as usize;
    let end_col = std::cmp::max(start.col, end.col) as usize;
    Ok((start_row, start_col, end_row, end_col))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_value_prefix(input: &str) -> Result<(Option<JsonValue>, &str), String> {
    let trimmed = input.trim_start();
    if trimmed.is_empty() {
        return Err("expected literal".to_string());
    }

    let lower = trimmed.to_ascii_lowercase();
    for (token, value) in [
        ("null", None),
        ("undefined", None),
        ("true", Some(JsonValue::Bool(true))),
        ("false", Some(JsonValue::Bool(false))),
    ] {
        if lower.starts_with(token) {
            let remainder = &trimmed[token.len()..];
            let next = remainder.chars().next();
            if matches!(next, Some(ch) if ch.is_ascii_alphanumeric() || ch == '_') {
                // e.g. "nullish" -> not a token.
            } else {
                return Ok((value, remainder));
            }
        }
    }

    if trimmed.starts_with('"') || trimmed.starts_with('\'') {
        let (value, remainder) = parse_typescript_string_literal_prefix(trimmed)?;
        return Ok((Some(JsonValue::String(value)), remainder));
    }

    // Parse a simple number literal (no exponent).
    let bytes = trimmed.as_bytes();
    let mut idx = 0usize;
    if matches!(bytes.get(idx), Some(b'+' | b'-')) {
        idx += 1;
    }
    let start_digits = idx;
    while matches!(bytes.get(idx), Some(b) if b.is_ascii_digit()) {
        idx += 1;
    }
    if idx == start_digits {
        return Err(format!("unsupported TypeScript literal: {trimmed}"));
    }
    if matches!(bytes.get(idx), Some(b'.')) {
        idx += 1;
        let start_frac = idx;
        while matches!(bytes.get(idx), Some(b) if b.is_ascii_digit()) {
            idx += 1;
        }
        if idx == start_frac {
            return Err(format!("invalid numeric literal: {trimmed}"));
        }
    }
    let literal = &trimmed[..idx];
    let remainder = &trimmed[idx..];
    if literal.contains('.') {
        let float_value = literal
            .parse::<f64>()
            .map_err(|_| format!("invalid numeric literal: {literal}"))?;
        let num = serde_json::Number::from_f64(float_value)
            .ok_or_else(|| format!("invalid numeric literal: {literal}"))?;
        return Ok((Some(JsonValue::Number(num)), remainder));
    }
    let int_value = literal
        .parse::<i64>()
        .map_err(|_| format!("invalid numeric literal: {literal}"))?;
    Ok((Some(JsonValue::from(int_value)), remainder))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_value_matrix(expr: &str) -> Result<Vec<Vec<Option<JsonValue>>>, String> {
    use crate::resource_limits::{MAX_RANGE_CELLS_PER_CALL, MAX_RANGE_DIM};

    let mut rest = expr.trim_start();
    rest = rest
        .strip_prefix('[')
        .ok_or_else(|| "expected matrix literal like [[1,2],[3,4]]".to_string())?;

    let mut rows: Vec<Vec<Option<JsonValue>>> = Vec::new();
    let mut total_cells = 0usize;
    loop {
        rest = rest.trim_start();
        if let Some(next) = rest.strip_prefix(']') {
            rest = next;
            break;
        }
        rest = rest
            .strip_prefix('[')
            .ok_or_else(|| "expected row literal like [1,2]".to_string())?;

        if rows.len() >= MAX_RANGE_DIM {
            return Err(format!(
                "matrix literal is too large (max {MAX_RANGE_DIM} rows)"
            ));
        }

        let mut row: Vec<Option<JsonValue>> = Vec::new();
        loop {
            rest = rest.trim_start();
            if let Some(next) = rest.strip_prefix(']') {
                rest = next;
                break;
            }

            if row.len() >= MAX_RANGE_DIM {
                return Err(format!(
                    "matrix literal row is too large (max {MAX_RANGE_DIM} columns)"
                ));
            }

            total_cells = total_cells.saturating_add(1);
            if total_cells > MAX_RANGE_CELLS_PER_CALL {
                return Err(format!(
                    "matrix literal is too large (max {MAX_RANGE_CELLS_PER_CALL} cells)"
                ));
            }
            let (value, remainder) = parse_typescript_value_prefix(rest)?;
            row.push(value);
            rest = remainder.trim_start();
            if let Some(next) = rest.strip_prefix(',') {
                rest = next;
                continue;
            }
            if let Some(next) = rest.strip_prefix(']') {
                rest = next;
                break;
            }
            return Err(format!("expected ',' or ']' in row literal, got {rest:?}"));
        }
        rows.push(row);

        rest = rest.trim_start();
        if let Some(next) = rest.strip_prefix(',') {
            rest = next;
            continue;
        }
        if let Some(next) = rest.strip_prefix(']') {
            rest = next;
            break;
        }
        return Err(format!(
            "expected ',' or ']' after row literal, got {rest:?}"
        ));
    }

    if !rest.trim().trim_end_matches(';').trim().is_empty() {
        return Err(format!(
            "unexpected trailing tokens after matrix literal: {rest}"
        ));
    }

    Ok(rows)
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_formulas_matrix(expr: &str) -> Result<Vec<Vec<Option<String>>>, String> {
    let matrix = parse_typescript_value_matrix(expr)?;
    let mut out: Vec<Vec<Option<String>>> = Vec::new();
    for row in matrix {
        let mut row_out: Vec<Option<String>> = Vec::new();
        for value in row {
            match value {
                None => row_out.push(None),
                Some(JsonValue::String(s)) => row_out.push(Some(s)),
                Some(other) => {
                    return Err(format!(
                        "expected formula string literal or null, got {other}"
                    ))
                }
            }
        }
        out.push(row_out);
    }
    Ok(out)
}

#[cfg(any(feature = "desktop", test))]
#[derive(Clone, Debug)]
enum TypeScriptBinding {
    Scalar(Option<JsonValue>),
    Matrix(Vec<Vec<Option<JsonValue>>>),
}

#[cfg(any(feature = "desktop", test))]
fn is_typescript_identifier(input: &str) -> bool {
    let mut chars = input.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !(first.is_ascii_alphabetic() || first == '_') {
        return false;
    }
    chars.all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
}

#[cfg(any(feature = "desktop", test))]
fn split_typescript_top_level_commas(input: &str) -> Vec<&str> {
    let trimmed = input.trim();
    let mut parts: Vec<&str> = Vec::new();
    let mut start = 0usize;

    let mut in_string: Option<char> = None;
    let mut escape = false;
    let mut paren_depth: i32 = 0;
    let mut bracket_depth: i32 = 0;
    let mut brace_depth: i32 = 0;

    for (idx, ch) in trimmed.char_indices() {
        if let Some(quote) = in_string {
            if escape {
                escape = false;
                continue;
            }
            if ch == '\\' {
                escape = true;
                continue;
            }
            if ch == quote {
                in_string = None;
            }
            continue;
        }

        match ch {
            '"' | '\'' => in_string = Some(ch),
            '(' => paren_depth += 1,
            ')' => paren_depth = (paren_depth - 1).max(0),
            '[' => bracket_depth += 1,
            ']' => bracket_depth = (bracket_depth - 1).max(0),
            '{' => brace_depth += 1,
            '}' => brace_depth = (brace_depth - 1).max(0),
            ',' if paren_depth == 0 && bracket_depth == 0 && brace_depth == 0 => {
                parts.push(trimmed[start..idx].trim());
                start = idx + ch.len_utf8();
            }
            _ => {}
        }
    }

    parts.push(trimmed[start..].trim());
    parts
}

#[cfg(any(feature = "desktop", test))]
fn resolve_typescript_scalar_expr(
    expr: &str,
    bindings: &std::collections::HashMap<String, TypeScriptBinding>,
) -> Result<Option<JsonValue>, String> {
    match parse_typescript_value_expr(expr) {
        Ok(value) => Ok(value),
        Err(parse_err) => {
            let ident = expr.trim().trim_end_matches(';').trim();
            if is_typescript_identifier(ident) {
                if let Some(TypeScriptBinding::Scalar(value)) = bindings.get(ident) {
                    return Ok(value.clone());
                }
            }
            Err(parse_err)
        }
    }
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_array_from_fill_matrix(
    expr: &str,
    bindings: &std::collections::HashMap<String, TypeScriptBinding>,
) -> Result<Vec<Vec<Option<JsonValue>>>, String> {
    use crate::resource_limits::{MAX_RANGE_CELLS_PER_CALL, MAX_RANGE_DIM};

    let trimmed = expr.trim().trim_end_matches(';').trim();
    let after = trimmed
        .strip_prefix("Array.from")
        .ok_or_else(|| format!("unsupported matrix expression: {trimmed}"))?;
    let after = after.trim_start();
    let after = after
        .strip_prefix('(')
        .ok_or_else(|| format!("unsupported matrix expression: {trimmed}"))?;

    let (args, remainder) = parse_typescript_call_args(after)?;
    if !remainder.trim().is_empty() {
        return Err(format!(
            "unexpected trailing tokens after Array.from(...): {remainder}"
        ));
    }

    let parts = split_typescript_top_level_commas(&args);
    if parts.len() != 2 {
        return Err(format!("unsupported Array.from(...) arguments: {args}"));
    }

    let length_arg = parts[0];
    let mut rest = length_arg.trim();
    rest = rest
        .strip_prefix('{')
        .ok_or_else(|| format!("unsupported Array.from length arg: {length_arg}"))?;
    rest = rest
        .strip_suffix('}')
        .ok_or_else(|| format!("unsupported Array.from length arg: {length_arg}"))?;
    rest = rest.trim_start();
    rest = rest
        .strip_prefix("length")
        .ok_or_else(|| format!("unsupported Array.from length arg: {length_arg}"))?;
    rest = rest.trim_start();
    rest = rest
        .strip_prefix(':')
        .ok_or_else(|| format!("unsupported Array.from length arg: {length_arg}"))?;
    rest = rest.trim_start();
    let digits = rest
        .chars()
        .take_while(|ch| ch.is_ascii_digit())
        .collect::<String>();
    if digits.is_empty() {
        return Err(format!("unsupported Array.from length arg: {length_arg}"));
    }
    let rows = digits
        .parse::<usize>()
        .map_err(|_| format!("invalid Array.from length: {digits}"))?;
    if rows == 0 {
        return Err("Array.from length must be > 0".to_string());
    }

    let fill_arg = parts[1];
    let fill_str = fill_arg.trim();
    let array_idx = fill_str
        .find("Array(")
        .ok_or_else(|| format!("unsupported Array.from fill arg: {fill_arg}"))?;
    let after_array = &fill_str[array_idx + "Array(".len()..];
    let after_array = after_array.trim_start();
    let digits = after_array
        .chars()
        .take_while(|ch| ch.is_ascii_digit())
        .collect::<String>();
    if digits.is_empty() {
        return Err(format!("unsupported Array.from fill arg: {fill_arg}"));
    }
    let cols = digits
        .parse::<usize>()
        .map_err(|_| format!("invalid Array(...) length: {digits}"))?;
    if cols == 0 {
        return Err("Array(...) length must be > 0".to_string());
    }

    if rows > MAX_RANGE_DIM || cols > MAX_RANGE_DIM {
        return Err(format!(
            "Array.from matrix is too large ({rows}x{cols}; max dimension {MAX_RANGE_DIM})"
        ));
    }
    let cell_count = (rows as u128) * (cols as u128);
    if cell_count > MAX_RANGE_CELLS_PER_CALL as u128 {
        return Err(format!(
            "Array.from matrix is too large ({rows}x{cols}; max {MAX_RANGE_CELLS_PER_CALL} cells)"
        ));
    }

    let mut rest = after_array[digits.len()..].trim_start();
    rest = rest
        .strip_prefix(')')
        .ok_or_else(|| format!("unsupported Array.from fill arg: {fill_arg}"))?;
    rest = rest.trim_start();
    rest = rest
        .strip_prefix(".fill(")
        .ok_or_else(|| format!("unsupported Array.from fill arg: {fill_arg}"))?;

    let (fill_expr, remainder) = parse_typescript_call_args(rest)?;
    if !remainder.trim().is_empty() {
        return Err(format!(
            "unexpected trailing tokens after Array(...).fill(...): {remainder}"
        ));
    }

    let fill_value = resolve_typescript_scalar_expr(&fill_expr, bindings)?;
    let mut matrix: Vec<Vec<Option<JsonValue>>> = Vec::new();
    for _ in 0..rows {
        matrix.push((0..cols).map(|_| fill_value.clone()).collect());
    }
    Ok(matrix)
}

#[cfg(any(feature = "desktop", test))]
fn resolve_typescript_value_matrix_expr(
    expr: &str,
    bindings: &std::collections::HashMap<String, TypeScriptBinding>,
) -> Result<Vec<Vec<Option<JsonValue>>>, String> {
    let trimmed = expr.trim().trim_end_matches(';').trim();
    if trimmed.starts_with('[') {
        return parse_typescript_value_matrix(trimmed);
    }
    if trimmed.starts_with("Array.from") {
        return parse_typescript_array_from_fill_matrix(trimmed, bindings);
    }
    if is_typescript_identifier(trimmed) {
        match bindings.get(trimmed) {
            Some(TypeScriptBinding::Matrix(matrix)) => return Ok(matrix.clone()),
            Some(TypeScriptBinding::Scalar(_)) => {
                return Err(format!("expected matrix expression, got scalar {trimmed}"))
            }
            None => return Err(format!("unknown identifier: {trimmed}")),
        }
    }
    Err(format!(
        "unsupported TypeScript matrix expression: {trimmed}"
    ))
}

#[cfg(any(feature = "desktop", test))]
fn resolve_typescript_formulas_matrix_expr(
    expr: &str,
    bindings: &std::collections::HashMap<String, TypeScriptBinding>,
) -> Result<Vec<Vec<Option<String>>>, String> {
    let trimmed = expr.trim().trim_end_matches(';').trim();
    if trimmed.starts_with('[') {
        return parse_typescript_formulas_matrix(trimmed);
    }

    let matrix = resolve_typescript_value_matrix_expr(trimmed, bindings)?;
    let mut out: Vec<Vec<Option<String>>> = Vec::new();
    for row in matrix {
        let mut row_out: Vec<Option<String>> = Vec::new();
        for value in row {
            match value {
                None => row_out.push(None),
                Some(JsonValue::String(s)) => row_out.push(Some(s)),
                Some(other) => {
                    return Err(format!(
                        "expected formula string literal or null, got {other}"
                    ))
                }
            }
        }
        out.push(row_out);
    }
    Ok(out)
}

#[cfg(any(feature = "desktop", test))]
fn run_typescript_migration_script(state: &mut AppState, code: &str) -> TypeScriptRunResult {
    use std::collections::HashMap;

    let active_sheet_id = match state.get_workbook() {
        Ok(workbook) => {
            let active_index = state.macro_runtime_context().active_sheet;
            workbook
                .sheets
                .get(active_index)
                .or_else(|| workbook.sheets.first())
                .map(|s| s.id.clone())
                .unwrap_or_else(|| "Sheet1".to_string())
        }
        Err(err) => {
            return TypeScriptRunResult {
                ok: false,
                updates: Vec::new(),
                error: Some(err.to_string()),
            }
        }
    };

    let mut updates = Vec::<CellUpdateData>::new();
    let mut error: Option<String> = None;
    let mut bindings: HashMap<String, TypeScriptBinding> = HashMap::new();

    for raw_line in code.lines() {
        let line = raw_line.trim();
        if line.is_empty() {
            continue;
        }
        if line.starts_with("//") {
            continue;
        }
        if line.starts_with("export ") || line.starts_with("import ") {
            continue;
        }
        if line == "{" || line == "}" {
            continue;
        }
        if line.starts_with("const ") || line.starts_with("let ") || line.starts_with("var ") {
            let rest = if let Some(rest) = line.strip_prefix("const ") {
                rest
            } else if let Some(rest) = line.strip_prefix("let ") {
                rest
            } else {
                line.strip_prefix("var ").unwrap_or("")
            };

            if let Some((name_raw, expr_raw)) = rest.split_once('=') {
                let name = name_raw.trim();
                let expr = expr_raw.trim().trim_end_matches(';').trim();
                if is_typescript_identifier(name) {
                    if let Ok(value) = parse_typescript_value_expr(expr) {
                        bindings.insert(name.to_string(), TypeScriptBinding::Scalar(value));
                    } else if expr.trim_start().starts_with('[') {
                        if let Ok(matrix) = parse_typescript_value_matrix(expr) {
                            bindings.insert(name.to_string(), TypeScriptBinding::Matrix(matrix));
                        }
                    } else if expr.trim_start().starts_with("Array.from") {
                        if let Ok(matrix) = parse_typescript_array_from_fill_matrix(expr, &bindings)
                        {
                            bindings.insert(name.to_string(), TypeScriptBinding::Matrix(matrix));
                        }
                    }
                }
            }

            continue;
        }

        if let Some(idx) = line.find(".setValue(") {
            let target_expr = line[..idx].trim();
            let after = &line[idx + ".setValue(".len()..];
            match parse_typescript_call_args(after) {
                Ok((args, remainder)) => {
                    if !remainder
                        .trim_start()
                        .trim_start_matches(';')
                        .trim()
                        .is_empty()
                    {
                        error = Some(format!("unsupported setValue call: {line}"));
                        break;
                    }
                    match parse_typescript_range_selector(target_expr) {
                        Ok((start_row, start_col, end_row, end_col)) => {
                            if start_row != end_row || start_col != end_col {
                                error = Some(format!(
                                    "setValue is only valid for single cells (got range {start_row},{start_col}..{end_row},{end_col})"
                                ));
                                break;
                            }
                            match resolve_typescript_scalar_expr(&args, &bindings) {
                                Ok(value) => {
                                    match state.set_cell(
                                        &active_sheet_id,
                                        start_row,
                                        start_col,
                                        value,
                                        None,
                                    ) {
                                        Ok(mut changed) => updates.append(&mut changed),
                                        Err(e) => {
                                            error = Some(e.to_string());
                                            break;
                                        }
                                    }
                                }
                                Err(e) => {
                                    error = Some(e);
                                    break;
                                }
                            }
                        }
                        Err(e) => {
                            error = Some(e);
                            break;
                        }
                    }
                    continue;
                }
                Err(e) => {
                    error = Some(e);
                    break;
                }
            }
        }

        if let Some(idx) = line.find(".setValues(") {
            let target_expr = line[..idx].trim();
            let after = &line[idx + ".setValues(".len()..];
            match parse_typescript_call_args(after) {
                Ok((args, remainder)) => {
                    if !remainder
                        .trim_start()
                        .trim_start_matches(';')
                        .trim()
                        .is_empty()
                    {
                        error = Some(format!("unsupported setValues call: {line}"));
                        break;
                    }

                    let (start_row, start_col, end_row, end_col) =
                        match parse_typescript_range_selector(target_expr) {
                            Ok(coords) => coords,
                            Err(e) => {
                                error = Some(e);
                                break;
                            }
                        };

                    let matrix = match resolve_typescript_value_matrix_expr(&args, &bindings) {
                        Ok(m) => m,
                        Err(e) => {
                            error = Some(e);
                            break;
                        }
                    };

                    let row_count = end_row - start_row + 1;
                    let col_count = end_col - start_col + 1;
                    if matrix.len() != row_count || matrix.iter().any(|row| row.len() != col_count)
                    {
                        error = Some(format!(
                            "setValues expected {row_count}x{col_count} matrix for range ({start_row},{start_col})..({end_row},{end_col}), got {}x{}",
                            matrix.len(),
                            matrix.first().map(|row| row.len()).unwrap_or(0)
                        ));
                        break;
                    }

                    let mut payload: Vec<Vec<(Option<JsonValue>, Option<String>)>> = Vec::new();
                    for row in matrix {
                        payload.push(
                            row.into_iter()
                                .map(|value| (value, None))
                                .collect::<Vec<_>>(),
                        );
                    }

                    match state.set_range(
                        &active_sheet_id,
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        payload,
                    ) {
                        Ok(mut changed) => updates.append(&mut changed),
                        Err(e) => {
                            error = Some(e.to_string());
                            break;
                        }
                    }

                    continue;
                }
                Err(e) => {
                    error = Some(e);
                    break;
                }
            }
        }

        if let Some(idx) = line.find(".setFormulas(") {
            let target_expr = line[..idx].trim();
            let after = &line[idx + ".setFormulas(".len()..];
            match parse_typescript_call_args(after) {
                Ok((args, remainder)) => {
                    if !remainder
                        .trim_start()
                        .trim_start_matches(';')
                        .trim()
                        .is_empty()
                    {
                        error = Some(format!("unsupported setFormulas call: {line}"));
                        break;
                    }
                    let (start_row, start_col, end_row, end_col) =
                        match parse_typescript_range_selector(target_expr) {
                            Ok(coords) => coords,
                            Err(e) => {
                                error = Some(e);
                                break;
                            }
                        };

                    let matrix = match resolve_typescript_formulas_matrix_expr(&args, &bindings) {
                        Ok(m) => m,
                        Err(e) => {
                            error = Some(e);
                            break;
                        }
                    };

                    let row_count = end_row - start_row + 1;
                    let col_count = end_col - start_col + 1;
                    if matrix.len() != row_count || matrix.iter().any(|row| row.len() != col_count)
                    {
                        error = Some(format!(
                            "setFormulas expected {row_count}x{col_count} matrix for range ({start_row},{start_col})..({end_row},{end_col}), got {}x{}",
                            matrix.len(),
                            matrix.first().map(|row| row.len()).unwrap_or(0)
                        ));
                        break;
                    }

                    let mut payload: Vec<Vec<(Option<JsonValue>, Option<String>)>> = Vec::new();
                    for row in matrix {
                        payload.push(
                            row.into_iter()
                                .map(|formula| (None, formula))
                                .collect::<Vec<_>>(),
                        );
                    }

                    match state.set_range(
                        &active_sheet_id,
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        payload,
                    ) {
                        Ok(mut changed) => updates.append(&mut changed),
                        Err(e) => {
                            error = Some(e.to_string());
                            break;
                        }
                    }

                    continue;
                }
                Err(e) => {
                    error = Some(e);
                    break;
                }
            }
        }

        let Some((lhs_raw, rhs_raw)) = line.split_once('=') else {
            // Ignore non-assignment statements (loops, conditionals, etc.).
            continue;
        };
        let lhs_raw = lhs_raw.trim();
        let rhs_raw = rhs_raw.trim();

        let (assign_kind, lhs_target) = if let Some(prefix) = lhs_raw.strip_suffix(".value") {
            ("value", prefix.trim())
        } else if let Some(prefix) = lhs_raw.strip_suffix(".formula") {
            ("formula", prefix.trim())
        } else {
            continue;
        };

        let (row, col) = match parse_typescript_cell_selector(lhs_target) {
            Ok(coords) => coords,
            Err(_) => continue,
        };

        let result = match assign_kind {
            "value" => match resolve_typescript_scalar_expr(rhs_raw, &bindings) {
                Ok(value) => state.set_cell(&active_sheet_id, row, col, value, None),
                Err(e) => Err(AppStateError::WhatIf(e)),
            },
            "formula" => match resolve_typescript_scalar_expr(rhs_raw, &bindings) {
                Ok(None) => state.set_cell(&active_sheet_id, row, col, None, None),
                Ok(Some(JsonValue::String(formula))) => {
                    state.set_cell(&active_sheet_id, row, col, None, Some(formula))
                }
                Ok(Some(other)) => Err(AppStateError::WhatIf(format!(
                    "expected formula string literal or null, got {other}"
                ))),
                Err(e) => Err(AppStateError::WhatIf(e)),
            },
            _ => continue,
        };

        match result {
            Ok(mut changed) => updates.append(&mut changed),
            Err(e) => {
                error = Some(e.to_string());
                break;
            }
        }
    }

    // De-dupe updates by last write (keep report stable).
    let mut out: Vec<CellUpdateData> = Vec::new();
    let mut idx_by_key: HashMap<(String, usize, usize), usize> = HashMap::new();
    for update in updates {
        let key = (update.sheet_id.clone(), update.row, update.col);
        if let Some(idx) = idx_by_key.get(&key).copied() {
            out[idx] = update;
        } else {
            idx_by_key.insert(key, out.len());
            out.push(update);
        }
    }

    TypeScriptRunResult {
        ok: error.is_none(),
        updates: out.into_iter().map(cell_update_from_state).collect(),
        error,
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn validate_vba_migration(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    macro_id: LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>,
    target: MigrationTarget,
    code: IpcScriptCode,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MigrationValidationReport, String> {
    ipc_origin::ensure_main_window(window.label(), "macro execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro execution", ipc_origin::Verb::Is)?;

    let macro_id = macro_id.into_inner();
    let code = code.into_inner();
    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        use crate::macros::MacroExecutionOptions;
        use std::collections::BTreeSet;

        let (vba_blocked_result, workbook, macro_ctx) = {
            let mut state = shared.lock().unwrap();
            let mut trust_store = trust_shared.lock().unwrap();
            trust_store.ensure_loaded();

            let blocked = {
                let workbook_id = workbook_id.as_ref().map(|id| id.as_ref());
                let workbook = state.get_workbook_mut().map_err(app_error)?;
                enforce_macro_trust(workbook, workbook_id, &trust_store)?
            };

            let vba_blocked_result = blocked.map(macro_blocked_result);
            let macro_ctx = state.macro_runtime_context();
            let workbook = state.get_workbook_mut().map_err(app_error)?.clone();
            (vba_blocked_result, workbook, macro_ctx)
        };

        let mut vba_state = AppState::new();
        vba_state.load_workbook(workbook.clone());
        vba_state
            .set_macro_runtime_context(macro_ctx)
            .map_err(|e| e.to_string())?;

        let mut script_state = AppState::new();
        script_state.load_workbook(workbook);
        script_state
            .set_macro_runtime_context(macro_ctx)
            .map_err(|e| e.to_string())?;

        let vba = if let Some(blocked_result) = vba_blocked_result {
            blocked_result
        } else {
            match vba_state.run_macro(&macro_id, MacroExecutionOptions::default()) {
                Ok(outcome) => macro_result_from_outcome(outcome),
                Err(err) => MacroRunResult {
                    ok: false,
                    output: Vec::new(),
                    updates: Vec::new(),
                    error: Some(MacroError {
                        message: err.to_string(),
                        code: Some("macro_error".to_string()),
                        blocked: None,
                    }),
                    permission_request: None,
                },
            }
        };

        let mut python = None;
        let mut typescript = None;

        match target {
            MigrationTarget::Python => {
                let python_context = {
                    let workbook = script_state.get_workbook().map_err(|e| e.to_string())?;
                    let fallback_sheet_id = workbook
                        .sheets
                        .first()
                        .map(|s| s.id.clone())
                        .ok_or_else(|| "workbook contains no sheets".to_string())?;
                    let active_sheet_id = workbook
                        .sheets
                        .get(macro_ctx.active_sheet)
                        .map(|s| s.id.clone())
                        .unwrap_or_else(|| fallback_sheet_id.clone());

                    let selection =
                        macro_ctx
                            .selection
                            .unwrap_or(formula_vba_runtime::VbaRangeRef {
                                sheet: macro_ctx.active_sheet,
                                start_row: macro_ctx.active_cell.0,
                                start_col: macro_ctx.active_cell.1,
                                end_row: macro_ctx.active_cell.0,
                                end_col: macro_ctx.active_cell.1,
                            });
                    let selection_sheet_id = workbook
                        .sheets
                        .get(selection.sheet)
                        .map(|s| s.id.clone())
                        .unwrap_or_else(|| active_sheet_id.clone());
                    PythonRunContext {
                        active_sheet_id: Some(active_sheet_id.clone()),
                        selection: Some(PythonSelection {
                            sheet_id: selection_sheet_id,
                            start_row: selection.start_row.saturating_sub(1) as usize,
                            start_col: selection.start_col.saturating_sub(1) as usize,
                            end_row: selection.end_row.saturating_sub(1) as usize,
                            end_col: selection.end_col.saturating_sub(1) as usize,
                        }),
                    }
                };
                python = Some(
                    crate::python::run_python_script(
                        &mut script_state,
                        &code,
                        None,
                        None,
                        None,
                        Some(python_context),
                    )
                    .map_err(|e| e.to_string())?,
                );
            }
            MigrationTarget::TypeScript => {
                typescript = Some(run_typescript_migration_script(&mut script_state, &code));
            }
        };

        let mut mismatches = Vec::new();

        let mut touched = BTreeSet::<(String, usize, usize)>::new();
        for update in &vba.updates {
            touched.insert((update.sheet_id.clone(), update.row, update.col));
        }
        if let Some(python_run) = python.as_ref() {
            for update in &python_run.updates {
                touched.insert((update.sheet_id.clone(), update.row, update.col));
            }
        }
        if let Some(ts_run) = typescript.as_ref() {
            for update in &ts_run.updates {
                touched.insert((update.sheet_id.clone(), update.row, update.col));
            }
        }

        for (sheet_id, row, col) in touched {
            let vba_cell = cell_value_from_state(&vba_state, &sheet_id, row, col)?;
            let script_cell = cell_value_from_state(&script_state, &sheet_id, row, col)?;
            if vba_cell != script_cell {
                mismatches.push(MigrationValidationMismatch {
                    sheet_id,
                    row,
                    col,
                    vba: vba_cell,
                    script: script_cell,
                });
            }
        }

        let script_ok = match target {
            MigrationTarget::Python => python.as_ref().map(|r| r.ok).unwrap_or(false),
            MigrationTarget::TypeScript => typescript.as_ref().map(|r| r.ok).unwrap_or(false),
        };

        let mut error_messages: Vec<String> = Vec::new();
        if !vba.ok {
            if let Some(err) = vba.error.as_ref() {
                error_messages.push(err.message.clone());
            } else {
                error_messages.push("VBA macro failed".to_string());
            }
        }
        if !script_ok {
            match target {
                MigrationTarget::Python => {
                    if let Some(run) = python.as_ref() {
                        if let Some(err) = run.error.as_ref() {
                            error_messages.push(err.message.clone());
                        } else {
                            error_messages.push("Python script failed".to_string());
                        }
                    }
                }
                MigrationTarget::TypeScript => {
                    if let Some(run) = typescript.as_ref() {
                        if let Some(err) = run.error.as_ref() {
                            error_messages.push(err.clone());
                        } else {
                            error_messages.push("TypeScript migration failed".to_string());
                        }
                    }
                }
            }
        }
        let error = if error_messages.is_empty() {
            None
        } else {
            Some(error_messages.join(" | "))
        };

        let ok = vba.ok && script_ok && mismatches.is_empty();

        Ok(MigrationValidationReport {
            ok,
            macro_id,
            target,
            mismatches,
            vba,
            python,
            typescript,
            error,
        })
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(any(feature = "desktop", test))]
fn macro_error_code_from_message(message: &str) -> Option<&'static str> {
    if message.starts_with("macro produced too many cell updates") {
        return Some("macro_updates_limit_exceeded");
    }
    if message.starts_with("cell value string is too large") {
        return Some("macro_cell_value_too_large");
    }
    if message.starts_with("cell formula is too large") {
        return Some("macro_cell_formula_too_large");
    }
    None
}

#[cfg(any(feature = "desktop", test))]
fn macro_result_from_outcome(outcome: crate::macros::MacroExecutionOutcome) -> MacroRunResult {
    MacroRunResult {
        ok: outcome.ok,
        output: outcome.output,
        updates: outcome
            .updates
            .into_iter()
            .map(cell_update_from_state)
            .collect(),
        error: outcome.error.map(|message| {
            let code = macro_error_code_from_message(&message).map(|c| c.to_string());
            MacroError {
                message,
                code,
                blocked: None,
            }
        }),
        permission_request: outcome.permission_request,
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn fire_workbook_open(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    permissions: Option<LimitedMacroPermissions>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    ipc_origin::ensure_main_window(window.label(), "macro execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro execution", ipc_origin::Verb::Is)?;

    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let mut trust_store = trust_shared.lock().unwrap();
            trust_store.ensure_loaded();
            let workbook_id = workbook_id.as_ref().map(|id| id.as_ref());
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }
        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.map(|p| p.0).unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .fire_workbook_open(options)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn fire_workbook_before_close(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    permissions: Option<LimitedMacroPermissions>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    ipc_origin::ensure_main_window(window.label(), "macro execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro execution", ipc_origin::Verb::Is)?;

    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let mut trust_store = trust_shared.lock().unwrap();
            trust_store.ensure_loaded();
            let workbook_id = workbook_id.as_ref().map(|id| id.as_ref());
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }
        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.map(|p| p.0).unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .fire_workbook_before_close(options)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn fire_worksheet_change(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    sheet_id: LimitedString<{ crate::ipc_limits::MAX_SHEET_ID_BYTES }>,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    permissions: Option<LimitedMacroPermissions>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    ipc_origin::ensure_main_window(window.label(), "macro execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro execution", ipc_origin::Verb::Is)?;

    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let mut trust_store = trust_shared.lock().unwrap();
            trust_store.ensure_loaded();
            let workbook_id = workbook_id.as_ref().map(|id| id.as_ref());
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }
        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.map(|p| p.0).unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .fire_worksheet_change(
                sheet_id.as_ref(),
                start_row,
                start_col,
                end_row,
                end_col,
                options,
            )
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn fire_selection_change(
    window: tauri::WebviewWindow,
    workbook_id: Option<LimitedString<MAX_IPC_SECURE_STORE_KEY_BYTES>>,
    sheet_id: LimitedString<{ crate::ipc_limits::MAX_SHEET_ID_BYTES }>,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    permissions: Option<LimitedMacroPermissions>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    ipc_origin::ensure_main_window(window.label(), "macro execution", ipc_origin::Verb::Is)?;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_trusted_origin(&url, "macro execution", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "macro execution", ipc_origin::Verb::Is)?;

    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let mut trust_store = trust_shared.lock().unwrap();
            trust_store.ensure_loaded();
            let workbook_id = workbook_id.as_ref().map(|id| id.as_ref());
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }
        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.map(|p| p.0).unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .fire_selection_change(
                sheet_id.as_ref(),
                start_row,
                start_col,
                end_row,
                end_col,
                options,
            )
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn check_for_updates(
    window: tauri::WebviewWindow,
    source: crate::updater::UpdateCheckSource,
) -> Result<(), String> {
    use tauri::Manager as _;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "update checks", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_trusted_origin(&url, "update checks", ipc_origin::Verb::Are)?;
    ipc_origin::ensure_stable_origin(&window, "update checks", ipc_origin::Verb::Are)?;

    let app = window.app_handle();
    crate::updater::spawn_update_check(&app, source);
    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn open_external_url(
    window: tauri::Window,
    url: LimitedString<MAX_IPC_URL_BYTES>,
) -> Result<(), String> {
    use tauri::Manager as _;
    ipc_origin::ensure_main_window(window.label(), "external URL opening", ipc_origin::Verb::Is)?;
    {
        // Prevent arbitrary remote web content from using IPC to open external URLs. This is a
        // defense-in-depth check: even though Tauri's security model should prevent remote origins
        // from accessing the invoke API by default, keep the command itself resilient.
        //
        // Mirrors the trusted-origin checks used by other privileged commands in `main.rs`.
        let Some(webview) = window.app_handle().get_webview_window(window.label()) else {
            return Err("main webview window not available".to_string());
        };
        let webview_url = webview.url().map_err(|err| err.to_string())?;
        ipc_origin::ensure_trusted_origin(
            &webview_url,
            "external URL opening",
            ipc_origin::Verb::Is,
        )?;
        ipc_origin::ensure_stable_origin(&webview, "external URL opening", ipc_origin::Verb::Is)?;
    }

    let parsed = crate::external_url::validate_external_url(url.as_ref())?;

    window
        .shell()
        .open(parsed.as_str(), None)
        .map_err(|e| format!("Failed to open URL: {e}"))?;
    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn quit_app(window: tauri::WebviewWindow) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "app lifecycle", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "app lifecycle", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "app lifecycle", ipc_origin::Verb::Is)?;

    // We intentionally use a hard process exit here. The desktop shell already delegates
    // "should we quit?" decisions (event macros + unsaved prompts) to the frontend.
    // Once the frontend invokes this command, exiting immediately avoids re-entering the
    // CloseRequested handler (which prevents default close to support hide-to-tray).
    std::process::exit(0);
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn restart_app(window: tauri::WebviewWindow) -> Result<(), String> {
    use tauri::Manager as _;
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "app lifecycle", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "app lifecycle", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "app lifecycle", ipc_origin::Verb::Is)?;

    let app = window.app_handle();
    // For update flows we need a graceful shutdown so Tauri and its plugins (notably
    // `tauri-plugin-updater`) can complete any pending work before the process exits.
    //
    // On desktop targets Tauri provides `AppHandle::restart()`. On unsupported targets we fall
    // back to a best-effort graceful exit.
    #[cfg(any(target_os = "windows", target_os = "macos", target_os = "linux"))]
    {
        // `AppHandle::restart()` should terminate the process (by spawning a new instance and
        // exiting), but if it ever returns we fall back to `AppHandle::exit(0)` as a best-effort
        // graceful exit.
        #[allow(unreachable_code, unused_must_use)]
        {
            app.restart();
            app.exit(0);
        }
    }

    #[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
    app.exit(0);

    Ok(())
}

// Clipboard bridge commands.
//
// The frontend prefers the browser Clipboard API when available (so we can copy/paste rich HTML
// tables inside the WebView), but falls back to these commands for formats that are not reliably
// available via WebView clipboard integrations on Linux (Wayland/X11).
//
// These commands are intentionally thin wrappers around `crate::clipboard`, which handles
// platform dispatch and GTK main-thread requirements.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn read_clipboard(
    window: tauri::WebviewWindow,
) -> Result<crate::clipboard::ClipboardContent, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "clipboard access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "clipboard access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "clipboard access", ipc_origin::Verb::Is)?;

    // Clipboard APIs on macOS call into AppKit. AppKit is not thread-safe, and Tauri
    // commands can execute on a background thread, so we always dispatch to the main
    // thread before touching NSPasteboard.
    #[cfg(target_os = "macos")]
    {
        use tauri::Manager as _;
        return window
            .app_handle()
            .run_on_main_thread(crate::clipboard::read)
            .map_err(|e| e.to_string())?
            .map_err(|e| e.to_string());
    }

    #[cfg(not(target_os = "macos"))]
    {
        tauri::async_runtime::spawn_blocking(|| crate::clipboard::read().map_err(|e| e.to_string()))
            .await
            .map_err(|e| e.to_string())?
            .map_err(|e| e.to_string())
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn write_clipboard(
    window: tauri::WebviewWindow,
    text: crate::resource_limits::LimitedString<{ crate::clipboard::MAX_RICH_TEXT_BYTES }>,
    html: Option<crate::resource_limits::LimitedString<{ crate::clipboard::MAX_RICH_TEXT_BYTES }>>,
    rtf: Option<crate::resource_limits::LimitedString<{ crate::clipboard::MAX_RICH_TEXT_BYTES }>>,
    image_png_base64: Option<
        crate::resource_limits::LimitedString<{ crate::clipboard::MAX_IMAGE_PNG_BASE64_BYTES }>,
    >,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "clipboard access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "clipboard access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "clipboard access", ipc_origin::Verb::Is)?;

    let payload = crate::clipboard::ClipboardWritePayload {
        text: Some(text.into_inner()),
        html: html.map(crate::resource_limits::LimitedString::into_inner),
        rtf: rtf.map(crate::resource_limits::LimitedString::into_inner),
        image_png_base64: image_png_base64.map(crate::resource_limits::LimitedString::into_inner),
    };
    #[cfg(target_os = "macos")]
    {
        use tauri::Manager as _;
        return window
            .app_handle()
            .run_on_main_thread(move || crate::clipboard::write(&payload))
            .map_err(|e| e.to_string())?
            .map_err(|e| e.to_string());
    }

    #[cfg(not(target_os = "macos"))]
    {
        tauri::async_runtime::spawn_blocking(move || {
            crate::clipboard::write(&payload).map_err(|e| e.to_string())
        })
        .await
        .map_err(|e| e.to_string())?
    }
}

// -----------------------------------------------------------------------------
// Network + Marketplace proxy (desktop webview)
//
// Desktop extensions and the in-webview marketplace client run inside the Tauri WebView, which is
// governed by the app CSP (including `connect-src`). The production CSP is intentionally
// restrictive (HTTPS + WebSockets via `https:`/`ws:`/`wss:`; no `http:`).
//
// Network access for extensions is primarily enforced by Formula's permission model + extension
// worker guardrails (which replace `fetch`/`WebSocket` inside the worker); CSP is defense-in-depth.
//
// To avoid relying on permissive CORS headers for the `tauri://…` origin (and to keep networking
// behavior consistent across WebViews), the WebView prefers routing outbound HTTP(S) through these
// Tauri commands so the Rust backend performs the network request.

#[cfg(any(feature = "desktop", test))]
pub(crate) fn is_local_http_allowed(url: &Url) -> bool {
    if url.scheme() != "http" {
        return false;
    }

    match url.host() {
        Some(url::Host::Domain(domain)) => {
            let domain = domain.trim_end_matches('.').to_ascii_lowercase();
            domain == "localhost" || domain.ends_with(".localhost")
        }
        Some(url::Host::Ipv4(ip)) => ip.is_loopback(),
        Some(url::Host::Ipv6(ip)) => ip.is_loopback(),
        None => false,
    }
}

#[cfg(any(feature = "desktop", test))]
pub(crate) fn ensure_ipc_network_url_allowed(
    url: &Url,
    context: &str,
    debug_assertions: bool,
) -> Result<(), String> {
    match url.scheme() {
        "https" => Ok(()),
        "http" => {
            if debug_assertions || is_local_http_allowed(url) {
                Ok(())
            } else {
                let host = url.host_str().unwrap_or("<unknown>");
                Err(format!(
                    "{context}: http URLs are only allowed for localhost in release builds (got host '{host}'). Use https:// for remote hosts, or http://localhost for local development."
                ))
            }
        }
        other => Err(format!(
            "Unsupported url scheme for {context}: {other} (only http/https allowed)"
        )),
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn pyodide_index_url(
    window: tauri::WebviewWindow,
    download: Option<bool>,
) -> Result<Option<String>, String> {
    use tauri::Manager as _;
    ipc_origin::ensure_main_window_and_stable_origin(
        &window,
        "python runtime download",
        ipc_origin::Verb::Is,
    )?;

    let download = download.unwrap_or(false);

    // If Pyodide assets were explicitly bundled into `dist/` (opt-in via
    // `FORMULA_BUNDLE_PYODIDE_ASSETS=1`), prefer loading them from the embedded frontend assets
    // instead of downloading into the app-data cache.
    //
    // This keeps offline bundled builds functional without needing a first-run download, while
    // default builds still use the on-demand cache to keep installers small.
    let version_tag = crate::pyodide_assets::pyodide_version_tag();
    let bundled_probe = format!("pyodide/{version_tag}/full/pyodide.js");
    if window
        .app_handle()
        .asset_resolver()
        .get(bundled_probe)
        .is_some()
    {
        return Ok(Some(format!("/pyodide/{version_tag}/full/")));
    }

    crate::pyodide_assets::pyodide_index_url_from_cache(download, |progress| {
        let _ = window.emit(
            crate::pyodide_assets::PYODIDE_DOWNLOAD_PROGRESS_EVENT,
            progress,
        );
    })
    .await
    .map_err(|err| {
        format!(
            "Failed to prepare the Python runtime.\n\n\
Pyodide is downloaded on-demand and cached for future runs.\n\
Error: {err}"
        )
    })
}

#[cfg(feature = "desktop")]
pub use crate::network_fetch::NetworkFetchResult;

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn network_fetch(
    window: tauri::WebviewWindow,
    url: LimitedString<MAX_IPC_URL_BYTES>,
    init: Option<JsonValue>,
) -> Result<NetworkFetchResult, String> {
    let origin_url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "network access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&origin_url, "network access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "network access", ipc_origin::Verb::Is)?;

    let parsed_url = reqwest::Url::parse(url.as_ref()).map_err(|e| format!("Invalid url: {e}"))?;
    ensure_ipc_network_url_allowed(&parsed_url, "network_fetch", cfg!(debug_assertions))?;

    let url = url.into_inner();
    let init = init.unwrap_or(JsonValue::Null);
    crate::network_fetch::network_fetch_impl(url.as_ref(), &init).await
}

#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct MarketplaceSearchArgs {
    pub base_url: LimitedString<MAX_IPC_URL_BYTES>,
    pub q: Option<LimitedString<MAX_IPC_MARKETPLACE_QUERY_BYTES>>,
    pub category: Option<LimitedString<MAX_IPC_MARKETPLACE_QUERY_BYTES>>,
    pub tag: Option<LimitedString<MAX_IPC_MARKETPLACE_QUERY_BYTES>>,
    pub verified: Option<bool>,
    pub featured: Option<bool>,
    pub sort: Option<LimitedString<MAX_IPC_MARKETPLACE_QUERY_BYTES>>,
    pub limit: Option<u32>,
    pub offset: Option<u32>,
    pub cursor: Option<LimitedString<MAX_IPC_MARKETPLACE_QUERY_BYTES>>,
}

#[cfg(feature = "desktop")]
fn parse_marketplace_base_url(base_url: &str) -> Result<reqwest::Url, String> {
    let url = reqwest::Url::parse(base_url).map_err(|_| {
        "Marketplace baseUrl must be an absolute http(s) URL when running under Tauri".to_string()
    })?;
    ensure_ipc_network_url_allowed(&url, "Marketplace baseUrl", cfg!(debug_assertions))?;
    Ok(url)
}

#[cfg(any(feature = "desktop", test))]
fn marketplace_reqwest_client() -> Result<reqwest::Client, String> {
    let debug_assertions = cfg!(debug_assertions);
    reqwest::Client::builder()
        .redirect(reqwest::redirect::Policy::custom(move |attempt| {
            if attempt.previous().len() >= 10 {
                return attempt.stop();
            }
            match ensure_ipc_network_url_allowed(
                attempt.url(),
                "Marketplace redirect",
                debug_assertions,
            ) {
                Ok(()) => attempt.follow(),
                Err(_) => attempt.stop(),
            }
        }))
        .build()
        .map_err(|e| e.to_string())
}

#[cfg(any(feature = "desktop", test))]
async fn marketplace_get_response(url: reqwest::Url) -> Result<reqwest::Response, String> {
    // The marketplace base URL is treated as untrusted input coming from the WebView. Apply the
    // same scheme allowlist as `network_fetch`, then enforce it again for redirects.
    ensure_ipc_network_url_allowed(&url, "Marketplace request", cfg!(debug_assertions))?;

    let client = marketplace_reqwest_client()?;
    let response = client.get(url).send().await.map_err(|e| e.to_string())?;

    // If the redirect policy stopped following redirects (e.g. because a hop violates the IPC URL
    // policy), make that explicit rather than returning the raw 3xx response.
    if response.status().is_redirection() {
        if let Some(location) = response
            .headers()
            .get(reqwest::header::LOCATION)
            .and_then(|v| v.to_str().ok())
        {
            if let Ok(target) = response.url().join(location) {
                ensure_ipc_network_url_allowed(
                    &target,
                    "Marketplace redirect",
                    cfg!(debug_assertions),
                )?;
            }
        }
    }

    Ok(response)
}

#[cfg(any(feature = "desktop", test))]
async fn marketplace_fetch_json_with_limit(
    url: reqwest::Url,
    limit_bytes: usize,
    limit_context: &'static str,
    status_err_prefix: &'static str,
) -> Result<JsonValue, String> {
    let mut response = marketplace_get_response(url).await?;
    if !response.status().is_success() {
        return Err(format!("{status_err_prefix} ({})", response.status()));
    }

    let bytes = crate::network_limits::read_response_body_with_limit(
        &mut response,
        limit_bytes,
        limit_context,
    )
    .await?;
    serde_json::from_slice::<JsonValue>(&bytes).map_err(|e| e.to_string())
}

#[cfg(any(feature = "desktop", test))]
async fn marketplace_fetch_optional_json_with_limit(
    url: reqwest::Url,
    limit_bytes: usize,
    limit_context: &'static str,
    status_err_prefix: &'static str,
) -> Result<Option<JsonValue>, String> {
    let mut response = marketplace_get_response(url).await?;

    if response.status().as_u16() == 404 {
        return Ok(None);
    }
    if !response.status().is_success() {
        return Err(format!("{status_err_prefix} ({})", response.status()));
    }

    let bytes = crate::network_limits::read_response_body_with_limit(
        &mut response,
        limit_bytes,
        limit_context,
    )
    .await?;
    let json = serde_json::from_slice::<JsonValue>(&bytes).map_err(|e| e.to_string())?;
    Ok(Some(json))
}

#[cfg(any(feature = "desktop", test))]
async fn marketplace_fetch_optional_download_payload(
    url: reqwest::Url,
    status_err_prefix: &'static str,
) -> Result<Option<MarketplaceDownloadPayload>, String> {
    let response = marketplace_get_response(url).await?;
    if response.status().as_u16() == 404 {
        return Ok(None);
    }
    if !response.status().is_success() {
        return Err(format!("{status_err_prefix} ({})", response.status()));
    }

    marketplace_download_payload_from_response(response)
        .await
        .map(Some)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn marketplace_search(
    window: tauri::WebviewWindow,
    args: MarketplaceSearchArgs,
) -> Result<JsonValue, String> {
    let origin_url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "marketplace access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&origin_url, "marketplace access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "marketplace access", ipc_origin::Verb::Is)?;

    let mut url = parse_marketplace_base_url(args.base_url.as_ref())?;
    {
        let mut segments = url
            .path_segments_mut()
            .map_err(|_| "Invalid marketplace baseUrl".to_string())?;
        segments.push("search");
    }

    {
        let mut qp = url.query_pairs_mut();
        if let Some(v) = args
            .q
            .as_ref()
            .map(|s| s.as_ref())
            .filter(|s| !s.trim().is_empty())
        {
            qp.append_pair("q", v);
        }
        if let Some(v) = args
            .category
            .as_ref()
            .map(|s| s.as_ref())
            .filter(|s| !s.trim().is_empty())
        {
            qp.append_pair("category", v);
        }
        if let Some(v) = args
            .tag
            .as_ref()
            .map(|s| s.as_ref())
            .filter(|s| !s.trim().is_empty())
        {
            qp.append_pair("tag", v);
        }
        if let Some(v) = args.verified {
            qp.append_pair("verified", if v { "true" } else { "false" });
        }
        if let Some(v) = args.featured {
            qp.append_pair("featured", if v { "true" } else { "false" });
        }
        if let Some(v) = args
            .sort
            .as_ref()
            .map(|s| s.as_ref())
            .filter(|s| !s.trim().is_empty())
        {
            qp.append_pair("sort", v);
        }
        if let Some(v) = args.limit {
            qp.append_pair("limit", &v.to_string());
        }
        if let Some(v) = args.offset {
            qp.append_pair("offset", &v.to_string());
        }
        if let Some(v) = args
            .cursor
            .as_ref()
            .map(|s| s.as_ref())
            .filter(|s| !s.trim().is_empty())
        {
            qp.append_pair("cursor", v);
        }
    }

    marketplace_fetch_json_with_limit(
        url,
        crate::network_limits::MARKETPLACE_JSON_MAX_BODY_BYTES,
        "marketplace_search",
        "Marketplace search failed",
    )
    .await
}

#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct MarketplaceGetExtensionArgs {
    pub base_url: LimitedString<MAX_IPC_URL_BYTES>,
    pub id: LimitedString<MAX_IPC_MARKETPLACE_ID_BYTES>,
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn marketplace_get_extension(
    window: tauri::WebviewWindow,
    args: MarketplaceGetExtensionArgs,
) -> Result<Option<JsonValue>, String> {
    let origin_url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "marketplace access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&origin_url, "marketplace access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "marketplace access", ipc_origin::Verb::Is)?;

    let mut url = parse_marketplace_base_url(args.base_url.as_ref())?;
    {
        let mut segments = url
            .path_segments_mut()
            .map_err(|_| "Invalid marketplace baseUrl".to_string())?;
        segments.push("extensions");
        segments.push(args.id.as_ref().trim());
    }

    marketplace_fetch_optional_json_with_limit(
        url,
        crate::network_limits::MARKETPLACE_JSON_MAX_BODY_BYTES,
        "marketplace_get_extension",
        "Marketplace getExtension failed",
    )
    .await
}

#[cfg(feature = "desktop")]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct MarketplaceDownloadArgs {
    pub base_url: LimitedString<MAX_IPC_URL_BYTES>,
    pub id: LimitedString<MAX_IPC_MARKETPLACE_ID_BYTES>,
    pub version: LimitedString<MAX_IPC_MARKETPLACE_VERSION_BYTES>,
}

#[cfg(any(feature = "desktop", test))]
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct MarketplaceDownloadPayload {
    pub bytes_base64: String,
    pub signature_base64: Option<String>,
    pub sha256: Option<String>,
    pub format_version: Option<u32>,
    pub publisher: Option<String>,
    pub publisher_key_id: Option<String>,
    pub scan_status: Option<String>,
    pub files_sha256: Option<String>,
}

#[cfg(any(feature = "desktop", test))]
fn marketplace_bounded_header_string(
    headers: &reqwest::header::HeaderMap,
    header_name: &'static str,
) -> Result<Option<String>, String> {
    use crate::resource_limits::MAX_MARKETPLACE_HEADER_BYTES;

    let Some(value) = headers.get(header_name) else {
        return Ok(None);
    };

    let byte_len = value.as_bytes().len();
    if byte_len > MAX_MARKETPLACE_HEADER_BYTES {
        return Err(format!(
            "Marketplace header '{header_name}' exceeded MAX_MARKETPLACE_HEADER_BYTES ({byte_len} > {MAX_MARKETPLACE_HEADER_BYTES} bytes)"
        ));
    }

    Ok(value.to_str().ok().map(|s| s.to_string()))
}

#[cfg(any(feature = "desktop", test))]
async fn marketplace_download_payload_from_response(
    mut response: reqwest::Response,
) -> Result<MarketplaceDownloadPayload, String> {
    use crate::resource_limits::MAX_MARKETPLACE_PACKAGE_BYTES;
    use base64::{engine::general_purpose::STANDARD, Engine as _};

    let (
        signature_base64,
        sha256,
        format_version,
        publisher,
        publisher_key_id,
        scan_status,
        files_sha256,
    ) = {
        let headers = response.headers();

        let signature_base64 = marketplace_bounded_header_string(headers, "x-package-signature")?;
        let sha256 = marketplace_bounded_header_string(headers, "x-package-sha256")?;
        let format_version =
            marketplace_bounded_header_string(headers, "x-package-format-version")?
                .and_then(|s| s.parse::<u32>().ok());
        let publisher = marketplace_bounded_header_string(headers, "x-publisher")?;
        let publisher_key_id = marketplace_bounded_header_string(headers, "x-publisher-key-id")?;
        let scan_status = marketplace_bounded_header_string(headers, "x-package-scan-status")?;
        let files_sha256 = marketplace_bounded_header_string(headers, "x-package-files-sha256")?;

        (
            signature_base64,
            sha256,
            format_version,
            publisher,
            publisher_key_id,
            scan_status,
            files_sha256,
        )
    };

    let bytes = crate::network_limits::read_response_body_with_limit(
        &mut response,
        MAX_MARKETPLACE_PACKAGE_BYTES,
        "marketplace_download_package",
    )
    .await?;

    let bytes_base64 = STANDARD.encode(&bytes);

    Ok(MarketplaceDownloadPayload {
        bytes_base64,
        signature_base64,
        sha256,
        format_version,
        publisher,
        publisher_key_id,
        scan_status,
        files_sha256,
    })
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn marketplace_download_package(
    window: tauri::WebviewWindow,
    args: MarketplaceDownloadArgs,
) -> Result<Option<MarketplaceDownloadPayload>, String> {
    let origin_url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "marketplace access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&origin_url, "marketplace access", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "marketplace access", ipc_origin::Verb::Is)?;

    let mut url = parse_marketplace_base_url(args.base_url.as_ref())?;
    {
        let mut segments = url
            .path_segments_mut()
            .map_err(|_| "Invalid marketplace baseUrl".to_string())?;
        segments.push("extensions");
        segments.push(args.id.as_ref().trim());
        segments.push("download");
        segments.push(args.version.as_ref().trim());
    }

    marketplace_fetch_optional_download_payload(url, "Marketplace download failed").await
}
#[cfg(test)]
mod tests {
    use super::*;
    use crate::file_io::read_xlsx_blocking;
    use crate::resource_limits::{MAX_MARKETPLACE_HEADER_BYTES, MAX_MARKETPLACE_PACKAGE_BYTES};
    use base64::{engine::general_purpose::STANDARD, Engine as _};
    use formula_xlsx::drawingml::{
        PreservedDrawingParts, PreservedSheetPicture, SheetRelationshipStub,
    };
    use std::io::Write;
    use std::path::Path;
    use tempfile::TempDir;
    use tokio::io::{AsyncReadExt, AsyncWriteExt};
    use tokio::net::TcpListener;

    #[test]
    fn ipc_pivot_field_ref_deserializes_and_converts_data_model_refs() {
        let field: IpcPivotFieldRef = serde_json::from_value(serde_json::json!("Region")).unwrap();
        assert_eq!(
            PivotFieldRef::from(field),
            PivotFieldRef::CacheFieldName("Region".to_string())
        );

        let field: IpcPivotFieldRef =
            serde_json::from_value(serde_json::json!("Sales[Amount]")).unwrap();
        assert_eq!(
            PivotFieldRef::from(field),
            PivotFieldRef::DataModelColumn {
                table: "Sales".to_string(),
                column: "Amount".to_string(),
            }
        );

        let field: IpcPivotFieldRef =
            serde_json::from_value(serde_json::json!("[Total Sales]")).unwrap();
        assert_eq!(
            PivotFieldRef::from(field),
            PivotFieldRef::DataModelMeasure("Total Sales".to_string())
        );

        let field: IpcPivotFieldRef =
            serde_json::from_value(serde_json::json!({"table":"Sales","column":"Amount"})).unwrap();
        assert_eq!(
            PivotFieldRef::from(field),
            PivotFieldRef::DataModelColumn {
                table: "Sales".to_string(),
                column: "Amount".to_string(),
            }
        );

        let field: IpcPivotFieldRef =
            serde_json::from_value(serde_json::json!({"measure":"Total Sales"})).unwrap();
        assert_eq!(
            PivotFieldRef::from(field),
            PivotFieldRef::DataModelMeasure("Total Sales".to_string())
        );

        let field: IpcPivotFieldRef =
            serde_json::from_value(serde_json::json!({"name":"Total Sales"})).unwrap();
        assert_eq!(
            PivotFieldRef::from(field),
            PivotFieldRef::DataModelMeasure("Total Sales".to_string())
        );
    }

    #[test]
    fn imported_sheet_background_images_falls_back_to_preserved_parts_when_origin_bytes_missing() {
        use std::collections::BTreeMap;

        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.sheets[0].xlsx_worksheet_part = Some("xl/worksheets/sheet1.xml".to_string());

        workbook.preserved_drawing_parts = Some(PreservedDrawingParts {
            content_types_xml: b"<Types/>".to_vec(),
            parts: BTreeMap::from([("xl/media/bg.png".to_string(), vec![0x01, 0x02, 0x03])]),
            sheet_drawings: BTreeMap::new(),
            sheet_pictures: BTreeMap::from([(
                "Sheet1".to_string(),
                PreservedSheetPicture {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    picture_xml: br#"<picture r:id="rId1"/>"#.to_vec(),
                    picture_rel: SheetRelationshipStub {
                        rel_id: "rId1".to_string(),
                        target: "../media/bg.png".to_string(),
                    },
                },
            )]),
            sheet_ole_objects: BTreeMap::new(),
            sheet_controls: BTreeMap::new(),
            sheet_drawing_hfs: BTreeMap::new(),
            chart_sheets: BTreeMap::new(),
        });

        // Ensure we exercise the preserved-parts path (no origin bytes).
        workbook.origin_xlsx_bytes = None;

        let payloads = sheet_background_image_payloads_from_preserved_parts(&workbook);
        assert_eq!(payloads.len(), 1);

        let infos = imported_sheet_background_images_from_preserved_payloads(payloads);
        assert_eq!(infos.len(), 1);
        assert_eq!(infos[0].sheet_name, "Sheet1");
        assert_eq!(infos[0].worksheet_part, "xl/worksheets/sheet1.xml");
        assert_eq!(infos[0].image_id, "bg.png");
        assert_eq!(infos[0].bytes_base64, "AQID");
        assert_eq!(infos[0].mime_type, "image/png");
    }

    #[tokio::test]
    async fn marketplace_download_payload_rejects_oversized_package_bytes() {
        let listener = TcpListener::bind("127.0.0.1:0")
            .await
            .expect("bind test server");
        let addr = listener.local_addr().expect("listener addr");

        let server = tokio::spawn(async move {
            let (mut stream, _) = listener.accept().await.expect("accept");
            let mut buf = [0u8; 1024];
            let _ = stream.read(&mut buf).await;

            let headers = "HTTP/1.1 200 OK\r\nConnection: close\r\n\r\n";
            stream
                .write_all(headers.as_bytes())
                .await
                .expect("write headers");

            let mut remaining = MAX_MARKETPLACE_PACKAGE_BYTES + 1;
            let chunk = vec![b'a'; 16 * 1024];
            while remaining > 0 {
                let n = remaining.min(chunk.len());
                if stream.write_all(&chunk[..n]).await.is_err() {
                    break;
                }
                remaining -= n;
            }

            let _ = stream.shutdown().await;
        });

        let url = reqwest::Url::parse(&format!("http://{addr}/download")).expect("parse url");
        let err = marketplace_fetch_optional_download_payload(url, "Marketplace download failed")
            .await
            .expect_err("expected package byte limit error");
        assert!(
            err.contains("Response body too large for marketplace_download_package"),
            "unexpected error: {err}"
        );
        assert!(
            err.contains(&MAX_MARKETPLACE_PACKAGE_BYTES.to_string()),
            "unexpected error: {err}"
        );
        assert!(
            err.contains(&(MAX_MARKETPLACE_PACKAGE_BYTES + 1).to_string()),
            "unexpected error: {err}"
        );

        server.await.expect("server task");
    }

    #[tokio::test]
    async fn marketplace_download_payload_rejects_oversized_signature_header() {
        let listener = TcpListener::bind("127.0.0.1:0")
            .await
            .expect("bind test server");
        let addr = listener.local_addr().expect("listener addr");

        let server = tokio::spawn(async move {
            let (mut stream, _) = listener.accept().await.expect("accept");
            let mut buf = [0u8; 1024];
            let _ = stream.read(&mut buf).await;

            let signature = "a".repeat(MAX_MARKETPLACE_HEADER_BYTES + 1);
            let response = format!(
                "HTTP/1.1 200 OK\r\nx-package-signature: {signature}\r\nContent-Length: 0\r\nConnection: close\r\n\r\n"
            );
            stream
                .write_all(response.as_bytes())
                .await
                .expect("write response");
            let _ = stream.shutdown().await;
        });

        let url = reqwest::Url::parse(&format!("http://{addr}/download")).expect("parse url");
        let err = marketplace_fetch_optional_download_payload(url, "Marketplace download failed")
            .await
            .expect_err("expected header byte limit error");
        assert!(
            err.contains("MAX_MARKETPLACE_HEADER_BYTES"),
            "unexpected error: {err}"
        );
        assert!(
            err.contains("x-package-signature"),
            "unexpected error: {err}"
        );

        server.await.expect("server task");
    }

    #[tokio::test]
    async fn marketplace_download_payload_roundtrips_small_response_bytes() {
        let listener = TcpListener::bind("127.0.0.1:0")
            .await
            .expect("bind test server");
        let addr = listener.local_addr().expect("listener addr");

        let body: Vec<u8> = b"hello marketplace".to_vec();
        let expected = body.clone();
        let body_len = body.len();

        let server = tokio::spawn(async move {
            let (mut stream, _) = listener.accept().await.expect("accept");
            let mut buf = [0u8; 1024];
            let _ = stream.read(&mut buf).await;

            let headers = format!(
                "HTTP/1.1 200 OK\r\nx-package-signature: sig\r\nContent-Length: {body_len}\r\nConnection: close\r\n\r\n"
            );
            stream
                .write_all(headers.as_bytes())
                .await
                .expect("write headers");
            stream.write_all(&body).await.expect("write body");
            let _ = stream.shutdown().await;
        });

        let url = reqwest::Url::parse(&format!("http://{addr}/download")).expect("parse url");
        let payload =
            marketplace_fetch_optional_download_payload(url, "Marketplace download failed")
                .await
                .expect("expected download payload")
                .expect("expected payload");
        assert_eq!(payload.signature_base64, Some("sig".to_string()));

        let decoded = STANDARD
            .decode(payload.bytes_base64.as_bytes())
            .expect("base64 decode payload");
        assert_eq!(decoded, expected);

        server.await.expect("server task");
    }

    #[tokio::test]
    async fn marketplace_json_limit_allows_small_json() {
        let listener = TcpListener::bind("127.0.0.1:0")
            .await
            .expect("bind test server");
        let addr = listener.local_addr().expect("listener addr");

        let body = br#"{"ok":true}"#.to_vec();
        let body_len = body.len();

        let server = tokio::spawn(async move {
            let (mut stream, _) = listener.accept().await.expect("accept");
            let mut buf = [0u8; 1024];
            let _ = stream.read(&mut buf).await;

            let headers = format!(
                "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: {body_len}\r\nConnection: close\r\n\r\n"
            );
            stream
                .write_all(headers.as_bytes())
                .await
                .expect("write headers");
            stream.write_all(&body).await.expect("write body");
            let _ = stream.shutdown().await;
        });

        let url = reqwest::Url::parse(&format!("http://{addr}/search")).expect("parse url");
        let parsed = marketplace_fetch_json_with_limit(
            url,
            1024,
            "marketplace_json_test",
            "Marketplace json test failed",
        )
        .await
        .expect("expected JSON body to be within limit");
        assert_eq!(parsed, serde_json::json!({ "ok": true }));

        server.await.expect("server task");
    }

    #[tokio::test]
    async fn marketplace_optional_json_returns_none_on_404() {
        let listener = TcpListener::bind("127.0.0.1:0")
            .await
            .expect("bind test server");
        let addr = listener.local_addr().expect("listener addr");

        let server = tokio::spawn(async move {
            let (mut stream, _) = listener.accept().await.expect("accept");
            let mut buf = [0u8; 1024];
            let _ = stream.read(&mut buf).await;

            let headers = "HTTP/1.1 404 Not Found\r\nConnection: close\r\n\r\n";
            stream
                .write_all(headers.as_bytes())
                .await
                .expect("write headers");
            let _ = stream.shutdown().await;
        });

        let url =
            reqwest::Url::parse(&format!("http://{addr}/extensions/missing")).expect("parse url");
        let parsed = marketplace_fetch_optional_json_with_limit(
            url,
            1024,
            "marketplace_optional_json_test",
            "Marketplace optional json test failed",
        )
        .await
        .expect("expected request to succeed");
        assert!(parsed.is_none(), "expected 404 to map to None");

        server.await.expect("server task");
    }

    #[tokio::test]
    async fn marketplace_json_limit_rejects_oversized_json_without_content_length() {
        // Omit Content-Length to ensure the limit is enforced while streaming.
        let listener = TcpListener::bind("127.0.0.1:0")
            .await
            .expect("bind test server");
        let addr = listener.local_addr().expect("listener addr");

        let limit_bytes = 64usize;
        let payload = format!(r#"{{"data":"{}"}}"#, "a".repeat(128)).into_bytes();

        let server = tokio::spawn(async move {
            let (mut stream, _) = listener.accept().await.expect("accept");
            let mut buf = [0u8; 1024];
            let _ = stream.read(&mut buf).await;

            let headers =
                "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nConnection: close\r\n\r\n";
            stream
                .write_all(headers.as_bytes())
                .await
                .expect("write headers");
            // Write in chunks to exercise streaming limit enforcement.
            for chunk in payload.chunks(16) {
                if stream.write_all(chunk).await.is_err() {
                    break;
                }
            }
            let _ = stream.shutdown().await;
        });

        let url = reqwest::Url::parse(&format!("http://{addr}/search")).expect("parse url");
        let err = marketplace_fetch_json_with_limit(
            url,
            limit_bytes,
            "marketplace_json_test",
            "Marketplace json test failed",
        )
        .await
        .expect_err("expected JSON body size to be rejected");
        assert!(
            err.contains("Response body too large") && err.contains(&limit_bytes.to_string()),
            "unexpected error: {err}"
        );

        server.await.expect("server task");
    }

    #[test]
    fn ipc_network_scheme_policy_allows_https_everywhere() {
        let url = Url::parse("https://example.com/").expect("parse url");
        ensure_ipc_network_url_allowed(&url, "test", false).expect("https should be allowed");
    }

    #[test]
    fn ipc_network_scheme_policy_allows_http_localhost_in_release_mode() {
        let url = Url::parse("http://localhost:3000/").expect("parse url");
        ensure_ipc_network_url_allowed(&url, "test", false)
            .expect("http localhost should be allowed in release mode");
    }

    #[test]
    fn ipc_network_scheme_policy_denies_http_remote_in_release_mode() {
        let url = Url::parse("http://example.com/").expect("parse url");
        let err = ensure_ipc_network_url_allowed(&url, "test", false)
            .expect_err("http remote should be denied in release mode");
        assert!(
            err.contains("http URLs are only allowed for localhost in release builds"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn ipc_network_scheme_policy_allows_http_remote_in_debug_mode() {
        let url = Url::parse("http://example.com/").expect("parse url");
        ensure_ipc_network_url_allowed(&url, "test", true)
            .expect("http remote should be allowed in debug mode");
    }

    #[test]
    fn local_http_allowlist_matches_expected_hosts() {
        for candidate in [
            "http://localhost/",
            "http://foo.localhost/",
            "http://127.0.0.1/",
            "http://[::1]/",
        ] {
            let url = Url::parse(candidate).expect("parse url");
            assert!(
                is_local_http_allowed(&url),
                "expected {candidate} to be allowed"
            );
        }

        let url = Url::parse("http://example.com/").expect("parse url");
        assert!(
            !is_local_http_allowed(&url),
            "expected remote http host to be rejected"
        );
    }

    #[test]
    fn normalize_tab_color_rgb_normalizes_hex_strings() {
        assert_eq!(
            normalize_tab_color_rgb("#00ff00").expect("should normalize rgb"),
            "FF00FF00"
        );
        assert_eq!(
            normalize_tab_color_rgb("00ff00").expect("should normalize rgb"),
            "FF00FF00"
        );
        assert_eq!(
            normalize_tab_color_rgb("80ff0000").expect("should preserve alpha"),
            "80FF0000"
        );
        assert_eq!(
            normalize_tab_color_rgb("#80ff0000").expect("should preserve alpha"),
            "80FF0000"
        );
    }

    #[test]
    fn normalize_tab_color_rgb_rejects_invalid_values() {
        assert!(
            normalize_tab_color_rgb("").is_err(),
            "expected empty rgb to fail"
        );
        assert!(
            normalize_tab_color_rgb("   ").is_err(),
            "expected whitespace rgb to fail"
        );

        let err = normalize_tab_color_rgb("#xyz").expect_err("expected invalid hex to fail");
        assert!(
            err.contains("6-digit") || err.contains("8-digit"),
            "unexpected error: {err}"
        );
        let err = normalize_tab_color_rgb("#12345").expect_err("expected wrong length to fail");
        assert!(
            err.contains("6-digit") || err.contains("8-digit"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn sheet_visibility_serializes_as_expected() {
        let json = serde_json::to_value(SheetVisibility::VeryHidden).expect("serialize visibility");
        assert_eq!(json, serde_json::Value::String("veryHidden".to_string()));

        let parsed: SheetVisibility =
            serde_json::from_value(serde_json::Value::String("veryHidden".to_string()))
                .expect("deserialize visibility");
        assert_eq!(parsed, SheetVisibility::VeryHidden);
    }

    #[test]
    fn limited_range_cell_edits_deserializes_small_payloads() {
        let value = serde_json::json!([
            [
                { "value": 1 },
                { "value": 2, "formula": "=1+1" }
            ],
            [
                { "value": 3 },
                { "value": 4 }
            ]
        ]);

        let parsed: LimitedRangeCellEdits =
            serde_json::from_value(value).expect("expected limited matrix to deserialize");
        assert_eq!(parsed.0.len(), 2);
        assert_eq!(parsed.0[0].len(), 2);
    }

    #[test]
    fn cell_edit_ipc_rejects_object_and_array_values() {
        let err = serde_json::from_str::<RangeCellEdit>(r#"{"value":{"a":1}}"#)
            .expect_err("expected object value to be rejected");
        assert!(
            err.to_string().contains("scalar") && err.to_string().contains("object"),
            "unexpected error: {err}"
        );

        let err = serde_json::from_str::<RangeCellEdit>(r#"{"value":[1,2,3]}"#)
            .expect_err("expected array value to be rejected");
        assert!(
            err.to_string().contains("scalar") && err.to_string().contains("array"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn cell_edit_ipc_rejects_oversized_cell_value_strings() {
        let max = crate::resource_limits::MAX_CELL_VALUE_STRING_BYTES;
        let oversized = "x".repeat(max + 1);
        let json = format!(r#"{{"value":"{oversized}"}}"#);
        let err =
            serde_json::from_str::<RangeCellEdit>(&json).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains(&max.to_string()),
            "expected error message to mention limit: {err}"
        );
    }

    #[test]
    fn cell_edit_ipc_value_string_limit_counts_utf8_bytes() {
        let max = crate::resource_limits::MAX_CELL_VALUE_STRING_BYTES;
        let emoji = "💩";
        let bytes_per_emoji = emoji.len();
        assert!(bytes_per_emoji > 1, "expected multi-byte UTF-8 string");

        let fits = emoji.repeat(max / bytes_per_emoji);
        let json = format!(r#"{{"value":"{fits}"}}"#);
        serde_json::from_str::<RangeCellEdit>(&json).expect("expected value at limit to parse");

        let oversized = emoji.repeat(max / bytes_per_emoji + 1);
        let json = format!(r#"{{"value":"{oversized}"}}"#);
        let err =
            serde_json::from_str::<RangeCellEdit>(&json).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains("cell value string")
                && err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn cell_edit_ipc_rejects_unknown_range_cell_edit_fields() {
        let err = serde_json::from_str::<RangeCellEdit>(r#"{"value":1,"extra":[1,2,3]}"#)
            .expect_err("expected unknown field to be rejected");
        assert!(
            err.to_string().contains("unknown field") && err.to_string().contains("extra"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn cell_edit_ipc_rejects_oversized_formula_strings() {
        let max = crate::resource_limits::MAX_CELL_FORMULA_BYTES;
        let oversized = "x".repeat(max + 1);
        let json = format!(r#"{{"formula":"{oversized}"}}"#);
        let err =
            serde_json::from_str::<RangeCellEdit>(&json).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains("cell formula") && err.to_string().contains(&max.to_string()),
            "expected error message to mention limit: {err}"
        );
    }

    #[test]
    fn cell_edit_ipc_formula_limit_counts_utf8_bytes() {
        let max = crate::resource_limits::MAX_CELL_FORMULA_BYTES;
        let emoji = "💩";
        let bytes_per_emoji = emoji.len();
        assert!(bytes_per_emoji > 1, "expected multi-byte UTF-8 string");

        // This payload is exactly at the byte limit, but well under the limit in Unicode scalar
        // count. This ensures we are enforcing UTF-8 byte length, not character count.
        let fits = emoji.repeat(max / bytes_per_emoji);
        let json = format!(r#"{{"formula":"{fits}"}}"#);
        serde_json::from_str::<RangeCellEdit>(&json).expect("expected formula at limit to parse");

        let oversized = emoji.repeat(max / bytes_per_emoji + 1);
        let json = format!(r#"{{"formula":"{oversized}"}}"#);
        let err = serde_json::from_str::<RangeCellEdit>(&json)
            .expect_err("expected formula over byte limit to be rejected");
        assert!(
            err.to_string().contains("cell formula") && err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn cell_edit_ipc_rejects_oversized_range_payload_cell_count() {
        type SmallLimitedRangeCellEdits = LimitedMatrix<RangeCellEdit, 100, 10>;

        let max = 10usize;
        let mut json = String::from("[[");
        for idx in 0..=max {
            if idx > 0 {
                json.push(',');
            }
            json.push_str("{}");
        }
        json.push_str("]]");

        let err = serde_json::from_str::<SmallLimitedRangeCellEdits>(&json)
            .expect_err("expected cell count limit to be rejected");
        assert!(
            err.to_string().contains("cells") && err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn limited_cell_value_deserializes_scalar_values() {
        let parsed: LimitedCellValue =
            serde_json::from_str("null").expect("expected null to deserialize");
        assert_eq!(parsed, LimitedCellValue::Null);

        let parsed: LimitedCellValue =
            serde_json::from_str("true").expect("expected bool to deserialize");
        assert_eq!(parsed, LimitedCellValue::Bool(true));

        let parsed: LimitedCellValue =
            serde_json::from_str("123").expect("expected number to deserialize");
        match parsed {
            LimitedCellValue::Number(n) => assert_eq!(n, 123.0),
            other => panic!("unexpected value: {other:?}"),
        }

        let parsed: LimitedCellValue =
            serde_json::from_str(r#""hello""#).expect("expected string to deserialize");
        match parsed {
            LimitedCellValue::String(s) => assert_eq!(s.as_ref(), "hello"),
            other => panic!("unexpected value: {other:?}"),
        }
    }

    #[test]
    fn limited_range_cell_edits_rejects_too_many_rows() {
        let max = crate::resource_limits::MAX_RANGE_DIM;
        let mut json = String::from("[");
        for idx in 0..=max {
            if idx > 0 {
                json.push(',');
            }
            json.push_str("[]");
        }
        json.push(']');

        let err = serde_json::from_str::<LimitedRangeCellEdits>(&json)
            .expect_err("expected oversized row count to be rejected");
        assert!(
            err.to_string().contains("max") && err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn limited_range_cell_edits_rejects_too_many_cols() {
        let max = crate::resource_limits::MAX_RANGE_DIM;
        let mut json = String::from("[[");
        for idx in 0..=max {
            if idx > 0 {
                json.push(',');
            }
            json.push_str("{}");
        }
        json.push_str("]]");

        let err = serde_json::from_str::<LimitedRangeCellEdits>(&json)
            .expect_err("expected oversized column count to be rejected");
        assert!(
            err.to_string().contains("max") && err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn limited_f64_vec_rejects_too_many_entries() {
        let max = crate::resource_limits::MAX_RANGE_DIM;
        let mut json = String::from("[");
        for idx in 0..=max {
            if idx > 0 {
                json.push(',');
            }
            json.push_str("0");
        }
        json.push(']');

        let err =
            serde_json::from_str::<LimitedF64Vec>(&json).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains("max") && err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    fn minimal_pivot_config_json() -> serde_json::Value {
        serde_json::json!({
            "rowFields": [],
            "columnFields": [],
            "valueFields": [],
            "filterFields": [],
            "layout": "tabular",
            "subtotals": "none",
            "grandTotals": { "rows": true, "columns": true }
        })
    }

    #[test]
    fn ipc_pivot_config_rejects_too_many_row_fields() {
        let max = crate::resource_limits::MAX_PIVOT_FIELDS;

        let mut cfg = minimal_pivot_config_json();
        let row_fields = (0..=max)
            .map(|idx| {
                serde_json::json!({
                    "sourceField": format!("field-{idx}"),
                    "sortOrder": "ascending",
                    "manualSort": null
                })
            })
            .collect::<Vec<_>>();
        cfg.as_object_mut().unwrap().insert(
            "rowFields".to_string(),
            serde_json::Value::Array(row_fields),
        );

        let err =
            serde_json::from_value::<IpcPivotConfig>(cfg).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn ipc_pivot_config_rejects_too_many_calculated_fields() {
        let max = crate::resource_limits::MAX_PIVOT_CALCULATED_FIELDS;

        let mut cfg = minimal_pivot_config_json();
        let calculated_fields = (0..=max)
            .map(|idx| {
                serde_json::json!({
                    "name": format!("Calc{idx}"),
                    "formula": "=1+1"
                })
            })
            .collect::<Vec<_>>();
        cfg.as_object_mut().unwrap().insert(
            "calculatedFields".to_string(),
            serde_json::Value::Array(calculated_fields),
        );

        let err =
            serde_json::from_value::<IpcPivotConfig>(cfg).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn ipc_pivot_config_rejects_too_many_calculated_items() {
        let max = crate::resource_limits::MAX_PIVOT_CALCULATED_ITEMS;

        let mut cfg = minimal_pivot_config_json();
        let calculated_items = (0..=max)
            .map(|idx| {
                serde_json::json!({
                    "field": "Category",
                    "name": format!("Item{idx}"),
                    "formula": "=1"
                })
            })
            .collect::<Vec<_>>();
        cfg.as_object_mut().unwrap().insert(
            "calculatedItems".to_string(),
            serde_json::Value::Array(calculated_items),
        );

        let err =
            serde_json::from_value::<IpcPivotConfig>(cfg).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn ipc_pivot_config_rejects_too_many_manual_sort_items() {
        let max = crate::resource_limits::MAX_PIVOT_MANUAL_SORT_ITEMS;

        let mut cfg = minimal_pivot_config_json();
        let manual_sort = (0..=max)
            .map(|idx| serde_json::json!({ "type": "text", "value": format!("Item{idx}") }))
            .collect::<Vec<_>>();
        let row_fields = serde_json::json!([{
            "sourceField": "Category",
            "sortOrder": "manual",
            "manualSort": manual_sort
        }]);
        cfg.as_object_mut()
            .unwrap()
            .insert("rowFields".to_string(), row_fields);

        let err =
            serde_json::from_value::<IpcPivotConfig>(cfg).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn ipc_pivot_config_rejects_too_many_filter_allowed_values() {
        let max = crate::resource_limits::MAX_PIVOT_FILTER_ALLOWED_VALUES;

        let mut cfg = minimal_pivot_config_json();
        let allowed = (0..=max)
            .map(|idx| serde_json::json!({ "type": "text", "value": format!("Value{idx}") }))
            .collect::<Vec<_>>();
        let filter_fields = serde_json::json!([{
            "sourceField": "Region",
            "allowed": allowed
        }]);
        cfg.as_object_mut()
            .unwrap()
            .insert("filterFields".to_string(), filter_fields);

        let err =
            serde_json::from_value::<IpcPivotConfig>(cfg).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn ipc_pivot_config_rejects_oversized_pivot_text() {
        let max = crate::resource_limits::MAX_PIVOT_TEXT_BYTES;

        let mut cfg = minimal_pivot_config_json();
        let row_fields = serde_json::json!([{
            "sourceField": "x".repeat(max + 1),
            "sortOrder": "ascending",
            "manualSort": null
        }]);
        cfg.as_object_mut()
            .unwrap()
            .insert("rowFields".to_string(), row_fields);

        let err =
            serde_json::from_value::<IpcPivotConfig>(cfg).expect_err("expected size limit to fail");
        assert!(
            err.to_string().contains(&max.to_string()),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn limited_string_rejects_oversized_payloads() {
        type ShortString = LimitedString<4>;

        let err = serde_json::from_str::<ShortString>("\"abcde\"")
            .expect_err("expected oversized string to fail");
        assert!(
            err.to_string().contains("max") && err.to_string().contains("4"),
            "unexpected error: {err}"
        );
    }

    struct SizeHintSeqDeserializer {
        len: usize,
    }

    struct SizeHintSeqAccess {
        len: usize,
    }

    impl<'de> de::SeqAccess<'de> for SizeHintSeqAccess {
        type Error = de::value::Error;

        fn next_element_seed<T>(&mut self, _seed: T) -> Result<Option<T::Value>, Self::Error>
        where
            T: de::DeserializeSeed<'de>,
        {
            panic!("unexpected element deserialization (size_hint guard should have failed first)");
        }

        fn size_hint(&self) -> Option<usize> {
            Some(self.len)
        }
    }

    impl<'de> serde::Deserializer<'de> for SizeHintSeqDeserializer {
        type Error = de::value::Error;

        fn deserialize_any<V>(self, visitor: V) -> Result<V::Value, Self::Error>
        where
            V: de::Visitor<'de>,
        {
            self.deserialize_seq(visitor)
        }

        fn deserialize_seq<V>(self, visitor: V) -> Result<V::Value, Self::Error>
        where
            V: de::Visitor<'de>,
        {
            visitor.visit_seq(SizeHintSeqAccess { len: self.len })
        }

        serde::forward_to_deserialize_any! {
            bool i8 i16 i32 i64 u8 u16 u32 u64 f32 f64 char str string bytes byte_buf option unit
            unit_struct newtype_struct tuple tuple_struct map struct enum identifier ignored_any
        }
    }

    #[test]
    fn apply_sheet_formatting_deltas_request_rejects_too_many_row_format_deltas() {
        let max = crate::resource_limits::MAX_SHEET_FORMATTING_ROW_DELTAS;
        let err =
            <LimitedSheetRowFormatDeltas as Deserialize>::deserialize(SizeHintSeqDeserializer {
                len: max + 1,
            })
            .expect_err("expected size_hint guard to reject oversized row deltas")
            .to_string();
        assert!(
            err.contains("rowFormats") && err.contains(&max.to_string()),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn apply_sheet_formatting_deltas_request_rejects_too_many_col_format_deltas() {
        let max = crate::resource_limits::MAX_SHEET_FORMATTING_COL_DELTAS;
        let err =
            <LimitedSheetColFormatDeltas as Deserialize>::deserialize(SizeHintSeqDeserializer {
                len: max + 1,
            })
            .expect_err("expected size_hint guard to reject oversized col deltas")
            .to_string();
        assert!(
            err.contains("colFormats") && err.contains(&max.to_string()),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn apply_sheet_formatting_deltas_request_rejects_too_many_cell_format_deltas() {
        let max = crate::resource_limits::MAX_SHEET_FORMATTING_CELL_DELTAS;
        let err =
            <LimitedSheetCellFormatDeltas as Deserialize>::deserialize(SizeHintSeqDeserializer {
                len: max + 1,
            })
            .expect_err("expected size_hint guard to reject oversized cell deltas")
            .to_string();
        assert!(
            err.contains("cellFormats") && err.contains(&max.to_string()),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn apply_sheet_formatting_deltas_request_rejects_too_many_run_columns() {
        let max = crate::resource_limits::MAX_SHEET_FORMATTING_RUN_COLS;
        let err = <LimitedSheetFormatRunsByColDeltas as Deserialize>::deserialize(
            SizeHintSeqDeserializer { len: max + 1 },
        )
        .expect_err("expected size_hint guard to reject oversized run columns")
        .to_string();
        assert!(
            err.contains("formatRunsByCol") && err.contains(&max.to_string()),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn apply_sheet_view_deltas_request_rejects_too_many_col_width_deltas() {
        let max = crate::resource_limits::MAX_SHEET_VIEW_COL_WIDTH_DELTAS;
        let err =
            <LimitedSheetColWidthDeltas as Deserialize>::deserialize(SizeHintSeqDeserializer {
                len: max + 1,
            })
            .expect_err("expected size_hint guard to reject oversized col width deltas")
            .to_string();
        assert!(
            err.contains("colWidths") && err.contains(&max.to_string()),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn apply_sheet_view_deltas_request_rejects_too_many_row_height_deltas() {
        let max = crate::resource_limits::MAX_SHEET_VIEW_ROW_HEIGHT_DELTAS;
        let err =
            <LimitedSheetRowHeightDeltas as Deserialize>::deserialize(SizeHintSeqDeserializer {
                len: max + 1,
            })
            .expect_err("expected size_hint guard to reject oversized row height deltas")
            .to_string();
        assert!(
            err.contains("rowHeights") && err.contains(&max.to_string()),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn apply_sheet_formatting_deltas_request_rejects_too_many_runs_per_column() {
        let max = crate::resource_limits::MAX_SHEET_FORMATTING_RUNS_PER_COL;

        let mut runs = Vec::with_capacity(max + 1);
        for idx in 0..=max {
            runs.push(serde_json::json!({
                "startRow": idx as i64,
                "endRowExclusive": idx as i64 + 1,
                "format": {}
            }));
        }

        let value = serde_json::json!({
            "sheetId": "Sheet1",
            "formatRunsByCol": [{
                "col": 0,
                "runs": runs
            }]
        });

        let err = serde_json::from_value::<ApplySheetFormattingDeltasRequest>(value)
            .expect_err("expected run limit to be enforced during deserialization")
            .to_string();
        assert!(
            err.contains("formatRunsByCol") && err.contains(&max.to_string()),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn limited_macro_permissions_rejects_too_many_entries() {
        let max = crate::resource_limits::MAX_MACRO_PERMISSION_ENTRIES;
        let mut json = String::from("[");
        for idx in 0..=max {
            if idx > 0 {
                json.push(',');
            }
            json.push_str("\"filesystem_read\"");
        }
        json.push(']');

        let err = serde_json::from_str::<LimitedMacroPermissions>(&json)
            .expect_err("expected oversized permissions array to be rejected");
        assert!(
            err.to_string().contains(&max.to_string()),
            "expected error to mention limit: {err}"
        );
    }

    #[test]
    fn limited_vec_rejects_too_many_sheet_ids() {
        type ShortSheetId = LimitedString<4>;
        type SheetIds = LimitedVec<ShortSheetId, 4>;

        let value = serde_json::json!(["a", "b", "c", "d", "e"]);
        let err = serde_json::from_value::<SheetIds>(value)
            .expect_err("expected oversized array to fail");
        assert!(
            err.to_string().contains("max") && err.to_string().contains("4"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_permissions_rejects_too_many_allowlist_entries() {
        let max = crate::resource_limits::MAX_PYTHON_NETWORK_ALLOWLIST_ENTRIES;
        let mut json = String::from(
            "{\"filesystem\":\"none\",\"network\":\"allowlist\",\"networkAllowlist\":[",
        );
        for idx in 0..=max {
            if idx > 0 {
                json.push(',');
            }
            json.push_str("\"example.com\"");
        }
        json.push_str("]}");

        let err = serde_json::from_str::<PythonPermissions>(&json)
            .expect_err("expected oversized allowlist to be rejected");
        assert!(
            err.to_string().contains(&max.to_string()),
            "expected error to mention limit: {err}"
        );
    }

    #[test]
    fn limited_vec_rejects_too_many_print_area_ranges() {
        type PrintRanges = LimitedVec<PrintCellRange, 4>;

        let value = serde_json::json!([
            { "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 1 },
            { "start_row": 2, "end_row": 2, "start_col": 1, "end_col": 1 },
            { "start_row": 3, "end_row": 3, "start_col": 1, "end_col": 1 },
            { "start_row": 4, "end_row": 4, "start_col": 1, "end_col": 1 },
            { "start_row": 5, "end_row": 5, "start_col": 1, "end_col": 1 }
        ]);

        let err = serde_json::from_value::<PrintRanges>(value)
            .expect_err("expected oversized print area ranges to fail");
        assert!(
            err.to_string().contains("max") && err.to_string().contains("4"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn limited_script_code_rejects_oversized_payloads() {
        type ShortCode = LimitedScriptCode<4>;

        let err = serde_json::from_str::<ShortCode>("\"abcde\"")
            .expect_err("expected oversized code to fail");
        assert!(
            err.to_string().contains("Script is too large") && err.to_string().contains("4"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn sheet_formatting_metadata_size_check_rejects_oversized_payloads() {
        let max = crate::resource_limits::MAX_SHEET_FORMATTING_METADATA_BYTES;

        // JSON strings add 2 bytes for the surrounding quotes.
        let ok = JsonValue::String("a".repeat(max.saturating_sub(2)));
        assert!(
            validate_sheet_formatting_metadata_size(&ok).is_ok(),
            "expected max-sized payload to be accepted"
        );

        let too_big = JsonValue::String("a".repeat(max.saturating_sub(1)));
        let err = validate_sheet_formatting_metadata_size(&too_big)
            .expect_err("expected oversized payload to be rejected");
        assert!(
            err.contains("Sheet formatting metadata") && err.contains(&max.to_string()),
            "unexpected error message: {err}"
        );
    }

    #[test]
    fn python_permissions_rejects_allowlist_entries_that_exceed_max_bytes() {
        let max = crate::resource_limits::MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES;
        let oversized_entry = "x".repeat(max + 1);
        let json = format!(
            "{{\"filesystem\":\"none\",\"network\":\"allowlist\",\"networkAllowlist\":[\"{oversized_entry}\"]}}"
        );

        let err = serde_json::from_str::<PythonPermissions>(&json)
            .expect_err("expected oversized allowlist entry to be rejected");
        assert!(
            err.to_string().contains(&max.to_string()),
            "expected error to mention limit: {err}"
        );
    }

    #[test]
    fn sheet_formatting_payload_for_ipc_returns_default_when_too_large() {
        let max = crate::resource_limits::MAX_SHEET_FORMATTING_METADATA_BYTES;
        let too_big = JsonValue::String("a".repeat(max.saturating_sub(1)));

        let clamped = sheet_formatting_payload_for_ipc("Sheet1", Some(&too_big));
        assert_eq!(clamped, default_sheet_formatting_payload());
    }

    #[test]
    fn coerce_save_path_to_xlsx_rewrites_non_workbook_origins() {
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.csv"),
            "/tmp/foo.xlsx",
            "expected .csv saves to coerce to .xlsx"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.xls"),
            "/tmp/foo.xlsx",
            "expected .xls saves to coerce to .xlsx"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.txt"),
            "/tmp/foo.xlsx",
            "expected .txt saves to coerce to .xlsx"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.ods"),
            "/tmp/foo.xlsx",
            "expected .ods saves to coerce to .xlsx"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.parquet"),
            "/tmp/foo.xlsx",
            "expected .parquet saves to coerce to .xlsx"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.xlsx"),
            "/tmp/foo.xlsx",
            "expected .xlsx saves to remain unchanged"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.xlsm"),
            "/tmp/foo.xlsm",
            "expected .xlsm saves to remain unchanged"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.xlsb"),
            "/tmp/foo.xlsb",
            "expected .xlsb saves to remain unchanged"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.xltx"),
            "/tmp/foo.xltx",
            "expected .xltx saves to remain unchanged"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.xltm"),
            "/tmp/foo.xltm",
            "expected .xltm saves to remain unchanged"
        );
        assert_eq!(
            coerce_save_path_to_xlsx("/tmp/foo.xlam"),
            "/tmp/foo.xlam",
            "expected .xlam saves to remain unchanged"
        );
    }

    #[test]
    fn wants_origin_bytes_for_save_path_includes_xlsx_family() {
        for ext in ["xlsx", "xlsm", "xltx", "xltm", "xlam"] {
            let path = format!("/tmp/workbook.{ext}");
            assert!(
                wants_origin_bytes_for_save_path(&path),
                "expected wants_origin_bytes_for_save_path to accept {ext}"
            );
        }

        for ext in ["xlsb", "csv", "xls"] {
            let path = format!("/tmp/workbook.{ext}");
            assert!(
                !wants_origin_bytes_for_save_path(&path),
                "expected wants_origin_bytes_for_save_path to reject {ext}"
            );
        }
    }

    #[test]
    fn set_sheet_visibility_allows_very_hidden_from_ipc_enum() {
        // Ensure the IPC contract (`"veryHidden"`) can be deserialized and applied to the backend
        // state without rejection at the command layer.
        let visibility: SheetVisibility =
            serde_json::from_str("\"veryHidden\"").expect("deserialize veryHidden");
        assert_eq!(visibility, SheetVisibility::VeryHidden);
        assert_eq!(
            serde_json::to_string(&visibility).expect("serialize"),
            "\"veryHidden\""
        );

        // Make sure we have more than one visible sheet; Excel forbids hiding the last visible
        // sheet, and the backend enforces that invariant.
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        set_sheet_visibility_core(&mut state, "Sheet1", visibility)
            .expect("set veryHidden visibility");

        let info = state.workbook_info().expect("workbook info");
        let sheet1 = info
            .sheets
            .iter()
            .find(|sheet| sheet.id == "Sheet1")
            .expect("Sheet1 present");
        assert_eq!(
            sheet1.visibility,
            formula_model::SheetVisibility::VeryHidden
        );
    }

    #[test]
    fn macro_security_status_uses_in_memory_vba_project_signature_part_when_origin_bytes_missing() {
        // Build a minimal OLE container that looks like a VBA signature payload:
        // it contains a `\x05DigitalSignature` stream, but the bytes are not a valid PKCS#7 blob
        // so the verifier reports a parse error.
        let signature_part = {
            let cursor = std::io::Cursor::new(Vec::new());
            let mut ole = cfb::CompoundFile::create(cursor).expect("create signature OLE");
            let mut stream = ole
                .create_stream("\u{0005}DigitalSignature")
                .expect("create signature stream");
            stream
                .write_all(b"not-a-valid-pkcs7")
                .expect("write signature bytes");
            ole.into_inner().into_inner()
        };

        // No need for a fully-formed VBA project OLE for this test: signature parsing happens
        // against the signature part, and we only use the project bytes for fingerprinting.
        let mut workbook = Workbook::new_empty(None);
        workbook.vba_project_bin = Some(vec![1, 2, 3]);
        workbook.vba_project_signature_bin = Some(signature_part);
        workbook.origin_xlsx_bytes = None;

        let trust_store = crate::macro_trust::MacroTrustStore::new_ephemeral();
        let status =
            build_macro_security_status(&mut workbook, None, &trust_store).expect("macro status");
        let sig = status.signature.expect("signature info present");
        assert_eq!(
            sig.status,
            MacroSignatureStatus::SignedParseError,
            "expected signature status to come from vba_project_signature_bin even when origin_xlsx_bytes is None"
        );
    }

    #[test]
    fn macro_security_status_supports_raw_vba_project_signature_part_bytes() {
        // Some XLSM producers store `xl/vbaProjectSignature.bin` as a *raw* PKCS#7/CMS blob (not an
        // OLE container). When the bytes are invalid, the status should be a parse error (not
        // silently treated as unsigned).
        let signature_part = b"not-a-valid-pkcs7".to_vec();

        let mut workbook = Workbook::new_empty(None);
        workbook.vba_project_bin = Some(vec![1, 2, 3]);
        workbook.vba_project_signature_bin = Some(signature_part);
        workbook.origin_xlsx_bytes = None;

        let trust_store = crate::macro_trust::MacroTrustStore::new_ephemeral();
        let status =
            build_macro_security_status(&mut workbook, None, &trust_store).expect("macro status");
        let sig = status.signature.expect("signature info present");
        assert_eq!(sig.status, MacroSignatureStatus::SignedParseError);
    }

    #[test]
    fn macro_security_status_drops_oversized_vba_project_signature_part_from_origin_bytes() {
        let oversized_len = crate::resource_limits::MAX_VBA_PROJECT_SIGNATURE_BIN_BYTES + 1;
        let oversized_part = vec![0u8; oversized_len];

        // The signature part itself doesn't have to be valid; we just want to ensure we don't
        // allocate/parse it when it exceeds our configured max.
        let origin_zip = {
            let cursor = std::io::Cursor::new(Vec::new());
            let mut zip = zip::ZipWriter::new(cursor);
            let options = zip::write::FileOptions::<()>::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("xl/vbaProjectSignature.bin", options)
                .expect("start signature part");
            zip.write_all(&oversized_part)
                .expect("write signature part bytes");
            zip.finish().expect("finish zip").into_inner()
        };

        let mut workbook = Workbook::new_empty(None);
        workbook.vba_project_bin = Some(vec![1, 2, 3]);
        workbook.vba_project_signature_bin = None;
        workbook.origin_xlsx_bytes = Some(std::sync::Arc::<[u8]>::from(origin_zip));

        let trust_store = crate::macro_trust::MacroTrustStore::new_ephemeral();
        let status =
            build_macro_security_status(&mut workbook, None, &trust_store).expect("macro status");
        let sig = status.signature.expect("signature info present");
        assert_eq!(
            sig.status,
            MacroSignatureStatus::Unsigned,
            "expected oversized signature part to be dropped (treated as absent)"
        );
    }

    #[test]
    fn macro_result_from_outcome_sets_error_codes_for_host_limits() {
        let outcome = crate::macros::MacroExecutionOutcome {
            ok: false,
            output: Vec::new(),
            updates: Vec::new(),
            error: Some(format!(
                "macro produced too many cell updates (limit {})",
                crate::resource_limits::MAX_MACRO_UPDATES
            )),
            permission_request: None,
        };
        let result = macro_result_from_outcome(outcome);
        assert_eq!(
            result.error.and_then(|e| e.code),
            Some("macro_updates_limit_exceeded".to_string())
        );

        let outcome = crate::macros::MacroExecutionOutcome {
            ok: false,
            output: Vec::new(),
            updates: Vec::new(),
            error: Some(format!(
                "cell value string is too large (max {} bytes)",
                crate::resource_limits::MAX_CELL_VALUE_STRING_BYTES
            )),
            permission_request: None,
        };
        let result = macro_result_from_outcome(outcome);
        assert_eq!(
            result.error.and_then(|e| e.code),
            Some("macro_cell_value_too_large".to_string())
        );

        let outcome = crate::macros::MacroExecutionOutcome {
            ok: false,
            output: Vec::new(),
            updates: Vec::new(),
            error: Some(format!(
                "cell formula is too large (max {} bytes)",
                crate::resource_limits::MAX_CELL_FORMULA_BYTES
            )),
            permission_request: None,
        };
        let result = macro_result_from_outcome(outcome);
        assert_eq!(
            result.error.and_then(|e| e.code),
            Some("macro_cell_formula_too_large".to_string())
        );
    }

    #[test]
    fn normalize_tab_color_rgb_accepts_rgb_and_argb_hex() {
        assert_eq!(
            normalize_tab_color_rgb("ff00ff").expect("normalize RRGGBB"),
            "FFFF00FF"
        );
        assert_eq!(
            normalize_tab_color_rgb("#ff00ff").expect("normalize #RRGGBB"),
            "FFFF00FF"
        );
        assert_eq!(
            normalize_tab_color_rgb("80ff00ff").expect("normalize AARRGGBB"),
            "80FF00FF"
        );
        assert_eq!(
            normalize_tab_color_rgb("  #FF00FF  ").expect("normalize trims whitespace"),
            "FFFF00FF"
        );
    }

    #[test]
    fn normalize_tab_color_rgb_rejects_invalid_inputs() {
        assert!(normalize_tab_color_rgb("").is_err());
        assert!(normalize_tab_color_rgb("#").is_err());
        assert!(normalize_tab_color_rgb("GG00FF").is_err());
        assert!(normalize_tab_color_rgb("12345").is_err());
        assert!(normalize_tab_color_rgb("1234567").is_err());
        assert!(normalize_tab_color_rgb("123456789").is_err());
    }

    #[test]
    fn workbook_theme_palette_is_exposed_for_rt_simple_fixture() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsx/tests/fixtures/rt_simple.xlsx"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");
        let palette = workbook_theme_palette(&workbook).expect("palette should be present");

        for value in [
            palette.dk1,
            palette.lt1,
            palette.dk2,
            palette.lt2,
            palette.accent1,
            palette.accent2,
            palette.accent3,
            palette.accent4,
            palette.accent5,
            palette.accent6,
            palette.hlink,
            palette.followed_hlink,
        ] {
            assert!(
                value.len() == 7
                    && value.starts_with('#')
                    && value
                        .chars()
                        .skip(1)
                        .all(|c| c.is_ascii_hexdigit() && !c.is_ascii_lowercase()),
                "expected hex color like '#RRGGBB', got {value}"
            );
        }
    }

    #[test]
    fn read_xlsx_preserves_col_widths_and_hidden_cols() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/basic/row-col-attrs.xlsx"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");
        let sheet = workbook.sheets.first().expect("expected a sheet");

        // Fixture sanity check: Column B has a custom width and column C is hidden.
        assert_eq!(sheet.col_widths.get(&1).copied(), Some(25.0));
        assert!(sheet.hidden_cols.contains(&2));
    }

    #[test]
    fn open_workbook_sniffs_csv_even_when_extension_is_xlsx() {
        let dir = TempDir::new().expect("create temp dir");
        let path = dir.path().join("data.xlsx");
        std::fs::write(&path, b"a,b\n1,2\n").expect("write csv bytes");

        let workbook = crate::file_io::read_workbook_blocking(&path).expect("open workbook");
        assert!(
            workbook.origin_xlsx_bytes.is_none(),
            "expected CSV import to not preserve XLSX origin bytes"
        );
        assert!(
            workbook
                .sheets
                .first()
                .is_some_and(|sheet| sheet.columnar.is_some()),
            "expected CSV import to use columnar backing"
        );
    }

    #[test]
    fn open_workbook_sniffs_xlsx_even_when_extension_is_unknown() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsx/tests/fixtures/rt_simple.xlsx"
        ));
        let dir = TempDir::new().expect("create temp dir");
        let path = dir.path().join("data.unknown");
        std::fs::copy(fixture_path, &path).expect("copy xlsx fixture");

        let workbook = crate::file_io::read_workbook_blocking(&path).expect("open workbook");
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected XLSX import to preserve origin bytes"
        );
        assert!(
            workbook
                .sheets
                .first()
                .is_some_and(|sheet| sheet.columnar.is_none()),
            "expected XLSX import not to use columnar backing"
        );
    }

    #[test]
    fn inspect_workbook_encryption_returns_summary_for_encrypted_ooxml() {
        use std::io::Cursor;

        let base_dirs = directories::BaseDirs::new().expect("base dirs");
        let dir = tempfile::Builder::new()
            .prefix("formula-encrypted-ooxml")
            .tempdir_in(base_dirs.home_dir())
            .expect("create temp dir");
        let path = dir.path().join("encrypted.xlsx");

        let xml = r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"><keyData keyBits="256" hashAlgorithm="SHA512"/><keyEncryptors><keyEncryptor><encryptedKey spinCount="100000"/></keyEncryptor></keyEncryptors></encryption>"#;

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut stream = ole
                .create_stream("EncryptionInfo")
                .expect("create EncryptionInfo stream");
            let mut bytes = Vec::new();
            bytes.extend_from_slice(&4u16.to_le_bytes()); // VersionMajor (agile)
            bytes.extend_from_slice(&4u16.to_le_bytes()); // VersionMinor (agile)
            bytes.extend_from_slice(&0x40u32.to_le_bytes()); // Flags
            bytes.extend_from_slice(xml.as_bytes());
            stream.write_all(&bytes).expect("write EncryptionInfo");
        }
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let ole_bytes = ole.into_inner().into_inner();

        std::fs::write(&path, &ole_bytes).expect("write encrypted workbook");

        let summary = inspect_workbook_encryption(LimitedString::<MAX_IPC_PATH_BYTES>(
            path.to_string_lossy().to_string(),
        ))
        .expect("inspect_workbook_encryption should succeed")
        .expect("expected encryption summary");

        assert_eq!(summary.encryption_type, EncryptionTypeDto::Agile);
        assert_eq!(summary.hash_algorithm.as_deref(), Some("SHA512"));
        assert_eq!(summary.spin_count, Some(100000));
        assert_eq!(summary.key_bits, Some(256));
    }

    #[test]
    fn inspect_workbook_encryption_returns_summary_for_encrypted_ooxml_utf16le() {
        use std::io::{Cursor, Write as _};

        let base_dirs = directories::BaseDirs::new().expect("base dirs");
        let dir = tempfile::Builder::new()
            .prefix("formula-encrypted-ooxml-utf16")
            .tempdir_in(base_dirs.home_dir())
            .expect("create temp dir");
        let path = dir.path().join("encrypted.xlsx");

        let xml = r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"><keyData keyBits="256" hashAlgorithm="SHA512"/><keyEncryptors><keyEncryptor><encryptedKey spinCount="100000"/></keyEncryptor></keyEncryptors></encryption>"#;

        let mut xml_utf16 = Vec::new();
        // UTF-16LE BOM.
        xml_utf16.extend_from_slice(&[0xFF, 0xFE]);
        for unit in xml.encode_utf16() {
            xml_utf16.extend_from_slice(&unit.to_le_bytes());
        }
        // A trailing UTF-16 NUL terminator.
        xml_utf16.extend_from_slice(&[0, 0]);

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut stream = ole
                .create_stream("EncryptionInfo")
                .expect("create EncryptionInfo stream");
            let mut bytes = Vec::new();
            bytes.extend_from_slice(&4u16.to_le_bytes()); // VersionMajor (agile)
            bytes.extend_from_slice(&4u16.to_le_bytes()); // VersionMinor (agile)
            bytes.extend_from_slice(&0x40u32.to_le_bytes()); // Flags
            bytes.extend_from_slice(&xml_utf16);
            stream.write_all(&bytes).expect("write EncryptionInfo");
        }
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let ole_bytes = ole.into_inner().into_inner();

        std::fs::write(&path, &ole_bytes).expect("write encrypted workbook");

        let summary = inspect_workbook_encryption(LimitedString::<MAX_IPC_PATH_BYTES>(
            path.to_string_lossy().to_string(),
        ))
        .expect("inspect_workbook_encryption should succeed")
        .expect("expected encryption summary");

        assert_eq!(summary.encryption_type, EncryptionTypeDto::Agile);
        assert_eq!(summary.hash_algorithm.as_deref(), Some("SHA512"));
        assert_eq!(summary.spin_count, Some(100000));
        assert_eq!(summary.key_bits, Some(256));
    }

    #[test]
    fn inspect_workbook_encryption_returns_summary_for_encrypted_ooxml_utf16le_no_bom() {
        use std::io::{Cursor, Write as _};

        let base_dirs = directories::BaseDirs::new().expect("base dirs");
        let dir = tempfile::Builder::new()
            .prefix("formula-encrypted-ooxml-utf16le-no-bom")
            .tempdir_in(base_dirs.home_dir())
            .expect("create temp dir");
        let path = dir.path().join("encrypted.xlsx");

        let xml = r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"><keyData keyBits="256" hashAlgorithm="SHA512"/><keyEncryptors><keyEncryptor><encryptedKey spinCount="100000"/></keyEncryptor></keyEncryptors></encryption>"#;

        let mut xml_utf16 = Vec::new();
        for unit in xml.encode_utf16() {
            xml_utf16.extend_from_slice(&unit.to_le_bytes());
        }
        // A trailing UTF-16 NUL terminator.
        xml_utf16.extend_from_slice(&[0, 0]);

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut stream = ole
                .create_stream("EncryptionInfo")
                .expect("create EncryptionInfo stream");
            let mut bytes = Vec::new();
            bytes.extend_from_slice(&4u16.to_le_bytes()); // VersionMajor (agile)
            bytes.extend_from_slice(&4u16.to_le_bytes()); // VersionMinor (agile)
            bytes.extend_from_slice(&0x40u32.to_le_bytes()); // Flags
            bytes.extend_from_slice(&xml_utf16);
            stream.write_all(&bytes).expect("write EncryptionInfo");
        }
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let ole_bytes = ole.into_inner().into_inner();

        std::fs::write(&path, &ole_bytes).expect("write encrypted workbook");

        let summary = inspect_workbook_encryption(LimitedString::<MAX_IPC_PATH_BYTES>(
            path.to_string_lossy().to_string(),
        ))
        .expect("inspect_workbook_encryption should succeed")
        .expect("expected encryption summary");

        assert_eq!(summary.encryption_type, EncryptionTypeDto::Agile);
        assert_eq!(summary.hash_algorithm.as_deref(), Some("SHA512"));
        assert_eq!(summary.spin_count, Some(100000));
        assert_eq!(summary.key_bits, Some(256));
    }

    #[test]
    fn inspect_workbook_encryption_returns_summary_for_encrypted_ooxml_utf16be_no_bom() {
        use std::io::{Cursor, Write as _};

        let base_dirs = directories::BaseDirs::new().expect("base dirs");
        let dir = tempfile::Builder::new()
            .prefix("formula-encrypted-ooxml-utf16be")
            .tempdir_in(base_dirs.home_dir())
            .expect("create temp dir");
        let path = dir.path().join("encrypted.xlsx");

        let xml = r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"><keyData keyBits="256" hashAlgorithm="SHA512"/><keyEncryptors><keyEncryptor><encryptedKey spinCount="100000"/></keyEncryptor></keyEncryptors></encryption>"#;

        let mut xml_utf16 = Vec::new();
        for unit in xml.encode_utf16() {
            xml_utf16.extend_from_slice(&unit.to_be_bytes());
        }
        // A trailing UTF-16 NUL terminator.
        xml_utf16.extend_from_slice(&[0, 0]);

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut stream = ole
                .create_stream("EncryptionInfo")
                .expect("create EncryptionInfo stream");
            let mut bytes = Vec::new();
            bytes.extend_from_slice(&4u16.to_le_bytes()); // VersionMajor (agile)
            bytes.extend_from_slice(&4u16.to_le_bytes()); // VersionMinor (agile)
            bytes.extend_from_slice(&0x40u32.to_le_bytes()); // Flags
            bytes.extend_from_slice(&xml_utf16);
            stream.write_all(&bytes).expect("write EncryptionInfo");
        }
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let ole_bytes = ole.into_inner().into_inner();

        std::fs::write(&path, &ole_bytes).expect("write encrypted workbook");

        let summary = inspect_workbook_encryption(LimitedString::<MAX_IPC_PATH_BYTES>(
            path.to_string_lossy().to_string(),
        ))
        .expect("inspect_workbook_encryption should succeed")
        .expect("expected encryption summary");

        assert_eq!(summary.encryption_type, EncryptionTypeDto::Agile);
        assert_eq!(summary.hash_algorithm.as_deref(), Some("SHA512"));
        assert_eq!(summary.spin_count, Some(100000));
        assert_eq!(summary.key_bits, Some(256));
    }

    #[test]
    fn inspect_workbook_encryption_rejects_oversized_encryption_info_stream() {
        use std::io::{Cursor, Write as _};

        let base_dirs = directories::BaseDirs::new().expect("base dirs");
        let dir = tempfile::Builder::new()
            .prefix("formula-encrypted-ooxml-oversized")
            .tempdir_in(base_dirs.home_dir())
            .expect("create temp dir");
        let path = dir.path().join("encrypted.xlsx");

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut stream = ole
                .create_stream("EncryptionInfo")
                .expect("create EncryptionInfo stream");
            let oversized = vec![0u8; 1024 * 1024 + 1];
            stream.write_all(&oversized).expect("write EncryptionInfo");
        }
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let ole_bytes = ole.into_inner().into_inner();

        std::fs::write(&path, &ole_bytes).expect("write encrypted workbook");

        let err = inspect_workbook_encryption(LimitedString::<MAX_IPC_PATH_BYTES>(
            path.to_string_lossy().to_string(),
        ))
        .expect_err("expected oversized EncryptionInfo to error");

        assert!(
            err.contains("EncryptionInfo stream is too large"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn inspect_workbook_encryption_returns_summary_for_standard_encrypted_ooxml() {
        use std::io::{Cursor, Write as _};

        let base_dirs = directories::BaseDirs::new().expect("base dirs");
        let dir = tempfile::Builder::new()
            .prefix("formula-standard-encrypted-ooxml")
            .tempdir_in(base_dirs.home_dir())
            .expect("create temp dir");
        let path = dir.path().join("encrypted.xlsx");

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut stream = ole
                .create_stream("EncryptionInfo")
                .expect("create EncryptionInfo stream");
            let mut bytes = Vec::new();
            bytes.extend_from_slice(&4u16.to_le_bytes()); // VersionMajor (standard variants vary)
            bytes.extend_from_slice(&2u16.to_le_bytes()); // VersionMinor (standard)
            bytes.extend_from_slice(&0u32.to_le_bytes()); // Flags
            bytes.extend_from_slice(&32u32.to_le_bytes()); // Header size (EncryptionHeader)

            // Minimal CryptoAPI EncryptionHeader (8 * u32 = 32 bytes).
            bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
            bytes.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
            bytes.extend_from_slice(&0x0000_6610u32.to_le_bytes()); // algId (AES-256)
            bytes.extend_from_slice(&0x0000_8004u32.to_le_bytes()); // algIdHash (SHA1)
            bytes.extend_from_slice(&256u32.to_le_bytes()); // keySize (bits)
            bytes.extend_from_slice(&0u32.to_le_bytes()); // providerType
            bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved1
            bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved2

            stream.write_all(&bytes).expect("write EncryptionInfo");
        }
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let ole_bytes = ole.into_inner().into_inner();

        std::fs::write(&path, &ole_bytes).expect("write encrypted workbook");

        let summary = inspect_workbook_encryption(LimitedString::<MAX_IPC_PATH_BYTES>(
            path.to_string_lossy().to_string(),
        ))
        .expect("inspect_workbook_encryption should succeed")
        .expect("expected encryption summary");

        assert_eq!(summary.encryption_type, EncryptionTypeDto::Standard);
        assert_eq!(summary.cipher.as_deref(), Some("AES-256"));
        assert_eq!(summary.key_size, Some(256));
        assert_eq!(summary.hash_algorithm.as_deref(), Some("SHA1"));
        assert_eq!(summary.spin_count, None);
        assert_eq!(summary.key_bits, None);
    }

    #[cfg(feature = "parquet")]
    #[test]
    fn open_workbook_sniffs_parquet_even_when_extension_is_xlsx() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../packages/data-io/test/fixtures/simple.parquet"
        ));
        let dir = TempDir::new().expect("create temp dir");
        let path = dir.path().join("data.xlsx");
        std::fs::copy(fixture_path, &path).expect("copy parquet fixture");

        let workbook = crate::file_io::read_workbook_blocking(&path).expect("open workbook");
        assert!(
            workbook.origin_xlsx_bytes.is_none(),
            "expected Parquet import to not preserve XLSX origin bytes"
        );
        assert!(
            workbook
                .sheets
                .first()
                .is_some_and(|sheet| sheet.columnar.is_some()),
            "expected Parquet import to use columnar backing"
        );
    }

    #[test]
    fn list_dir_errors_when_entry_limit_reached() {
        use std::fs::{create_dir, remove_dir, File};

        let base_dirs = directories::BaseDirs::new().expect("base dirs");
        let dir = tempfile::Builder::new()
            .prefix("formula-list-dir-entry-limit")
            .tempdir_in(base_dirs.home_dir())
            .expect("create temp dir");
        for idx in 0..crate::resource_limits::MAX_LIST_DIR_ENTRIES {
            let path = dir.path().join(format!("file_{idx}.txt"));
            File::create(path).expect("create temp file");
        }

        let ok = list_dir_blocking(dir.path().to_str().unwrap(), false)
            .expect("expected list_dir to succeed when at the entry limit");
        assert_eq!(
            ok.len(),
            crate::resource_limits::MAX_LIST_DIR_ENTRIES,
            "expected exactly MAX_LIST_DIR_ENTRIES results"
        );

        // Adding even an empty directory should exceed the traversal limit (directory entries
        // count toward the cap, even though only files are returned).
        let extra_dir = dir.path().join("extra_dir");
        create_dir(&extra_dir).expect("create extra dir");

        let err = list_dir_blocking(dir.path().to_str().unwrap(), false)
            .expect_err("expected list_dir to error once entry limit is exceeded");
        assert!(
            err.contains(&format!(
                "Directory listing exceeded limit (max {} entries)",
                crate::resource_limits::MAX_LIST_DIR_ENTRIES
            )),
            "unexpected error: {err}"
        );

        remove_dir(&extra_dir).expect("remove extra dir");

        // Now add one more file and ensure we get a clear error.
        let extra_path = dir.path().join(format!(
            "file_{}.txt",
            crate::resource_limits::MAX_LIST_DIR_ENTRIES
        ));
        File::create(extra_path).expect("create extra temp file");

        let err = list_dir_blocking(dir.path().to_str().unwrap(), false)
            .expect_err("expected list_dir to error once entry limit is exceeded");
        assert!(
            err.contains(&format!(
                "Directory listing exceeded limit (max {} entries)",
                crate::resource_limits::MAX_LIST_DIR_ENTRIES
            )),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn list_dir_errors_when_depth_limit_reached() {
        use std::fs::{create_dir, File};

        let base_dirs = directories::BaseDirs::new().expect("base dirs");
        let dir = tempfile::Builder::new()
            .prefix("formula-list-dir-depth-limit")
            .tempdir_in(base_dirs.home_dir())
            .expect("create temp dir");
        let mut current = dir.path().to_path_buf();
        for depth in 0..=crate::resource_limits::MAX_LIST_DIR_DEPTH {
            current = current.join(format!("d{depth}"));
            create_dir(&current).expect("create nested dir");
        }

        // Add a file at the deepest level to ensure traversal would want to descend that far.
        File::create(current.join("deep.txt")).expect("create deep file");

        let err = list_dir_blocking(dir.path().to_str().unwrap(), true)
            .expect_err("expected list_dir to error once depth limit is reached");
        assert!(
            err.contains(&format!(
                "Directory listing exceeded depth limit (max {} levels)",
                crate::resource_limits::MAX_LIST_DIR_DEPTH
            )),
            "unexpected error: {err}"
        );
    }

    #[cfg(unix)]
    #[test]
    fn list_dir_does_not_follow_symlinked_directories() {
        use std::fs::{create_dir, File};
        use std::os::unix::fs::symlink;

        // Create temp dirs inside the allowed filesystem scope (home dir) so `list_dir_blocking`
        // can traverse them during tests.
        let base_dirs = directories::BaseDirs::new().expect("base dirs");
        let root: TempDir = tempfile::Builder::new()
            .prefix("formula-list-dir-symlink-root")
            .tempdir_in(base_dirs.home_dir())
            .expect("create root temp dir");
        let outside: TempDir = tempfile::Builder::new()
            .prefix("formula-list-dir-symlink-outside")
            .tempdir_in(base_dirs.home_dir())
            .expect("create outside temp dir");

        // A file that should be discoverable via a real directory in the requested subtree.
        let real_dir = root.path().join("real");
        create_dir(&real_dir).expect("create real dir");
        File::create(real_dir.join("inside.txt")).expect("create inside file");

        // A file that exists outside of the subtree. We create a symlinked directory inside the
        // root pointing to it; list_dir must not traverse it.
        File::create(outside.path().join("outside.txt")).expect("create outside file");
        symlink(outside.path(), root.path().join("link")).expect("create symlinked dir");

        let out = list_dir_blocking(root.path().to_str().unwrap(), true)
            .expect("list_dir should succeed");

        assert!(
            out.iter().any(|entry| entry.path.ends_with("inside.txt")),
            "expected to see inside.txt, got {out:?}"
        );
        assert!(
            out.iter().all(|entry| !entry.path.ends_with("outside.txt")),
            "expected not to traverse symlinked dir, got {out:?}"
        );
    }

    #[test]
    fn file_read_helpers_reject_non_regular_files() {
        let dir = TempDir::new().expect("create temp dir");
        let path = dir.path();

        let err = read_text_file_blocking(path).expect_err("expected directory read to fail");
        assert!(
            err.contains("Path is not a regular file"),
            "unexpected error: {err}"
        );

        let err = read_binary_file_blocking(path).expect_err("expected directory read to fail");
        assert!(
            err.contains("Path is not a regular file"),
            "unexpected error: {err}"
        );

        let err = read_binary_file_range_blocking(path, 0, 1)
            .expect_err("expected directory read to fail");
        assert!(
            err.contains("Path is not a regular file"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn typescript_migration_interpreter_applies_basic_range_and_cell_assignments() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = crate::state::AppState::new();
        state.load_workbook(workbook);

        let code = r#"
export default async function main(ctx) {
  const sheet = ctx.activeSheet;
  const fill = 7;
  const values = Array.from({ length: 2 }, () => Array(2).fill(fill));
  await sheet.getRange("$A$1:$B$2").setValues(values);
  await sheet.getRange("C1").setValue(fill);

  const formula = "=A1+B1";
  await sheet.getRange("D1:E1").setFormulas(Array.from({ length: 1 }, () => Array(2).fill(formula)));
 }
"#;

        let result = run_typescript_migration_script(&mut state, code);
        assert!(result.ok, "expected ok, got {:?}", result.error);

        let a1 = state.get_cell("Sheet1", 0, 0).expect("A1 exists");
        assert_eq!(a1.value.display(), "7");

        let b2 = state.get_cell("Sheet1", 1, 1).expect("B2 exists");
        assert_eq!(b2.value.display(), "7");

        let c1 = state.get_cell("Sheet1", 0, 2).expect("C1 exists");
        assert_eq!(c1.value.display(), "7");

        let d1 = state.get_cell("Sheet1", 0, 3).expect("D1 exists");
        assert_eq!(d1.formula.as_deref(), Some("=A1+B1"));
        assert_eq!(d1.value.display(), "14");

        let e1 = state.get_cell("Sheet1", 0, 4).expect("E1 exists");
        assert_eq!(e1.formula.as_deref(), Some("=A1+B1"));
        assert_eq!(e1.value.display(), "14");
    }

    #[test]
    fn typescript_migration_interpreter_rejects_oversized_fill_matrices() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = crate::state::AppState::new();
        state.load_workbook(workbook);

        let dim = crate::resource_limits::MAX_RANGE_DIM;
        let code = format!(
            r#"
export default async function main(ctx) {{
  const sheet = ctx.activeSheet;
  await sheet.getRange("A1").setValues(Array.from({{ length: {dim} }}, () => Array({dim}).fill(1)));
}}
"#
        );

        let result = run_typescript_migration_script(&mut state, &code);
        assert!(!result.ok, "expected script to fail due to size limits");
        let err = result.error.unwrap_or_default();
        assert!(
            err.contains("too large") || err.contains("max"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn typescript_migration_interpreter_respects_active_sheet_from_macro_context() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        let mut state = crate::state::AppState::new();
        state.load_workbook(workbook);

        // Simulate the user having Sheet2 selected when kicking off a migration validation.
        state
            .set_macro_ui_context("Sheet2", 0, 0, None)
            .expect("set macro ui context");

        let code = r#"
export default async function main(ctx) {
  const sheet = ctx.activeSheet;
  sheet.range("A1").value = 99;
}
"#;

        let result = run_typescript_migration_script(&mut state, code);
        assert!(result.ok, "expected ok, got {:?}", result.error);

        let sheet2_a1 = state.get_cell("Sheet2", 0, 0).expect("Sheet2!A1 exists");
        assert_eq!(sheet2_a1.value.display(), "99");

        let sheet1_a1 = state.get_cell("Sheet1", 0, 0).expect("Sheet1!A1 exists");
        assert_eq!(
            sheet1_a1.value.display(),
            "",
            "expected Sheet1!A1 to remain empty"
        );
    }

    #[test]
    fn extract_imported_sheet_background_images_from_xlsx_bytes_finds_picture_relationship() {
        use std::io::Cursor;
        use zip::write::FileOptions;
        use zip::ZipWriter;

        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <picture r:id="rId2"/>
</worksheet>"#;

        let sheet_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml).unwrap();
        zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
        zip.write_all(workbook_rels).unwrap();
        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet_xml).unwrap();
        zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options).unwrap();
        zip.write_all(sheet_rels).unwrap();
        zip.start_file("xl/media/image1.png", options).unwrap();
        zip.write_all(b"png-bytes").unwrap();

        let bytes = zip.finish().unwrap().into_inner();

        let extracted = extract_imported_sheet_background_images_from_xlsx_bytes(&bytes);
        assert_eq!(extracted.len(), 1);
        assert_eq!(extracted[0].sheet_name, "Sheet1");
        assert_eq!(extracted[0].worksheet_part, "xl/worksheets/sheet1.xml");
        assert_eq!(extracted[0].image_id, "image1.png");
        assert_eq!(extracted[0].mime_type, "image/png");
        assert_eq!(extracted[0].bytes_base64, STANDARD.encode(b"png-bytes"));
    }

    #[cfg(feature = "desktop")]
    #[test]
    fn restart_app_command_signature_compiles() {
        // We can't exercise a real restart in tests, but we can assert that the command
        // compiles with Tauri's supported restart API.
        let _ = restart_app as fn(tauri::WebviewWindow) -> Result<(), String>;
    }
}
