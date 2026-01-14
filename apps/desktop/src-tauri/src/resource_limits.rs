//! Backend-enforced resource limits for APIs that accept untrusted inputs.
//!
//! The desktop backend accepts requests from a WebView. If that WebView is compromised it can
//! send arbitrarily large ranges/exports, or access very large directories/files. These
//! conservative limits ensure the backend fails fast and deterministically instead of exhausting
//! memory/CPU (OOM/freeze).

use serde::de::{self, Visitor};
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde_json::value::RawValue;
use serde_json::Value as JsonValue;
use std::fmt;

/// Maximum number of filesystem entries processed by the `list_dir` command.
///
/// The limit is enforced even when `recursive=false`, since a single directory can contain an
/// unbounded number of files.
///
/// Note: `list_dir` returns file entries only, but this limit counts all directory entries
/// encountered during traversal (files + directories). This ensures CPU usage is bounded even for
/// directory trees that contain very few files.
pub const MAX_LIST_DIR_ENTRIES: usize = 50_000;

/// Maximum recursion depth for the `list_dir` command when `recursive=true`.
///
/// Depth starts at `0` for the root directory passed to `list_dir`.
pub const MAX_LIST_DIR_DEPTH: usize = 20;

/// Maximum number of cells that `get_range`/`set_range` will process in a single call.
pub const MAX_RANGE_CELLS_PER_CALL: usize = 1_000_000;

/// Maximum number of rows or columns allowed in a single `get_range`/`set_range` call.
///
/// This prevents pathological "skinny" ranges (e.g. 1×1_000_000) which can still be expensive even
/// when the total cell count is capped.
pub const MAX_RANGE_DIM: usize = 10_000;

/// Maximum number of cells that a single PDF export is allowed to render.
///
/// PDF generation is substantially more expensive than a plain range read, so this limit is
/// intentionally more conservative than `MAX_RANGE_CELLS_PER_CALL`.
pub const MAX_PDF_CELLS_PER_CALL: usize = 250_000;

/// Maximum size (in bytes) of a generated PDF before it is rejected.
///
/// Note: this is checked after generation, so it primarily protects against extremely large
/// base64 responses rather than CPU time.
pub const MAX_PDF_BYTES: usize = 20 * 1024 * 1024; // 20 MiB

/// Maximum number of filesystem paths accepted for a single `file-dropped` event payload.
///
/// Drag/drop payloads originate from the OS and can be arbitrarily large (e.g. the user drops a
/// massive directory selection). We cap the emitted path list to keep memory usage deterministic
/// and avoid sending unexpectedly large payloads over the JS event bridge.
pub const MAX_FILE_DROPPED_PATHS: usize = 1_000;

/// Maximum size (in bytes) of a single path string accepted for a `file-dropped` payload.
///
/// This is a UTF-8 byte length check performed on `PathBuf::to_string_lossy()` output.
pub const MAX_FILE_DROPPED_PATH_BYTES: usize = 8_192;

/// Maximum size (in bytes) of a marketplace extension package download.
///
/// The desktop backend base64-encodes the downloaded package for IPC, which expands the payload by
/// ~33%. This limit is applied to the *raw* bytes to keep memory usage bounded even if the
/// marketplace server (or a network attacker) attempts to return an arbitrarily large response.
pub const MAX_MARKETPLACE_PACKAGE_BYTES: usize = 20 * 1024 * 1024; // 20 MiB

/// Maximum size (in bytes) of individual marketplace metadata headers (e.g. signatures/hashes).
///
/// Headers are expected to be small (an ed25519 signature is 64 bytes, SHA-256 is 32 bytes), but
/// the backend still treats them as untrusted input and enforces a cap to avoid large string
/// allocations.
pub const MAX_MARKETPLACE_HEADER_BYTES: usize = 4 * 1024; // 4 KiB

/// Maximum on-disk workbook size (in bytes) allowed for `open_workbook`.
///
/// Large spreadsheet files can lead to unbounded memory/CPU usage during parsing (even when the
/// package itself is streamed) due to decompression and intermediate allocations. This serves as a
/// coarse backstop against accidental opens and malicious inputs.
pub const MAX_WORKBOOK_OPEN_BYTES: u64 = 512 * 1024 * 1024; // 512 MiB

/// Maximum size (in bytes) of the XLSX/XLSM package we will retain in
/// [`crate::file_io::Workbook::origin_xlsx_bytes`].
///
/// Retaining the original package enables patch-based saves (we can patch edited worksheet XML
/// onto the original ZIP). However, keeping arbitrarily large packages in memory is a footgun:
/// a compromised webview can OOM the Rust backend by opening a huge workbook.
///
/// When a workbook exceeds this limit we still open it, but we *do not* keep the raw bytes.
/// Subsequent saves fall back to the regeneration-based code path.
pub const MAX_ORIGIN_XLSX_BYTES: usize = 64 * 1024 * 1024; // 64 MiB

/// Maximum uncompressed size (in bytes) for any single ZIP part we will inflate while preserving
/// DrawingML-related parts during workbook open.
///
/// Preserved drawing parts are used to round-trip unsupported drawing/chart structures when the
/// user saves. This is a best-effort feature: if a workbook contains extremely large drawing parts
/// (or forged ZIP metadata), we prefer dropping preservation rather than risking OOM.
pub const MAX_PRESERVED_DRAWING_PART_BYTES: usize = 64 * 1024 * 1024; // 64 MiB

/// Maximum total uncompressed size (in bytes) across all ZIP parts inflated while preserving
/// DrawingML-related parts during workbook open.
pub const MAX_PRESERVED_DRAWING_TOTAL_BYTES: usize = 128 * 1024 * 1024; // 128 MiB

/// Maximum uncompressed size (in bytes) for any single ZIP part we will inflate while preserving
/// pivot-related parts during workbook open.
///
/// Like preserved drawings, pivot preservation is best-effort: when pivot parts are enormous we
/// drop preservation rather than allocating unbounded memory.
pub const MAX_PRESERVED_PIVOT_PART_BYTES: usize = 64 * 1024 * 1024; // 64 MiB

/// Maximum total uncompressed size (in bytes) across all ZIP parts inflated while preserving
/// pivot-related parts during workbook open.
pub const MAX_PRESERVED_PIVOT_TOTAL_BYTES: usize = 128 * 1024 * 1024; // 128 MiB

/// Maximum uncompressed size (in bytes) for any single ZIP part when inflating an XLSX/XLSM
/// package for *IPC-only* inspection/extraction.
///
/// Some IPC commands need to inspect the original workbook bytes (e.g. extracting imported
/// background images or drawing objects). These code paths are best-effort and should never
/// allocate unbounded memory when given a ZIP bomb.
pub const MAX_IPC_XLSX_PACKAGE_PART_BYTES: usize = 64 * 1024 * 1024; // 64 MiB

/// Maximum total uncompressed size (in bytes) across all ZIP parts when inflating an XLSX/XLSM
/// package for *IPC-only* inspection/extraction.
pub const MAX_IPC_XLSX_PACKAGE_TOTAL_BYTES: usize = 128 * 1024 * 1024; // 128 MiB

/// Maximum number of bytes allowed for `open_workbook`.
///
/// This reads the `FORMULA_MAX_WORKBOOK_OPEN_BYTES` environment variable.
/// - In debug builds we allow any value (so developers can test both branches).
/// - In release builds we only honor stricter values (tightening); relaxing the limit is ignored.
pub fn max_workbook_open_bytes() -> u64 {
    let default = MAX_WORKBOOK_OPEN_BYTES;
    let Some(value) = std::env::var("FORMULA_MAX_WORKBOOK_OPEN_BYTES")
        .ok()
        .and_then(|v| v.parse::<u64>().ok())
    else {
        return default;
    };
    if cfg!(debug_assertions) {
        value
    } else {
        value.min(default)
    }
}

/// Maximum size (in bytes) of the optional `xl/vbaProject.bin` payload extracted when opening
/// workbooks.
///
/// This guards against memory-DoS from oversized optional ZIP parts, while still allowing
/// legitimate macro-enabled workbooks to round-trip with macros preserved.
pub const MAX_VBA_PROJECT_BIN_BYTES: usize = 16 * 1024 * 1024; // 16 MiB

/// Maximum size (in bytes) of the optional `xl/vbaProjectSignature.bin` payload extracted when
/// opening workbooks.
pub const MAX_VBA_PROJECT_SIGNATURE_BIN_BYTES: usize = 1024 * 1024; // 1 MiB

/// Maximum number of bytes allowed for `Workbook::origin_xlsx_bytes`.
///
/// This reads the `FORMULA_MAX_ORIGIN_XLSX_BYTES` environment variable.
/// - In debug builds we allow any value (so developers can test both branches).
/// - In release builds we only honor stricter values (tightening); relaxing the limit is ignored.
pub fn max_origin_xlsx_bytes() -> usize {
    let default = MAX_ORIGIN_XLSX_BYTES;
    let Some(value) = std::env::var("FORMULA_MAX_ORIGIN_XLSX_BYTES")
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
    else {
        return default;
    };
    if cfg!(debug_assertions) {
        value
    } else {
        value.min(default)
    }
}

/// Maximum size (in bytes) of an imported worksheet background image payload returned over IPC.
///
/// Worksheet background images are extracted from untrusted XLSX packages and base64-encoded for
/// the Tauri IPC boundary. This limit is applied to the *raw* image bytes before base64 encoding.
pub const MAX_IMPORTED_SHEET_BACKGROUND_IMAGE_BYTES: usize = 10 * 1024 * 1024; // 10 MiB

/// Maximum total size (in bytes) of all imported worksheet background images returned over IPC.
///
/// This is a defense-in-depth cap to avoid sending very large base64 payloads across IPC when a
/// workbook contains many worksheets with background images.
pub const MAX_IMPORTED_SHEET_BACKGROUND_IMAGES_TOTAL_BYTES: usize = 20 * 1024 * 1024; // 20 MiB

/// Maximum number of macro permissions accepted over IPC.
///
/// This is intentionally small because macros currently support a small fixed set of permissions.
/// Bounding the array length prevents a compromised WebView from sending huge arrays and forcing
/// unbounded allocations during JSON deserialization.
pub const MAX_MACRO_PERMISSION_ENTRIES: usize = 8;

/// Maximum number of log lines captured from a single VBA macro execution.
///
/// VBA code can emit host output via `Debug.Print`/`MsgBox`. The desktop backend forwards this
/// output to the WebView over IPC. This limit ensures a compromised workbook or WebView cannot
/// force the backend to accumulate an unbounded `Vec<String>`.
pub const MAX_MACRO_OUTPUT_LINES: usize = 1_000;

/// Maximum total UTF-8 byte size of log output captured from a single VBA macro execution.
///
/// This is a coarse limit over the sum of bytes of all captured log lines (before IPC framing).
/// Together with `MAX_MACRO_OUTPUT_LINES` it bounds both memory usage and the size of the IPC
/// response payload.
pub const MAX_MACRO_OUTPUT_BYTES: usize = 1 * 1024 * 1024; // 1 MiB

pub(crate) fn macro_output_max_lines() -> usize {
    let default = MAX_MACRO_OUTPUT_LINES;
    let Some(value) = std::env::var("FORMULA_MACRO_OUTPUT_MAX_LINES")
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
    else {
        return default;
    };
    // Allow relaxing limits only in debug builds; in release we only honor stricter settings.
    if cfg!(debug_assertions) {
        value
    } else {
        value.min(default)
    }
}

pub(crate) fn macro_output_max_bytes() -> usize {
    let default = MAX_MACRO_OUTPUT_BYTES;
    let Some(value) = std::env::var("FORMULA_MACRO_OUTPUT_MAX_BYTES")
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
    else {
        return default;
    };
    if cfg!(debug_assertions) {
        value
    } else {
        value.min(default)
    }
}

/// Maximum number of cell updates a single VBA macro execution is allowed to generate.
///
/// Macros can write to many cells and trigger large recalculation fanout (dependent formulas),
/// producing large `Vec<CellUpdateData>` payloads. This limit bounds memory usage and prevents
/// sending extremely large IPC responses. When exceeded, macro execution is aborted with an
/// explicit runtime error (no truncation) so callers can handle the failure deterministically.
pub const MAX_MACRO_UPDATES: usize = 10_000;

// -----------------------------------------------------------------------------
// Native Python runner limits
// -----------------------------------------------------------------------------

/// Maximum number of entries allowed in the Python `network_allowlist` IPC field.
///
/// This allowlist is semantically small (hostnames / CIDR blocks) but untrusted. Bounding the
/// length prevents a compromised WebView from allocating arbitrarily large vectors while parsing
/// JSON IPC requests.
pub const MAX_PYTHON_NETWORK_ALLOWLIST_ENTRIES: usize = 256;

/// Maximum size in bytes of a single Python `network_allowlist` entry accepted over IPC.
///
/// Entries are expected to be small (domain names / IP ranges). Bounding the per-entry size avoids
/// large string allocations during IPC deserialization.
pub const MAX_PYTHON_NETWORK_ALLOWLIST_ENTRY_BYTES: usize = 512;

/// Maximum number of bytes captured from the native Python runner's stderr stream.
///
/// The WebView can execute arbitrary code in the Python subprocess. A compromised/buggy script can
/// spam stderr indefinitely; capturing stderr without a cap can OOM the Rust host.
///
/// 1 MiB keeps error messages and tracebacks useful while bounding host memory.
pub const MAX_PYTHON_STDERR_BYTES: usize = 1 * 1024 * 1024; // 1 MiB

/// Maximum size (in bytes) of a single JSON protocol line emitted by the native Python runner.
///
/// The host reads newline-delimited JSON messages from the runner. Each line is bounded to prevent
/// pathological allocations (e.g. a single multi-gigabyte "line") in the Rust process.
pub const MAX_PYTHON_PROTOCOL_LINE_BYTES: usize = 1 * 1024 * 1024; // 1 MiB

/// Maximum number of distinct cell updates accumulated during a single Python execution.
///
/// Updates are deduped by `(sheet_id, row, col)` before being returned to the frontend, but without
/// an overall cap a script can still touch an unbounded number of cells and force the backend to
/// retain (and eventually serialize) a massive update payload.
///
/// This is aligned with `MAX_RANGE_CELLS_PER_CALL` so a single "large range" operation doesn't
/// immediately fail, but scripts that fan out to multiple large operations will be rejected.
pub const MAX_PYTHON_UPDATES: usize = MAX_RANGE_CELLS_PER_CALL;

/// Maximum number of row-formatting deltas accepted by `apply_sheet_formatting_deltas` in a single
/// IPC call.
///
/// This bounds IPC allocations from untrusted WebView input and keeps the resulting persisted
/// formatting snapshot reasonably sized.
pub const MAX_SHEET_FORMATTING_ROW_DELTAS: usize = 50_000;

/// Maximum number of column-formatting deltas accepted by `apply_sheet_formatting_deltas` in a
/// single IPC call.
pub const MAX_SHEET_FORMATTING_COL_DELTAS: usize = 20_000;

/// Maximum number of column-width deltas accepted by `apply_sheet_view_deltas` in a single IPC call.
///
/// This is distinct from formatting deltas but has a similar risk profile: a compromised webview
/// could send an arbitrarily large array and force the backend to allocate unbounded memory during
/// JSON deserialization.
pub const MAX_SHEET_VIEW_COL_WIDTH_DELTAS: usize = 20_000;

/// Maximum number of row-height deltas accepted by `apply_sheet_view_deltas` in a single IPC call.
///
/// Row height edits can be applied to large ranges; this cap prevents a compromised webview from
/// sending pathologically large payloads that would allocate huge vectors during deserialization.
pub const MAX_SHEET_VIEW_ROW_HEIGHT_DELTAS: usize = 50_000;

/// Maximum number of cell-formatting deltas accepted by `apply_sheet_formatting_deltas` in a
/// single IPC call.
pub const MAX_SHEET_FORMATTING_CELL_DELTAS: usize = 200_000;

/// Maximum number of columns that can have their range-run formatting replaced in a single
/// `apply_sheet_formatting_deltas` call.
pub const MAX_SHEET_FORMATTING_RUN_COLS: usize = 20_000;

/// Maximum number of format runs allowed for a single column in `formatRunsByCol`.
///
/// Runs should be a compressed representation (ranges), so this can be kept fairly conservative.
pub const MAX_SHEET_FORMATTING_RUNS_PER_COL: usize = 4_096;

/// Maximum size (in bytes) of the per-sheet formatting metadata stored under
/// `SHEET_FORMATTING_METADATA_KEY` (`formula_ui_formatting`).
///
/// The formatting snapshot is persisted into the workbook (and encrypted document store). This cap
/// prevents a compromised WebView from writing arbitrarily large blobs to disk.
pub const MAX_SHEET_FORMATTING_METADATA_BYTES: usize = 2 * 1024 * 1024; // 2 MiB

/// Maximum size (in bytes) of a single UI style JSON payload accepted over IPC.
///
/// Formatting deltas include `format` objects (sheet default / row / col / cell / run formats). These
/// should be small (a handful of properties), but they originate from an untrusted WebView.
/// Bounding their serialized size prevents memory-DoS from a compromised frontend sending a single
/// huge `format` object inside an otherwise bounded delta list.
pub const MAX_SHEET_FORMATTING_FORMAT_BYTES: usize = 16 * 1024; // 16 KiB

/// Maximum size (in bytes) of a single string cell value accepted over IPC (`set_cell` /
/// `set_range`).
///
/// Rationale: Even with range-level cell count limits, a compromised webview could send a single
/// huge string and force the backend to allocate excessive memory while deserializing the command
/// payload. Excel caps cell text at 32,767 characters; we enforce a conservative per-cell byte
/// limit to bound allocations while still allowing large (but reasonable) inputs.
pub const MAX_CELL_VALUE_STRING_BYTES: usize = 64 * 1024; // 64 KiB

/// Maximum size (in bytes) of a single cell formula accepted over IPC (`set_cell` / `set_range`).
///
/// Rationale: The formula parser enforces Excel's 8,192-character limit, but without an IPC-level
/// bound a compromised webview can still force the backend to deserialize an arbitrarily large
/// string before we ever attempt to parse it.
pub const MAX_CELL_FORMULA_BYTES: usize = 32 * 1024; // 32 KiB

/// Maximum number of fields allowed in each pivot table axis (rows, columns, filters, values).
///
/// Pivot layouts rarely contain more than a handful of fields; allowing dozens is generous for real
/// workbooks while still preventing a compromised webview from sending a pathological payload that
/// forces large allocations during deserialization or expensive pivot processing.
pub const MAX_PIVOT_FIELDS: usize = 64;

/// Maximum number of calculated fields allowed in a pivot config.
///
/// Calculated fields can contain arbitrary formulas which are parsed/validated; keeping this
/// bounded prevents the backend from doing unbounded formula work and reduces memory pressure.
pub const MAX_PIVOT_CALCULATED_FIELDS: usize = 64;

/// Maximum number of calculated items allowed in a pivot config.
///
/// Calculated items multiply work because each item is scoped to a field and evaluated across the
/// pivot's unique values. Real-world workbooks typically use very few calculated items.
pub const MAX_PIVOT_CALCULATED_ITEMS: usize = 256;

/// Maximum number of manual sort items allowed for a single pivot field.
///
/// Manual sorting enumerates individual pivot items. For high-cardinality fields this becomes
/// unwieldy in the UI and expensive in the backend; cap it to ensure pivot configs remain small.
pub const MAX_PIVOT_MANUAL_SORT_ITEMS: usize = 512;

/// Maximum number of allowed values accepted in a single pivot filter allow-list.
///
/// The allow-list is specified over IPC and would otherwise be unbounded; large lists can cause
/// excessive allocations and slow filter evaluation.
pub const MAX_PIVOT_FILTER_ALLOWED_VALUES: usize = 1_000;

/// Maximum size (in bytes) for any user-provided pivot string (field names, formulas, item text).
///
/// Pivot configs cross the IPC boundary as JSON. Bounding string sizes prevents a compromised
/// webview from forcing large allocations and keeps pivot metadata reasonable.
pub const MAX_PIVOT_TEXT_BYTES: usize = 4 * 1024; // 4 KiB

/// A `String` wrapper that rejects inputs larger than `MAX_BYTES` during deserialization.
///
/// This is intended for privileged IPC endpoints: we want to fail fast *during* deserialization so
/// a compromised WebView can't force large allocations by sending huge JSON strings.
///
/// Note: Some formats (e.g. JSON strings with escapes) may still require buffering/decoding inside
/// the deserializer before a `&str` is available to the visitor. We still reject the vast majority
/// of cases (including unescaped base64) without allocating by implementing `visit_borrowed_str`.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct LimitedString<const MAX_BYTES: usize>(String);

impl<const MAX_BYTES: usize> LimitedString<MAX_BYTES> {
    #[inline]
    pub fn into_inner(self) -> String {
        self.0
    }

    #[inline]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl<const MAX_BYTES: usize> AsRef<str> for LimitedString<MAX_BYTES> {
    #[inline]
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl<const MAX_BYTES: usize> std::ops::Deref for LimitedString<MAX_BYTES> {
    type Target = str;

    #[inline]
    fn deref(&self) -> &Self::Target {
        self.as_str()
    }
}

impl<const MAX_BYTES: usize> From<LimitedString<MAX_BYTES>> for String {
    #[inline]
    fn from(value: LimitedString<MAX_BYTES>) -> Self {
        value.0
    }
}

impl<const MAX_BYTES: usize> Serialize for LimitedString<MAX_BYTES> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        serializer.serialize_str(self.as_str())
    }
}

impl<'de, const MAX_BYTES: usize> Deserialize<'de> for LimitedString<MAX_BYTES> {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        struct LimitedStringVisitor<const MAX: usize>;

        impl<'de, const MAX: usize> Visitor<'de> for LimitedStringVisitor<MAX> {
            type Value = LimitedString<MAX>;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                write!(formatter, "a string up to {MAX} bytes")
            }

            fn visit_borrowed_str<E>(self, value: &'de str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                let len = value.as_bytes().len();
                if len > MAX {
                    return Err(E::custom(format!(
                        "string exceeds maximum size ({len} > {MAX} bytes)"
                    )));
                }
                Ok(LimitedString(value.to_owned()))
            }

            fn visit_str<E>(self, value: &str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                let len = value.as_bytes().len();
                if len > MAX {
                    return Err(E::custom(format!(
                        "string exceeds maximum size ({len} > {MAX} bytes)"
                    )));
                }
                Ok(LimitedString(value.to_owned()))
            }

            fn visit_string<E>(self, value: String) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                let len = value.as_bytes().len();
                if len > MAX {
                    return Err(E::custom(format!(
                        "string exceeds maximum size ({len} > {MAX} bytes)"
                    )));
                }
                Ok(LimitedString(value))
            }
        }

        deserializer.deserialize_str(LimitedStringVisitor::<MAX_BYTES>)
    }
}

/// A `serde_json::Value` wrapper that rejects inputs larger than `MAX_BYTES` (in raw JSON bytes)
/// during deserialization.
///
/// This is intended for privileged IPC endpoints that need to accept arbitrary JSON objects while
/// still defending against a compromised WebView attempting to allocate huge nested values.
///
/// Implementation notes:
/// - We deserialize to `&RawValue` first (borrowed from the input) so we can check the raw JSON
///   length without allocating a full `Value`.
/// - After the size check passes, we parse the raw JSON into a `serde_json::Value`.
#[derive(Clone, Debug, PartialEq, Serialize)]
#[serde(transparent)]
pub struct LimitedJsonValue<const MAX_BYTES: usize>(pub JsonValue);

impl<const MAX_BYTES: usize> LimitedJsonValue<MAX_BYTES> {
    #[inline]
    pub fn into_inner(self) -> JsonValue {
        self.0
    }

    #[inline]
    pub fn as_ref(&self) -> &JsonValue {
        &self.0
    }
}

impl<const MAX_BYTES: usize> std::ops::Deref for LimitedJsonValue<MAX_BYTES> {
    type Target = JsonValue;

    #[inline]
    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl<const MAX_BYTES: usize> From<JsonValue> for LimitedJsonValue<MAX_BYTES> {
    #[inline]
    fn from(value: JsonValue) -> Self {
        Self(value)
    }
}

impl<const MAX_BYTES: usize> From<LimitedJsonValue<MAX_BYTES>> for JsonValue {
    #[inline]
    fn from(value: LimitedJsonValue<MAX_BYTES>) -> Self {
        value.0
    }
}

impl<'de, const MAX_BYTES: usize> Deserialize<'de> for LimitedJsonValue<MAX_BYTES> {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let raw = <&RawValue as Deserialize>::deserialize(deserializer)?;
        let len = raw.get().as_bytes().len();
        if len > MAX_BYTES {
            return Err(de::Error::custom(format!(
                "JSON payload is too large (max {MAX_BYTES} bytes)"
            )));
        }

        serde_json::from_str::<JsonValue>(raw.get())
            .map(LimitedJsonValue)
            .map_err(de::Error::custom)
    }
}
