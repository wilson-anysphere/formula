use formula_model::rich_text::RichText;

/// A shared string table entry (`BrtSI`) from the workbook shared strings part (typically
/// `xl/sharedStrings.bin`).
///
/// XLSB shared strings can contain rich text runs and/or phonetic (ruby) data.
/// We preserve the rich structure (and raw bytes for rich/phonetic values) so
/// a future writer can round-trip without data loss.
#[derive(Debug, Clone, PartialEq)]
pub struct SharedString {
    /// Parsed rich text. Plain strings have `runs.is_empty()`.
    pub rich_text: RichText,
    /// Opaque run formatting blobs, in the same order as `rich_text.runs`.
    ///
    /// XLSB rich runs reference font/style records from `xl/styles.bin`. We do
    /// not decode those yet; this stores the raw run formatting bytes so the
    /// original `BrtSI` can be reconstructed if needed.
    pub run_formats: Vec<Vec<u8>>,
    /// Opaque phonetic (ruby) payload, if present.
    pub phonetic: Option<Vec<u8>>,
    /// Raw `BrtSI` record payload bytes (excluding the record header).
    ///
    /// Only populated for rich/phonetic strings. Plain shared strings can be
    /// losslessly represented by `rich_text.text` alone.
    pub raw_si: Option<Vec<u8>>,
}

impl SharedString {
    pub fn plain_text(&self) -> &str {
        self.rich_text.plain_text()
    }
}
