use std::collections::HashMap;

use formula_model::{
    indexed_color_argb, Alignment, Border, BorderEdge, BorderStyle, CalculationMode, Color,
    DateSystem, Fill, FillPattern, Font, HorizontalAlignment, Protection, SheetVisibility, Style,
    TabColor, VerticalAlignment, WorkbookProtection, WorkbookWindow, WorkbookWindowState,
};

use super::{externsheet, records, strings, BiffVersion};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum BoundSheetType {
    Worksheet,
    MacroSheet,
    Chart,
    VisualBasicModule,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct BoundSheetInfo {
    pub(crate) name: String,
    pub(crate) offset: usize,
    /// Sheet visibility state (raw BIFF `hsState` value).
    pub(crate) hs_state: u8,
    /// Best-effort mapping of [`BoundSheetInfo::hs_state`] into a model visibility enum.
    ///
    /// Unknown values are preserved in [`BoundSheetInfo::hs_state`] and mapped to `None`.
    pub(crate) sheet_visibility: Option<SheetVisibility>,
    /// Sheet type (raw BIFF `dt` value).
    pub(crate) dt: u8,
    /// Best-effort mapping of [`BoundSheetInfo::dt`] into a simplified sheet type enum.
    ///
    /// Unknown values are preserved in [`BoundSheetInfo::dt`] and mapped to `None`.
    pub(crate) sheet_type: Option<BoundSheetType>,
}

// Record ids used by workbook-global parsing.
// See [MS-XLS] sections:
// - CODEPAGE: 2.4.52
// - BoundSheet8: 2.4.28
// - 1904: 2.4.169
// - FORMAT / FORMAT2: 2.4.90
// - PALETTE: 2.4.155
// - FONT: 2.4.92
// - XF: 2.4.353
// - WINDOW1: 2.4.346
const RECORD_CODEPAGE: u16 = 0x0042;
const RECORD_BOUNDSHEET: u16 = 0x0085;
const RECORD_1904: u16 = 0x0022;
const RECORD_CALCCOUNT: u16 = 0x000C;
const RECORD_CALCMODE: u16 = 0x000D;
const RECORD_PRECISION: u16 = 0x000E;
const RECORD_DELTA: u16 = 0x0010;
const RECORD_ITERATION: u16 = 0x0011;
// PROTECT [MS-XLS 2.4.203] (workbook globals: lock structure)
const RECORD_PROTECT: u16 = 0x0012;
// PASSWORD [MS-XLS 2.4.191] (workbook globals: protection password hash)
const RECORD_PASSWORD: u16 = 0x0013;
// WINDOWPROTECT [MS-XLS 2.4.347] (lock workbook windows)
const RECORD_WINDOWPROTECT: u16 = 0x0019;
const RECORD_WINDOW1: u16 = 0x003D;
const RECORD_FORMAT_BIFF8: u16 = 0x041E;
const RECORD_FORMAT2_BIFF5: u16 = 0x001E;
const RECORD_PALETTE: u16 = 0x0092;
const RECORD_FONT: u16 = 0x0031;
const RECORD_XF: u16 = 0x00E0;
const RECORD_SAVERECALC: u16 = 0x005F;
// EXTERNSHEET [MS-XLS 2.4.102] stores the mapping table for ixti -> sheet indices.
//
// This record can be large in workbooks with many 3D references and may be split across one or
// more `CONTINUE` records.
const RECORD_EXTERNSHEET: u16 = 0x0017;
// SHEETEXT [MS-XLS 2.4.269] (BIFF8 only) stores extended sheet metadata such as sheet tab color.
const RECORD_SHEETEXT: u16 = 0x0862;
// NAME [MS-XLS 2.4.150] (defined names / named ranges) may contain large formulas or description
// strings and may be split across one or more `CONTINUE` records.
const RECORD_NAME: u16 = 0x0018;

// XF type/protection flags: bit 2 is fStyle in BIFF5/8.
const XF_FLAG_STYLE: u16 = 0x0004;

// WINDOW1.grbit option flags [MS-XLS 2.4.346].
//
// Bit 0 (`fHidden`) indicates the workbook window is hidden (View -> Window -> Hide).
// Bit 1 (`fIconic`) indicates the workbook window is minimized.
// Bit 6 (`fMaximized`) indicates the workbook window is maximized.
const WINDOW1_GRBIT_HIDDEN: u16 = 0x0001;
const WINDOW1_GRBIT_ICONIC: u16 = 0x0002;
const WINDOW1_GRBIT_MAXIMIZED: u16 = 0x0040;

// Cap the number of workbook-global warnings we retain.
//
// Malformed/corrupt `.xls` files may contain many truncated records; without a cap, the warnings
// vector can grow without bound and consume large amounts of memory.
const MAX_GLOBAL_WARNINGS: usize = 100;
const GLOBAL_WARNINGS_SUPPRESSED_MSG: &str = "additional workbook-global warnings suppressed";

fn strip_embedded_nuls(s: &mut String) {
    if s.contains('\0') {
        s.retain(|c| c != '\0');
    }
}

fn hs_state_to_sheet_visibility(hs_state: u8) -> Option<SheetVisibility> {
    // BoundSheet8.hsState [MS-XLS 2.4.28]
    // 0x00 = visible, 0x01 = hidden, 0x02 = very hidden.
    match hs_state {
        0x00 => Some(SheetVisibility::Visible),
        0x01 => Some(SheetVisibility::Hidden),
        0x02 => Some(SheetVisibility::VeryHidden),
        _ => None,
    }
}

fn dt_to_sheet_type(dt: u8) -> Option<BoundSheetType> {
    // BoundSheet8.dt [MS-XLS 2.4.28]
    // 0x00 = worksheet (or dialog sheet), 0x01 = macro sheet, 0x02 = chart, 0x06 = VB module.
    match dt {
        0x00 => Some(BoundSheetType::Worksheet),
        0x01 => Some(BoundSheetType::MacroSheet),
        0x02 => Some(BoundSheetType::Chart),
        0x06 => Some(BoundSheetType::VisualBasicModule),
        _ => None,
    }
}

fn push_warning(out: &mut BiffWorkbookGlobals, msg: String) {
    if out.warnings.len() < MAX_GLOBAL_WARNINGS {
        out.warnings.push(msg);
        return;
    }

    // Once we hit the cap, emit a single note so callers/users know warnings were dropped.
    if out.warnings.len() == MAX_GLOBAL_WARNINGS {
        out.warnings
            .push(GLOBAL_WARNINGS_SUPPRESSED_MSG.to_string());
    }
}

/// Scan the workbook-global BIFF substream for a `CODEPAGE` record.
///
/// The result is used to decode 8-bit (ANSI) strings such as BIFF5 short strings and BIFF8
/// compressed `XLUnicodeString` payloads.
///
/// This scan is best-effort: it stops at the workbook-global `EOF` record, the next `BOF` record,
/// or the first malformed/truncated physical record. When the codepage is missing, defaults to the
/// Excel/Windows "ANSI" codepage (`1252`).
pub(crate) fn parse_biff_codepage(workbook_stream: &[u8]) -> u16 {
    let Ok(iter) = records::BestEffortSubstreamIter::from_offset(workbook_stream, 0) else {
        return 1252;
    };

    for record in iter {
        match record.record_id {
            // CODEPAGE [MS-XLS 2.4.52]
            RECORD_CODEPAGE => {
                if record.data.len() >= 2 {
                    return u16::from_le_bytes([record.data[0], record.data[1]]);
                }
            }
            // EOF terminates the workbook global substream.
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    // Default "ANSI" codepage used by Excel on Windows.
    1252
}

pub(crate) fn parse_biff_bound_sheets(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
) -> Result<Vec<BoundSheetInfo>, String> {
    let mut out = Vec::new();

    for record in records::BestEffortSubstreamIter::from_offset(workbook_stream, 0)? {
        match record.record_id {
            // BoundSheet8 [MS-XLS 2.4.28]
            RECORD_BOUNDSHEET => {
                if record.data.len() < 7 {
                    continue;
                }

                let sheet_offset = u32::from_le_bytes([
                    record.data[0],
                    record.data[1],
                    record.data[2],
                    record.data[3],
                ]) as usize;
                let hs_state = record.data[4];
                let dt = record.data[5];
                let Ok((mut name, _)) =
                    strings::parse_biff_short_string(&record.data[6..], biff, codepage)
                else {
                    continue;
                };
                strip_embedded_nuls(&mut name);
                out.push(BoundSheetInfo {
                    name,
                    offset: sheet_offset,
                    hs_state,
                    sheet_visibility: hs_state_to_sheet_visibility(hs_state),
                    dt,
                    sheet_type: dt_to_sheet_type(dt),
                });
            }
            // EOF terminates the workbook global substream.
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

/// Workbook-global BIFF records needed for stable number format and date system import.
#[derive(Debug, Clone)]
pub(crate) struct BiffWorkbookGlobals {
    /// True when the workbook stream contains a BIFF `FILEPASS` record, indicating it is
    /// encrypted/password-protected.
    pub(crate) is_encrypted: bool,
    pub(crate) date_system: DateSystem,
    pub(crate) calculation_mode: Option<CalculationMode>,
    pub(crate) iterative_enabled: Option<bool>,
    pub(crate) iterative_max_iterations: Option<u32>,
    pub(crate) iterative_max_change: Option<f64>,
    pub(crate) full_precision: Option<bool>,
    pub(crate) calculate_before_save: Option<bool>,
    pub(crate) workbook_protection: WorkbookProtection,
    pub(crate) active_tab_index: Option<u16>,
    pub(crate) workbook_window: Option<WorkbookWindow>,
    /// BIFF8 `EXTERNSHEET` entries (XTI), used for resolving 3D references (`ixti`) in formulas.
    pub(crate) extern_sheets: Vec<externsheet::ExternSheetEntry>,
    formats: HashMap<u16, String>,
    palette: Vec<u32>,
    fonts: Vec<BiffFont>,
    xfs: Vec<BiffXf>,
    /// Sheet tab colors parsed from BIFF8 `SHEETEXT` records, in stream order.
    ///
    /// These correspond to sheets in workbook order (same as `BoundSheet8` order) for BIFF8
    /// workbooks produced by Excel. When the counts differ, callers should treat this as
    /// best-effort metadata.
    pub(crate) sheet_tab_colors: Vec<Option<TabColor>>,
    pub(crate) warnings: Vec<String>,
}

impl Default for BiffWorkbookGlobals {
    fn default() -> Self {
        Self {
            is_encrypted: false,
            date_system: DateSystem::Excel1900,
            calculation_mode: None,
            iterative_enabled: None,
            iterative_max_iterations: None,
            iterative_max_change: None,
            full_precision: None,
            calculate_before_save: None,
            workbook_protection: WorkbookProtection::default(),
            active_tab_index: None,
            workbook_window: None,
            extern_sheets: Vec::new(),
            formats: HashMap::new(),
            palette: Vec::new(),
            fonts: Vec::new(),
            xfs: Vec::new(),
            sheet_tab_colors: Vec::new(),
            warnings: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct BiffFont {
    name: String,
    size_100pt: u16,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    color_idx: u16,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BiffXfKind {
    Cell,
    Style,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct BiffXfApplyFlags {
    num_fmt: bool,
    font: bool,
    alignment: bool,
    border: bool,
    fill: bool,
    protection: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct BiffXfProtection {
    locked: bool,
    hidden: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct BiffXfAlignment {
    horizontal: Option<HorizontalAlignment>,
    vertical: Option<VerticalAlignment>,
    wrap_text: bool,
    rotation: Option<i16>,
    indent: Option<u16>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct BiffXfBorderEdge {
    style: u8,
    color_idx: u16,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct BiffXfBorder {
    left: BiffXfBorderEdge,
    right: BiffXfBorderEdge,
    top: BiffXfBorderEdge,
    bottom: BiffXfBorderEdge,
    diagonal: BiffXfBorderEdge,
    diagonal_up: bool,
    diagonal_down: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct BiffXfFill {
    pattern: u8,
    fg_color_idx: u16,
    bg_color_idx: u16,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct BiffResolvedXfDiff {
    num_fmt: bool,
    font: bool,
    fill: bool,
    border: bool,
    alignment: bool,
    protection: bool,
}

impl BiffResolvedXfDiff {
    fn is_interesting(self) -> bool {
        self.num_fmt || self.font || self.fill || self.border || self.alignment || self.protection
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub(crate) struct BiffXf {
    pub(crate) font_idx: u16,
    pub(crate) num_fmt_id: u16,
    pub(crate) kind: Option<BiffXfKind>,
    pub(crate) parent_xf: Option<u16>,
    apply: BiffXfApplyFlags,
    protection: BiffXfProtection,
    alignment: BiffXfAlignment,
    border: BiffXfBorder,
    fill: BiffXfFill,
}

impl BiffWorkbookGlobals {
    /// Resolve an Excel number format code string for the given `xf_index`.
    ///
    /// Precedence:
    /// 1. `numFmtId == 0` → `None` ("General")
    /// 2. workbook `FORMAT` record → exact code
    /// 3. `formula_format::builtin_format_code(numFmtId)` → built-in code
    /// 4. otherwise → stable placeholder (`__builtin_numFmtId:{numFmtId}`)
    pub(crate) fn resolve_number_format_code(&self, xf_index: u32) -> Option<String> {
        let xf = self.xfs.get(xf_index as usize)?;
        let num_fmt_id = xf.num_fmt_id;

        if num_fmt_id == 0 {
            return None;
        }

        if let Some(code) = self.formats.get(&num_fmt_id) {
            return Some(code.clone());
        }

        if let Some(code) = formula_format::builtin_format_code(num_fmt_id) {
            return Some(code.to_string());
        }

        Some(format!(
            "{}{num_fmt_id}",
            formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX
        ))
    }

    #[allow(dead_code)]
    pub(crate) fn xf_count(&self) -> usize {
        self.xfs.len()
    }

    #[allow(dead_code)]
    pub(crate) fn resolve_style(&self, xf_index: u32) -> Style {
        let mut cache: Vec<Option<Style>> = Vec::new();
        if cache.try_reserve_exact(self.xfs.len()).is_err() {
            return Style::default();
        }
        cache.resize_with(self.xfs.len(), || None);
        let mut stack = Vec::new();
        self.resolve_style_inner(xf_index as usize, &mut cache, &mut stack)
            .unwrap_or_default()
    }

    #[allow(dead_code)]
    pub(crate) fn resolve_all_styles(&self) -> Vec<Style> {
        let mut cache: Vec<Option<Style>> = Vec::new();
        if cache.try_reserve_exact(self.xfs.len()).is_err() {
            return Vec::new();
        }
        cache.resize_with(self.xfs.len(), || None);
        let mut stack: Vec<usize> = Vec::new();
        for idx in 0..self.xfs.len() {
            let _ = self.resolve_style_inner(idx, &mut cache, &mut stack);
        }
        cache.into_iter().map(|s| s.unwrap_or_default()).collect()
    }

    /// Return a boolean mask (`xf_index -> is_interesting`) indicating which XF records resolve to a
    /// non-default [`Style`].
    ///
    /// This is used to avoid recording per-cell XF indices for cells that only use the default
    /// style (which can be the vast majority of cells in a sheet).
    pub(crate) fn xf_is_interesting_mask(&self) -> Vec<bool> {
        let mut cache: Vec<Option<BiffResolvedXfDiff>> = Vec::new();
        if cache.try_reserve_exact(self.xfs.len()).is_err() {
            return Vec::new();
        }
        cache.resize_with(self.xfs.len(), || None);
        let mut stack: Vec<usize> = Vec::new();
        let mut out = Vec::new();
        let _ = out.try_reserve_exact(self.xfs.len());
        for idx in 0..self.xfs.len() {
            let diff = self
                .resolve_xf_diff_inner(idx, &mut cache, &mut stack)
                .unwrap_or_default();
            out.push(diff.is_interesting());
        }
        out
    }

    /// Resolve the [`Style`] values for each XF index where `used_mask[idx] == true`.
    ///
    /// Styles are returned as `(xf_index, style)` pairs. Callers can then intern those styles into
    /// their destination workbook.
    pub(crate) fn resolve_styles_for_used_mask(&self, used_mask: &[bool]) -> Vec<(usize, Style)> {
        let mut cache: Vec<Option<Style>> = Vec::new();
        if cache.try_reserve_exact(self.xfs.len()).is_err() {
            return Vec::new();
        }
        cache.resize_with(self.xfs.len(), || None);
        let mut stack: Vec<usize> = Vec::new();
        let mut out: Vec<(usize, Style)> = Vec::new();

        let len = self.xfs.len().min(used_mask.len());
        for idx in 0..len {
            if !used_mask[idx] {
                continue;
            }
            let style = self
                .resolve_style_inner(idx, &mut cache, &mut stack)
                .unwrap_or_default();
            out.push((idx, style));
        }

        out
    }

    fn resolve_style_inner(
        &self,
        xf_index: usize,
        cache: &mut [Option<Style>],
        stack: &mut Vec<usize>,
    ) -> Option<Style> {
        if xf_index >= self.xfs.len() {
            log::warn!("XF index {xf_index} out of range (xf_count={})", self.xfs.len());
            return Some(Style::default());
        }

        if let Some(style) = cache[xf_index].clone() {
            return Some(style);
        }

        if stack.contains(&xf_index) {
            log::warn!("cycle detected while resolving XF inheritance at index {xf_index}");
            return Some(Style::default());
        }

        stack.push(xf_index);

        let xf = self.xfs[xf_index];
        let kind = xf.kind.unwrap_or(BiffXfKind::Cell);

        // Base style: parent XF (best-effort).
        let mut base = if let Some(parent) = xf.parent_xf {
            let parent_idx = parent as usize;
            if parent_idx != xf_index && parent_idx < self.xfs.len() {
                self.resolve_style_inner(parent_idx, cache, stack)
                    .unwrap_or_default()
            } else {
                Style::default()
            }
        } else {
            Style::default()
        };

        // Always apply style XFs; for cell XFs, obey attribute flags (best-effort).
        let apply = |flag: bool| kind == BiffXfKind::Style || flag;

        // Some BIFF writers appear to leave the attribute flags unset even when `ifmt` is
        // meaningful. Since number formats are critical for correct value rendering (especially
        // dates), treat any non-General `ifmt` as applied even when the "apply" bit is missing.
        if apply(xf.apply.num_fmt) || xf.num_fmt_id != 0 {
            base.number_format = self.resolve_number_format_code(xf_index as u32);
        }

        if apply(xf.apply.font) {
            base.font = self.resolve_font(xf.font_idx);
        }

        if apply(xf.apply.fill) {
            base.fill = self.resolve_fill(xf.fill);
        }

        if apply(xf.apply.border) {
            base.border = self.resolve_border(xf.border);
        }

        // Some BIFF writers appear to leave the "apply alignment" bit unset even when alignment
        // fields are meaningful. Similar to the number-format fallback above, treat non-default
        // alignment fields as applied even when the corresponding "apply" bit is missing.
        let alignment_non_default = self.alignment_is_non_default(xf.alignment);
        if apply(xf.apply.alignment) || alignment_non_default {
            base.alignment = self.resolve_alignment(xf.alignment);
        }

        if apply(xf.apply.protection) {
            base.protection = self.resolve_protection(xf.protection);
        }

        stack.pop();
        cache[xf_index] = Some(base.clone());
        Some(base)
    }

    fn resolve_xf_diff_inner(
        &self,
        xf_index: usize,
        cache: &mut [Option<BiffResolvedXfDiff>],
        stack: &mut Vec<usize>,
    ) -> Option<BiffResolvedXfDiff> {
        if xf_index >= self.xfs.len() {
            return Some(BiffResolvedXfDiff::default());
        }

        if let Some(diff) = cache[xf_index] {
            return Some(diff);
        }

        if stack.contains(&xf_index) {
            log::warn!("cycle detected while resolving XF inheritance at index {xf_index}");
            return Some(BiffResolvedXfDiff::default());
        }
        stack.push(xf_index);

        let xf = self.xfs[xf_index];
        let kind = xf.kind.unwrap_or(BiffXfKind::Cell);

        let mut base = if let Some(parent) = xf.parent_xf {
            let parent_idx = parent as usize;
            if parent_idx != xf_index && parent_idx < self.xfs.len() {
                self.resolve_xf_diff_inner(parent_idx, cache, stack)
                    .unwrap_or_default()
            } else {
                BiffResolvedXfDiff::default()
            }
        } else {
            BiffResolvedXfDiff::default()
        };

        let apply = |flag: bool| kind == BiffXfKind::Style || flag;

        if apply(xf.apply.num_fmt) || xf.num_fmt_id != 0 {
            base.num_fmt = xf.num_fmt_id != 0;
        }

        if apply(xf.apply.font) {
            base.font = self.font_is_non_default(xf.font_idx);
        }

        if apply(xf.apply.fill) {
            base.fill = xf.fill.pattern != 0;
        }

        if apply(xf.apply.border) {
            base.border = self.border_is_non_default(xf.border);
        }

        // Like style resolution, treat non-default alignment fields as applied even when the
        // "apply" bit is missing. This ensures we don't incorrectly classify a cell's XF record as
        // uninteresting (and therefore skip per-cell XF indices) when it only differs by
        // alignment.
        let alignment_non_default = self.alignment_is_non_default(xf.alignment);
        if apply(xf.apply.alignment) || alignment_non_default {
            base.alignment = alignment_non_default;
        }

        if apply(xf.apply.protection) {
            base.protection = self.protection_is_non_default(xf.protection);
        }

        stack.pop();
        cache[xf_index] = Some(base);
        Some(base)
    }

    fn font_is_non_default(&self, ifnt: u16) -> bool {
        let Some(base_font) = self.fonts.first() else {
            return false;
        };

        // BIFF quirk: font index 4 is reserved and omitted from the FONT record stream.
        let idx = if ifnt >= 4 { ifnt - 1 } else { ifnt } as usize;
        let Some(font) = self.fonts.get(idx) else {
            return false;
        };

        font != base_font
    }

    fn border_is_non_default(&self, border: BiffXfBorder) -> bool {
        border.left.style != 0
            || border.right.style != 0
            || border.top.style != 0
            || border.bottom.style != 0
            || border.diagonal.style != 0
            || border.diagonal_up
            || border.diagonal_down
            // If a border color is set but the style is "none", it does not affect rendering, so we
            // intentionally ignore `color_idx` when `style == 0`.
    }

    fn alignment_is_non_default(&self, alignment: BiffXfAlignment) -> bool {
        alignment.horizontal.is_some()
            || alignment.vertical.is_some()
            || alignment.wrap_text
            || alignment.rotation.is_some()
            || alignment.indent.is_some()
    }

    fn protection_is_non_default(&self, protection: BiffXfProtection) -> bool {
        protection.locked != true || protection.hidden != false
    }

    fn resolve_color_idx(&self, idx: u16) -> Option<Color> {
        // Excel uses multiple sentinels for "automatic" depending on the record/field.
        // - 0x0040 is the "automatic" ICV value used by many BIFF structures.
        // - 0x7FFF is used by some records such as FONT.
        if idx == 0x0040 || idx == 0x7FFF {
            return Some(Color::Auto);
        }

        // Palette entries correspond to indices starting at 8.
        if idx >= 8 {
            let pal_idx = (idx - 8) as usize;
            if let Some(&argb) = self.palette.get(pal_idx) {
                return Some(Color::Argb(argb));
            }
        }

        if let Some(argb) = indexed_color_argb(idx) {
            return Some(Color::Argb(argb));
        }

        Some(Color::Indexed(idx))
    }

    fn resolve_font(&self, ifnt: u16) -> Option<Font> {
        let Some(base_font) = self.fonts.first() else {
            return None;
        };

        // BIFF quirk: font index 4 is reserved and omitted from the FONT record stream.
        // When referencing fonts at indices >= 4, subtract 1 to index into the stored list.
        let idx = if ifnt >= 4 { ifnt - 1 } else { ifnt } as usize;
        let Some(font) = self.fonts.get(idx) else {
            log::warn!(
                "XF references out-of-range font index {ifnt} (resolved idx={idx}, font_count={})",
                self.fonts.len()
            );
            return None;
        };

        if font == base_font {
            return None;
        }

        Some(Font {
            name: Some(font.name.clone()),
            size_100pt: Some(font.size_100pt),
            bold: font.bold,
            italic: font.italic,
            underline: font.underline,
            strike: font.strike,
            color: self.resolve_color_idx(font.color_idx),
        })
    }

    fn resolve_fill(&self, fill: BiffXfFill) -> Option<Fill> {
        // BIFF uses a pattern code to indicate if a fill is present; pattern 0 means no fill.
        if fill.pattern == 0 {
            return None;
        }

        // BIFF fill patterns map directly to the OOXML `ST_PatternType` enumeration used by
        // `.xlsx`. `formula_model::FillPattern` intentionally preserves the XML token as a string
        // (`Other`) so it can round-trip losslessly.
        //
        // IMPORTANT: `FillPattern::Other` must contain a *valid* OOXML `patternType` token because
        // the XLSX writer will emit it verbatim. Avoid storing BIFF-specific placeholders like
        // `biff:<n>` here, as that would generate invalid `.xlsx` output.
        let pattern = match fill.pattern {
            1 => FillPattern::Solid,
            2 => FillPattern::Other("mediumGray".to_string()),
            3 => FillPattern::Other("darkGray".to_string()),
            4 => FillPattern::Other("lightGray".to_string()),
            5 => FillPattern::Other("darkHorizontal".to_string()),
            6 => FillPattern::Other("darkVertical".to_string()),
            7 => FillPattern::Other("darkDown".to_string()),
            8 => FillPattern::Other("darkUp".to_string()),
            9 => FillPattern::Other("darkGrid".to_string()),
            10 => FillPattern::Other("darkTrellis".to_string()),
            11 => FillPattern::Other("lightHorizontal".to_string()),
            12 => FillPattern::Other("lightVertical".to_string()),
            13 => FillPattern::Other("lightDown".to_string()),
            14 => FillPattern::Other("lightUp".to_string()),
            15 => FillPattern::Other("lightGrid".to_string()),
            16 => FillPattern::Other("lightTrellis".to_string()),
            0x11 => FillPattern::Gray125,
            0x12 => FillPattern::Other("gray0625".to_string()),
            other => {
                log::warn!("unknown BIFF fill pattern {other}; falling back to solid fill");
                FillPattern::Solid
            }
        };

        // Only interpret colors when the pattern is meaningful.
        let fg_color = self.resolve_color_idx(fill.fg_color_idx);
        let bg_color = self.resolve_color_idx(fill.bg_color_idx);

        Some(Fill {
            pattern,
            fg_color,
            bg_color,
        })
    }

    fn resolve_border(&self, border: BiffXfBorder) -> Option<Border> {
        let to_style = |code: u8| match code {
            0 => BorderStyle::None,
            // BIFF line style codes (subset), see [MS-XLS] 2.5.12:
            // 1=thin, 2=medium, 3=dashed, 4=dotted, 5=thick, 6=double.
            //
            // Excel defines additional styles (hair, dash-dot variants, etc.). `formula_model`
            // supports a smaller set; map those variants to the closest representable value so the
            // border is not dropped entirely during `.xls` -> `.xlsx` round-tripping.
            1 | 7 => BorderStyle::Thin, // 7 = hair
            2 => BorderStyle::Medium,
            3 | 8 | 9 | 10 | 11 | 12 | 13 => BorderStyle::Dashed,
            4 => BorderStyle::Dotted,
            5 => BorderStyle::Thick,
            6 => BorderStyle::Double,
            other => {
                log::warn!("unknown BIFF border style {other}; treating as thin");
                BorderStyle::Thin
            }
        };

        let edge = |e: BiffXfBorderEdge| {
            let style = to_style(e.style);
            BorderEdge {
                style,
                // Avoid treating default/unused border colors as meaningful when the border is "none".
                color: (style != BorderStyle::None).then(|| self.resolve_color_idx(e.color_idx))
                    .flatten(),
            }
        };

        let out = Border {
            left: edge(border.left),
            right: edge(border.right),
            top: edge(border.top),
            bottom: edge(border.bottom),
            diagonal: edge(border.diagonal),
            diagonal_up: border.diagonal_up,
            diagonal_down: border.diagonal_down,
        };

        if out == Border::default() {
            None
        } else {
            Some(out)
        }
    }

    fn resolve_alignment(&self, alignment: BiffXfAlignment) -> Option<Alignment> {
        let out = Alignment {
            horizontal: alignment.horizontal,
            vertical: alignment.vertical,
            wrap_text: alignment.wrap_text,
            rotation: alignment.rotation,
            indent: alignment.indent,
        };

        if out == Alignment::default() {
            None
        } else {
            Some(out)
        }
    }

    fn resolve_protection(&self, protection: BiffXfProtection) -> Option<Protection> {
        let out = Protection {
            locked: protection.locked,
            hidden: protection.hidden,
        };

        if out == Protection::default() {
            None
        } else {
            Some(out)
        }
    }
}

pub(crate) fn parse_biff_workbook_globals(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
) -> Result<BiffWorkbookGlobals, String> {
    let mut out = BiffWorkbookGlobals::default();

    // BIFF workbook streams always start with a `BOF` record at offset 0. Treat `FILEPASS`
    // (encryption) as meaningful only when that invariant holds, otherwise we risk flagging
    // non-Excel/garbled streams as encrypted when they merely contain the byte pattern.
    let starts_with_bof = matches!(
        records::read_biff_record(workbook_stream, 0),
        Some((record_id, _)) if records::is_bof_record(record_id)
    );

    let mut saw_eof = false;
    let mut continuation_parse_failed = false;
    let mut saw_external_supbook = false;

    let allows_continuation: fn(u16) -> bool = match biff {
        BiffVersion::Biff5 => workbook_globals_allows_continuation_biff5,
        BiffVersion::Biff8 => workbook_globals_allows_continuation_biff8,
    };

    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                push_warning(&mut out, format!("malformed BIFF record: {err}"));
                break;
            }
        };
        let record_id = record.record_id;
        let data = record.data.as_ref();

        // BOF indicates the start of a new substream; the workbook globals contain
        // a single BOF at offset 0, so a second BOF means we're past the globals
        // section (even if the EOF record is missing).
        if record.offset != 0 && records::is_bof_record(record_id) {
            saw_eof = true;
            break;
        }

        match record_id {
            // FILEPASS [MS-XLS 2.4.105]
            //
            // This record indicates the workbook is encrypted/password-protected. We do not
            // attempt to parse encryption details here; encryption is handled by the workbook-stream
            // reader before globals parsing.
            //
            // When the caller has already decrypted the stream (e.g. legacy XOR obfuscation), the
            // `FILEPASS` record remains present but subsequent records are now readable, so we do
            // not stop scanning.
            records::RECORD_FILEPASS if starts_with_bof => {
                out.is_encrypted = true;
            }
            // PROTECT [MS-XLS 2.4.203] (workbook globals: lock structure)
            RECORD_PROTECT => {
                if data.len() < 2 {
                    push_warning(&mut out, format!(
                        "truncated PROTECT record at offset {}",
                        record.offset
                    ));
                    continue;
                }
                let flag = u16::from_le_bytes([data[0], data[1]]);
                out.workbook_protection.lock_structure = flag != 0;
            }
            // WINDOWPROTECT [MS-XLS 2.4.347]
            RECORD_WINDOWPROTECT => {
                if data.len() < 2 {
                    push_warning(&mut out, format!(
                        "truncated WINDOWPROTECT record at offset {}",
                        record.offset
                    ));
                    continue;
                }
                let flag = u16::from_le_bytes([data[0], data[1]]);
                out.workbook_protection.lock_windows = flag != 0;
            }
            // PASSWORD [MS-XLS 2.4.191]
            RECORD_PASSWORD => {
                if data.len() < 2 {
                    push_warning(&mut out, format!(
                        "truncated PASSWORD record at offset {}",
                        record.offset
                    ));
                    continue;
                }
                let hash = u16::from_le_bytes([data[0], data[1]]);
                out.workbook_protection.password_hash = (hash != 0).then_some(hash);
            }
            // WINDOW1 [MS-XLS 2.4.346]
            RECORD_WINDOW1 => {
                // Payload layout (BIFF8): xWn, yWn, dxWn, dyWn, grbit, iTabCur, ...
                //
                // We treat window metadata as best-effort: malformed/truncated records produce
                // warnings but do not fail globals parsing.

                // Window geometry/state.
                if out.workbook_window.is_none() {
                    if data.len() < 8 {
                        push_warning(&mut out, format!(
                            "WINDOW1 record too short to read window geometry at offset {} (len={})",
                            record.offset,
                            data.len()
                        ));
                    } else {
                        // x/y are signed; dx/dy are unsigned.
                        let x = i16::from_le_bytes([data[0], data[1]]) as i32;
                        let y = i16::from_le_bytes([data[2], data[3]]) as i32;
                        let width = u16::from_le_bytes([data[4], data[5]]) as u32;
                        let height = u16::from_le_bytes([data[6], data[7]]) as u32;

                        let state = if data.len() < 10 {
                            push_warning(&mut out, format!(
                                "WINDOW1 record too short to read window state flags at offset {} (len={})",
                                record.offset,
                                data.len()
                            ));
                            None
                        } else {
                            let grbit = u16::from_le_bytes([data[8], data[9]]);
                            Some(if (grbit & (WINDOW1_GRBIT_HIDDEN | WINDOW1_GRBIT_ICONIC)) != 0 {
                                WorkbookWindowState::Minimized
                            } else if (grbit & WINDOW1_GRBIT_MAXIMIZED) != 0 {
                                WorkbookWindowState::Maximized
                            } else {
                                WorkbookWindowState::Normal
                            })
                        };

                        // Some `.xls` writers emit an all-zero WINDOW1 record. Treat that as
                        // missing metadata so we don't persist a meaningless 0x0 window.
                        let is_empty = x == 0
                            && y == 0
                            && width == 0
                            && height == 0
                            && matches!(state, None | Some(WorkbookWindowState::Normal));

                        if !is_empty {
                            out.workbook_window = Some(WorkbookWindow {
                                x: Some(x),
                                y: Some(y),
                                width: Some(width),
                                height: Some(height),
                                state,
                            });
                        }
                    }
                }

                if data.len() < 12 {
                    push_warning(
                        &mut out,
                        "WINDOW1 record too short to read active tab index".to_string(),
                    );
                } else if out.active_tab_index.is_none() {
                    let tab = u16::from_le_bytes([data[10], data[11]]);
                    out.active_tab_index = Some(tab);
                }
            }
            // 1904 [MS-XLS 2.4.169]
            RECORD_1904 => {
                if data.len() >= 2 {
                    let flag = u16::from_le_bytes([data[0], data[1]]);
                    if flag != 0 {
                        out.date_system = DateSystem::Excel1904;
                    }
                }
            }
            // CALCCOUNT [MS-XLS 2.4.40]
            RECORD_CALCCOUNT => {
                if data.len() < 2 {
                    push_warning(&mut out, format!(
                        "truncated CALCCOUNT record at offset {}",
                        record.offset
                    ));
                    continue;
                }
                let count = u16::from_le_bytes([data[0], data[1]]);
                out.iterative_max_iterations = Some(count as u32);
            }
            // CALCMODE [MS-XLS 2.4.41]
            RECORD_CALCMODE => {
                if data.len() < 2 {
                    push_warning(&mut out, format!(
                        "truncated CALCMODE record at offset {}",
                        record.offset
                    ));
                    continue;
                }
                let raw = u16::from_le_bytes([data[0], data[1]]);
                let mode = match raw {
                    0 => Some(CalculationMode::Manual),
                    1 => Some(CalculationMode::Automatic),
                    2 => Some(CalculationMode::AutomaticNoTable),
                    _ => None,
                };
                if let Some(mode) = mode {
                    out.calculation_mode = Some(mode);
                } else {
                    push_warning(&mut out, format!(
                        "ignored unknown CALCMODE value {raw} at offset {}",
                        record.offset
                    ));
                }
            }
            // PRECISION [MS-XLS 2.4.201]
            RECORD_PRECISION => {
                if data.len() < 2 {
                    push_warning(&mut out, format!(
                        "truncated PRECISION record at offset {}",
                        record.offset
                    ));
                    continue;
                }
                let flag = u16::from_le_bytes([data[0], data[1]]);
                // Non-zero means use full precision; zero means "precision as displayed".
                out.full_precision = Some(flag != 0);
            }
            // DELTA [MS-XLS 2.4.76]
            RECORD_DELTA => {
                if data.len() < 8 {
                    push_warning(&mut out, format!(
                        "truncated DELTA record at offset {}",
                        record.offset
                    ));
                    continue;
                }
                let mut bytes = [0u8; 8];
                bytes.copy_from_slice(&data[..8]);
                out.iterative_max_change = Some(f64::from_le_bytes(bytes));
            }
            // ITERATION [MS-XLS 2.4.118]
            RECORD_ITERATION => {
                if data.len() < 2 {
                    push_warning(&mut out, format!(
                        "truncated ITERATION record at offset {}",
                        record.offset
                    ));
                    continue;
                }
                let flag = u16::from_le_bytes([data[0], data[1]]);
                out.iterative_enabled = Some(flag != 0);
            }
            // FORMAT / FORMAT2 [MS-XLS 2.4.90]
            RECORD_FORMAT_BIFF8 | RECORD_FORMAT2_BIFF5 => {
                match parse_biff_format_record_strict(&record, codepage) {
                    Ok((num_fmt_id, code)) => {
                        out.formats.insert(num_fmt_id, code);
                    }
                    Err(_) if record.is_continued() => {
                        continuation_parse_failed = true;
                        if let Some((num_fmt_id, code)) =
                            parse_biff_format_record_best_effort(&record, codepage)
                        {
                            out.formats.insert(num_fmt_id, code);
                        }
                    }
                    Err(_) => {}
                }
            }
            // PALETTE [MS-XLS 2.4.155]
            RECORD_PALETTE => {
                if data.len() < 2 {
                    push_warning(
                        &mut out,
                        format!("truncated PALETTE record at offset {}", record.offset),
                    );
                } else {
                    let count = u16::from_le_bytes([data[0], data[1]]) as usize;
                    let expected_len = 2usize.saturating_add(count.saturating_mul(4));
                    if data.len() < expected_len {
                        push_warning(&mut out, format!(
                            "truncated PALETTE record at offset {} (ccv={count}, len={}, expected={expected_len})",
                            record.offset,
                            data.len()
                        ));
                    }
                }

                if let Some(palette) = parse_biff_palette_record(data) {
                    out.palette = palette;
                }
            }
            // FONT [MS-XLS 2.4.92]
            RECORD_FONT => match parse_biff_font_record(data, biff, codepage) {
                Ok(font) => out.fonts.push(font),
                Err(err) => push_warning(&mut out, format!("failed to parse FONT record: {err}")),
            },
            // XF [MS-XLS 2.4.353]
            RECORD_XF => {
                if let Ok(xf) = parse_biff_xf_record(data, biff) {
                    out.xfs.push(xf);
                }
            }
            // SAVERECALC [MS-XLS 2.4.248]
            RECORD_SAVERECALC => {
                if data.len() < 2 {
                    push_warning(&mut out, format!(
                        "truncated SAVERECALC record at offset {}",
                        record.offset
                    ));
                    continue;
                }
                let flag = u16::from_le_bytes([data[0], data[1]]);
                out.calculate_before_save = Some(flag != 0);
            }
            // EXTERNSHEET [MS-XLS 2.4.102]
            //
            // BIFF8 payload:
            //   [cXTI: u16]
            //   repeated cXTI times:
            //     [iSupBook: u16][itabFirst: u16][itabLast: u16]
            RECORD_EXTERNSHEET if biff == BiffVersion::Biff8 => {
                let parsed =
                    externsheet::parse_biff8_externsheet_record_data(data, record.offset);
                if parsed.entries.iter().any(|e| e.supbook != 0) {
                    saw_external_supbook = true;
                }
                out.extern_sheets.extend(parsed.entries);
                for warning in parsed.warnings {
                    push_warning(&mut out, warning);
                }
            }
            // SHEETEXT [MS-XLS 2.4.269]
            //
            // This is a BIFF8-only Future Record that may include sheet tab color metadata.
            RECORD_SHEETEXT if biff == BiffVersion::Biff8 => {
                match parse_biff_sheetext_tab_color(data) {
                    Ok(color) => out.sheet_tab_colors.push(color),
                    Err(err) => {
                        push_warning(&mut out, format!(
                            "failed to parse BIFF SHEETEXT record at offset {}: {err}",
                            record.offset
                        ));
                        // Preserve record count so callers can still align subsequent SheetExt
                        // records with workbook sheet order (best-effort).
                        out.sheet_tab_colors.push(None);
                    }
                }
            }
            // EOF terminates the workbook global substream.
            records::RECORD_EOF => {
                saw_eof = true;
                break;
            }
            _ => {}
        }
    }

    if saw_external_supbook {
        push_warning(
            &mut out,
            "workbook contains external SupBook references (EXTERNSHEET); external workbook refs are not yet supported"
                .to_string(),
        );
    }

    if continuation_parse_failed {
        push_warning(
            &mut out,
            "failed to parse one or more continued BIFF FORMAT records; number format codes may be truncated"
                .to_string(),
        );
    }

    if !saw_eof {
        // Some `.xls` files in the wild appear to be truncated or missing the
        // workbook-global EOF record. Treat this as a warning and return any
        // partial data we managed to parse so importers can still recover number
        // formats/date system where possible.
        push_warning(
            &mut out,
            "unexpected end of workbook globals stream (missing EOF)".to_string(),
        );
    }

    if !out.palette.is_empty() {
        apply_palette_to_tab_colors(&mut out.sheet_tab_colors, &out.palette);
    }

    Ok(out)
}

fn apply_palette_to_tab_colors(colors: &mut [Option<TabColor>], palette: &[u32]) {
    for color_opt in colors {
        let Some(color) = color_opt.as_mut() else {
            continue;
        };
        if color.rgb.is_some() {
            continue;
        }
        let Some(indexed) = color.indexed else {
            continue;
        };
        let Ok(idx) = u16::try_from(indexed) else {
            continue;
        };

        // BIFF palette entries correspond to indexed color values 8..=63 by default (56 entries).
        // Map `indexed=N` to palette[N-8] when in range so we can emit a stable ARGB string for
        // `.xls` -> `.xlsx` conversion. Fall back to Excel's standard indexed table when the
        // palette record is missing/truncated.
        let argb = idx
            .checked_sub(8)
            .and_then(|v| palette.get(v as usize).copied())
            .or_else(|| indexed_color_argb(idx));
        let Some(argb) = argb else {
            continue;
        };
        color.rgb = Some(format!("{argb:08X}"));
        // Prefer RGB when available; XLSX tabColor `indexed` depends on a compatible indexedColors
        // table, which we don't currently synthesize for legacy `.xls` conversion.
        color.indexed = None;
    }
}

/// Parse a BIFF8 `SHEETEXT` record and return the optional sheet tab color.
///
/// The `SHEETEXT` record is a Future Record Type (FRT) record and begins with an `FrtHeader`
/// (8 bytes). The remainder of the record includes various worksheet-level flags; the tab color is
/// stored as an `XColor` structure at the end of the record in files produced by Excel.
///
/// This parser is best-effort and only extracts `rgb`/`indexed` when they can be inferred.
fn parse_biff_sheetext_tab_color(data: &[u8]) -> Result<Option<TabColor>, String> {
    // `FrtHeader` is 8 bytes: rt (u16), grbitFrt (u16), reserved (u32).
    if data.len() < 8 {
        return Err(format!("SHEETEXT record too short (len={})", data.len()));
    }

    let payload = &data[8..];

    // Best-effort: In BIFF8 files produced by Excel, the tab color is stored as an `XColor`
    // structure (8 bytes) at the end of the record. Prefer that representation because it can
    // include RGB values directly.
    if payload.len() >= 8 {
        let xcolor = &payload[payload.len() - 8..];
        if let Some(color) = parse_biff_xcolor_tab_color(xcolor) {
            return Ok(Some(color));
        }
        // Fall through to additional heuristics in case we mis-detected the `XColor` layout.
    }

    // Fallback: treat the final 2 bytes as an indexed color (`ICV`).
    if payload.len() >= 2 {
        let idx_bytes = &payload[payload.len() - 2..];
        let idx = u16::from_le_bytes([idx_bytes[0], idx_bytes[1]]);
        if let Some(color) = icv_to_tab_color_indexed(idx) {
            return Ok(Some(color));
        }
    }

    // Fallback: treat the final 4 bytes as a `LongRGB`/`COLORREF` if present and non-zero.
    if payload.len() >= 4 {
        let rgb = &payload[payload.len() - 4..];
        if rgb[..3].iter().any(|b| *b != 0) {
            return Ok(Some(TabColor::rgb(long_rgb_to_argb_hex(rgb))));
        }
    }

    Ok(None)
}

fn parse_biff_xcolor_tab_color(data: &[u8]) -> Option<TabColor> {
    // Best-effort `XColor` parsing:
    //
    // `XColor` is used throughout BIFF8 to represent colors and can be specified in different
    // modes (indexed vs RGB). For sheet tab colors, Excel stores an `XColor` payload at the end of
    // the `SHEETEXT` record.
    //
    // Layout (best-effort, as used by Excel):
    // - [0..2]  xclrType (u16)
    // - [2..4]  icv / index (u16)
    // - [4..8]  `LongRGB` / `COLORREF` (4 bytes)
    if data.len() < 8 {
        return None;
    }

    let xclr_type = u16::from_le_bytes([data[0], data[1]]);
    let idx = u16::from_le_bytes([data[2], data[3]]);
    let rgb_bytes = &data[4..8];

    // Common `XColorType` values (best-effort):
    // 0 = automatic/none
    // 1 = indexed (legacy palette)
    // 2 = RGB
    // 3 = theme
    match xclr_type {
        0 => None,
        1 => icv_to_tab_color_indexed(idx),
        2 => Some(TabColor::rgb(long_rgb_to_argb_hex(rgb_bytes))),
        3 => {
            // Best-effort: treat the index field as an OOXML theme index.
            let mut out = TabColor::default();
            out.theme = Some(idx as u32);
            Some(out)
        }
        // Be conservative for unknown variants to avoid introducing incorrect tab colors.
        _ => None,
    }
}

fn icv_to_tab_color_indexed(idx: u16) -> Option<TabColor> {
    // BIFF `ICV` uses 0x7FFF for "automatic". Treat that (and 0) as "no tab color".
    const ICV_AUTOMATIC: u16 = 0x7FFF;
    if idx == 0 || idx == ICV_AUTOMATIC {
        return None;
    }
    let mut out = TabColor::default();
    out.indexed = Some(idx as u32);
    Some(out)
}

fn long_rgb_to_argb_hex(rgb: &[u8]) -> String {
    // BIFF `LongRGB`/`COLORREF` is stored as:
    // - red (u8)
    // - green (u8)
    // - blue (u8)
    // - reserved / alpha (u8, typically 0)
    let r = rgb.get(0).copied().unwrap_or(0);
    let g = rgb.get(1).copied().unwrap_or(0);
    let b = rgb.get(2).copied().unwrap_or(0);
    let a = rgb.get(3).copied().unwrap_or(0);

    let alpha = if a == 0 { 0xFF } else { a };
    format!("{alpha:02X}{r:02X}{g:02X}{b:02X}")
}

fn long_rgb_to_argb(rgb: &[u8]) -> u32 {
    let r = rgb.get(0).copied().unwrap_or(0);
    let g = rgb.get(1).copied().unwrap_or(0);
    let b = rgb.get(2).copied().unwrap_or(0);
    let a = rgb.get(3).copied().unwrap_or(0);
    let alpha = if a == 0 { 0xFF } else { a };
    ((alpha as u32) << 24) | ((r as u32) << 16) | ((g as u32) << 8) | (b as u32)
}

fn workbook_globals_allows_continuation_biff5(record_id: u16) -> bool {
    // FORMAT2 [MS-XLS 2.4.90]
    record_id == RECORD_FORMAT2_BIFF5
}

fn workbook_globals_allows_continuation_biff8(record_id: u16) -> bool {
    // Records in the workbook globals substream that are known to allow `CONTINUE` records.
    //
    // NOTE: The `LogicalBiffRecordIter` preserves fragment boundaries so parsers can handle BIFF8
    // continued strings (which inject an option-flags byte at the start of a continuation
    // fragment).
    record_id == RECORD_FORMAT_BIFF8 || record_id == RECORD_EXTERNSHEET || record_id == RECORD_NAME
}

fn parse_biff_xf_record(data: &[u8], biff: BiffVersion) -> Result<BiffXf, String> {
    if data.len() < 4 {
        return Err("XF record too short".to_string());
    }

    let font_idx = u16::from_le_bytes([data[0], data[1]]);
    let num_fmt_id = u16::from_le_bytes([data[2], data[3]]);

    let mut xf = BiffXf {
        font_idx,
        num_fmt_id,
        kind: None,
        parent_xf: None,
        apply: BiffXfApplyFlags::default(),
        protection: BiffXfProtection::default(),
        alignment: BiffXfAlignment::default(),
        border: BiffXfBorder::default(),
        fill: BiffXfFill::default(),
    };

    // Optional: in BIFF5/8 this is part of the "type/protection" flags field and bit 2 is `fStyle`.
    if data.len() >= 6 {
        let flags = u16::from_le_bytes([data[4], data[5]]);
        xf.kind = Some(if (flags & XF_FLAG_STYLE) != 0 {
            BiffXfKind::Style
        } else {
            BiffXfKind::Cell
        });
        xf.protection.locked = (flags & 0x0001) != 0;
        xf.protection.hidden = (flags & 0x0002) != 0;
        xf.parent_xf = Some(flags >> 4);
    } else {
        xf.protection.locked = true;
        xf.protection.hidden = false;
    }

    match biff {
        BiffVersion::Biff8 => {
            if data.len() >= 20 {
                parse_biff8_xf_payload(data, &mut xf);
            }
        }
        BiffVersion::Biff5 => {
            // BIFF5 XF records are 16 bytes; the field packing differs from BIFF8.
            // Parse best-effort using a BIFF8-compatible layout when enough bytes are present.
            if data.len() >= 16 {
                parse_biff5_xf_payload_best_effort(data, &mut xf);
            }
        }
    }

    Ok(xf)
}

fn parse_biff8_xf_payload(data: &[u8], xf: &mut BiffXf) {
    // BIFF8 XF record layout (20 bytes), best-effort:
    // [0..2]  ifnt
    // [2..4]  ifmt
    // [4..6]  type/protection/parent flags
    // [6]     alignment
    // [7]     rotation
    // [8]     text properties (indent, etc.)
    // [9]     attribute flags (apply bits)
    // [10..14] border1
    // [14..18] border2
    // [18..20] pattern (fill colors)
    let alignment = data[6];
    let rotation = data[7];
    let text_props = data[8];
    let used_attr = data[9];

    xf.alignment.horizontal = parse_biff_horizontal_alignment(alignment & 0x07);
    xf.alignment.wrap_text = (alignment & 0x08) != 0;
    xf.alignment.vertical = parse_biff_vertical_alignment((alignment >> 4) & 0x07);

    xf.alignment.rotation = (rotation != 0).then_some(rotation as i16);

    let indent = (text_props & 0x0F) as u16;
    xf.alignment.indent = (indent != 0).then_some(indent);

    // Attribute flags (apply bits) [MS-XLS 2.4.353]: best-effort mapping.
    //
    // Some BIFF files appear to leave this field as 0. Treat 0 as "apply all" so we preserve
    // formatting rather than dropping it.
    if used_attr == 0 {
        xf.apply.num_fmt = true;
        xf.apply.font = true;
        xf.apply.alignment = true;
        xf.apply.border = true;
        xf.apply.fill = true;
        xf.apply.protection = true;
    } else {
        xf.apply.num_fmt = (used_attr & 0x01) != 0;
        xf.apply.font = (used_attr & 0x02) != 0;
        xf.apply.alignment = (used_attr & 0x04) != 0;
        xf.apply.border = (used_attr & 0x08) != 0;
        xf.apply.fill = (used_attr & 0x10) != 0;
        xf.apply.protection = (used_attr & 0x20) != 0;
    }

    let border1 = u32::from_le_bytes([data[10], data[11], data[12], data[13]]);
    let border2 = u32::from_le_bytes([data[14], data[15], data[16], data[17]]);
    let pattern = u16::from_le_bytes([data[18], data[19]]);

    xf.border.left.style = (border1 & 0xF) as u8;
    xf.border.right.style = ((border1 >> 4) & 0xF) as u8;
    xf.border.top.style = ((border1 >> 8) & 0xF) as u8;
    xf.border.bottom.style = ((border1 >> 12) & 0xF) as u8;

    xf.border.left.color_idx = ((border1 >> 16) & 0x7F) as u16;
    xf.border.right.color_idx = ((border1 >> 23) & 0x7F) as u16;

    xf.border.diagonal_down = ((border1 >> 30) & 0x1) != 0;
    xf.border.diagonal_up = ((border1 >> 31) & 0x1) != 0;

    xf.border.top.color_idx = (border2 & 0x7F) as u16;
    xf.border.bottom.color_idx = ((border2 >> 7) & 0x7F) as u16;
    xf.border.diagonal.color_idx = ((border2 >> 14) & 0x7F) as u16;
    xf.border.diagonal.style = ((border2 >> 21) & 0x0F) as u8;

    xf.fill.pattern = ((border2 >> 25) & 0x3F) as u8;
    xf.fill.fg_color_idx = (pattern & 0x7F) as u16;
    xf.fill.bg_color_idx = ((pattern >> 7) & 0x7F) as u16;
}

fn parse_biff5_xf_payload_best_effort(data: &[u8], xf: &mut BiffXf) {
    // BIFF5 XF records are 16 bytes. Parse a subset by treating:
    // [6] alignment, [7] rotation, [8] text_props, [9] attr flags, [10..14] border1, [14..16] pattern.
    let alignment = data[6];
    let rotation = data[7];
    let text_props = data[8];
    let used_attr = data[9];

    xf.alignment.horizontal = parse_biff_horizontal_alignment(alignment & 0x07);
    xf.alignment.wrap_text = (alignment & 0x08) != 0;
    xf.alignment.vertical = parse_biff_vertical_alignment((alignment >> 4) & 0x07);
    xf.alignment.rotation = (rotation != 0).then_some(rotation as i16);

    let indent = (text_props & 0x0F) as u16;
    xf.alignment.indent = (indent != 0).then_some(indent);

    if used_attr == 0 {
        xf.apply.num_fmt = true;
        xf.apply.font = true;
        xf.apply.alignment = true;
        xf.apply.border = true;
        xf.apply.fill = true;
        xf.apply.protection = true;
    } else {
        xf.apply.num_fmt = (used_attr & 0x01) != 0;
        xf.apply.font = (used_attr & 0x02) != 0;
        xf.apply.alignment = (used_attr & 0x04) != 0;
        xf.apply.border = (used_attr & 0x08) != 0;
        xf.apply.fill = (used_attr & 0x10) != 0;
        xf.apply.protection = (used_attr & 0x20) != 0;
    }

    let border1 = u32::from_le_bytes([data[10], data[11], data[12], data[13]]);
    let pattern = u16::from_le_bytes([data[14], data[15]]);

    xf.border.left.style = (border1 & 0xF) as u8;
    xf.border.right.style = ((border1 >> 4) & 0xF) as u8;
    xf.border.top.style = ((border1 >> 8) & 0xF) as u8;
    xf.border.bottom.style = ((border1 >> 12) & 0xF) as u8;

    xf.border.left.color_idx = ((border1 >> 16) & 0x7F) as u16;
    xf.border.right.color_idx = ((border1 >> 23) & 0x7F) as u16;
    xf.border.diagonal_down = ((border1 >> 30) & 0x1) != 0;
    xf.border.diagonal_up = ((border1 >> 31) & 0x1) != 0;

    xf.fill.fg_color_idx = (pattern & 0x7F) as u16;
    xf.fill.bg_color_idx = ((pattern >> 7) & 0x7F) as u16;
}

fn parse_biff_horizontal_alignment(code: u8) -> Option<HorizontalAlignment> {
    match code {
        0 => None, // General
        1 => Some(HorizontalAlignment::Left),
        2 => Some(HorizontalAlignment::Center),
        3 => Some(HorizontalAlignment::Right),
        4 => Some(HorizontalAlignment::Fill),
        5 => Some(HorizontalAlignment::Justify),
        _ => None,
    }
}

fn parse_biff_vertical_alignment(code: u8) -> Option<VerticalAlignment> {
    match code {
        0 => Some(VerticalAlignment::Top),
        1 => Some(VerticalAlignment::Center),
        // Bottom is the default alignment in Excel; treat it as "unset" for stable style tables.
        2 => None,
        _ => None,
    }
}

fn parse_biff_palette_record(data: &[u8]) -> Option<Vec<u32>> {
    if data.len() < 2 {
        return None;
    }
    let count = u16::from_le_bytes([data[0], data[1]]) as usize;
    // `PALETTE` payload is `[ccv: u16]` followed by repeated `LongRGB` entries (4 bytes each).
    //
    // Cap the allocation based on the actual number of bytes available to avoid allocating large
    // vectors for corrupt/truncated records where `ccv` does not match the payload length.
    let max_colors = (data.len().saturating_sub(2)) / 4;
    let count = count.min(max_colors);

    let mut out = Vec::new();
    let _ = out.try_reserve_exact(count);
    let mut offset = 2usize;
    for _ in 0..count {
        let end = match offset.checked_add(4) {
            Some(v) => v,
            None => break,
        };
        let Some(rgb) = data.get(offset..end) else {
            break;
        };
        out.push(long_rgb_to_argb(rgb));
        offset = end;
    }
    Some(out)
}

fn parse_biff_font_record(data: &[u8], biff: BiffVersion, codepage: u16) -> Result<BiffFont, String> {
    if data.len() < 14 {
        return Err("FONT record too short".to_string());
    }

    let height_twips = u16::from_le_bytes([data[0], data[1]]);
    let grbit = u16::from_le_bytes([data[2], data[3]]);
    let color_idx = u16::from_le_bytes([data[4], data[5]]);
    let weight = u16::from_le_bytes([data[6], data[7]]);
    // let escapement = u16::from_le_bytes([data[8], data[9]]);
    let underline = data[10];
    // let family = data[11];
    // let charset = data[12];
    // let reserved = data[13];

    let (mut name, _) = strings::parse_biff_short_string(&data[14..], biff, codepage)?;
    // Excel stores some strings with embedded NUL bytes; strip them so font names round-trip
    // deterministically and match Excel’s visible semantics.
    strip_embedded_nuls(&mut name);

    let size_100pt = height_twips.saturating_mul(5);
    let bold = weight >= 700;
    let italic = (grbit & 0x0002) != 0;
    let strike = (grbit & 0x0008) != 0;

    Ok(BiffFont {
        name,
        size_100pt,
        bold,
        italic,
        underline: underline != 0,
        strike,
        color_idx,
    })
}

fn parse_biff_format_record_strict(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
) -> Result<(u16, String), String> {
    let record_id = record.record_id;
    let data = record.data.as_ref();
    if data.len() < 2 {
        return Err("FORMAT record too short".to_string());
    }

    let num_fmt_id = u16::from_le_bytes([data[0], data[1]]);
    let rest = &data[2..];

    let mut code = match record_id {
        // BIFF8 FORMAT uses `XLUnicodeString` (16-bit length) and may be split
        // across one or more `CONTINUE` records.
        RECORD_FORMAT_BIFF8 => {
            if record.is_continued() {
                let fragments: Vec<&[u8]> = record.fragments().collect();
                strings::parse_biff8_unicode_string_continued(&fragments, 2, codepage)?
            } else {
                strings::parse_biff8_unicode_string(rest, codepage)?.0
            }
        }
        // BIFF5 FORMAT2 uses a short ANSI string (8-bit length).
        RECORD_FORMAT2_BIFF5 => strings::parse_biff5_short_string(rest, codepage)?.0,
        _ => return Err(format!("unexpected FORMAT record id 0x{record_id:04X}")),
    };

    // Excel stores some strings with embedded NUL bytes; strip them for stable formatting.
    strip_embedded_nuls(&mut code);
    Ok((num_fmt_id, code))
}

fn parse_biff_format_record_best_effort(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
) -> Option<(u16, String)> {
    let first = record.first_fragment();
    if first.len() < 2 {
        return None;
    }
    let num_fmt_id = u16::from_le_bytes([first[0], first[1]]);
    let rest = first.get(2..).unwrap_or_default();

    let mut code = match record.record_id {
        RECORD_FORMAT_BIFF8 => strings::parse_biff8_unicode_string_best_effort(rest, codepage)?,
        RECORD_FORMAT2_BIFF5 => strings::parse_biff5_short_string_best_effort(rest, codepage)?,
        _ => return None,
    };
    strip_embedded_nuls(&mut code);
    Some((num_fmt_id, code))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u16).to_le_bytes());
        out.extend_from_slice(data);
        out
    }

    #[test]
    fn boundsheet_scan_stops_on_truncated_record() {
        // BoundSheet8 with a compressed 8-bit name (fHighByte=0).
        let mut bs_payload = Vec::new();
        bs_payload.extend_from_slice(&0x1234u32.to_le_bytes()); // sheet offset
        bs_payload.extend_from_slice(&[0, 0]); // visibility/type
        bs_payload.push(1); // cch
        bs_payload.push(0); // flags (compressed)
        bs_payload.push(b'A');

        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x0001u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[1, 2]); // missing 2 bytes

        let stream = [record(RECORD_BOUNDSHEET, &bs_payload), truncated].concat();
        let codepage = parse_biff_codepage(&stream);
        let sheets = parse_biff_bound_sheets(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(
            sheets,
            vec![BoundSheetInfo {
                name: "A".to_string(),
                offset: 0x1234,
                hs_state: 0,
                sheet_visibility: Some(SheetVisibility::Visible),
                dt: 0,
                sheet_type: Some(BoundSheetType::Worksheet),
            }]
        );
    }

    #[test]
    fn boundsheet_strips_embedded_nuls() {
        // BoundSheet8 with a compressed 8-bit name containing an embedded NUL byte.
        let mut bs_payload = Vec::new();
        bs_payload.extend_from_slice(&0x1234u32.to_le_bytes()); // sheet offset
        bs_payload.extend_from_slice(&[0, 0]); // visibility/type
        bs_payload.push(3); // cch
        bs_payload.push(0); // flags (compressed)
        bs_payload.extend_from_slice(&[b'A', 0x00, b'B']);

        let stream = [
            record(RECORD_BOUNDSHEET, &bs_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let codepage = parse_biff_codepage(&stream);
        let sheets = parse_biff_bound_sheets(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(
            sheets,
            vec![BoundSheetInfo {
                name: "AB".to_string(),
                offset: 0x1234,
                hs_state: 0,
                sheet_visibility: Some(SheetVisibility::Visible),
                dt: 0,
                sheet_type: Some(BoundSheetType::Worksheet),
            }]
        );
    }

    #[test]
    fn boundsheet_parses_visibility_and_type() {
        // BoundSheet8 with:
        // - hsState = 0x02 (very hidden)
        // - dt = 0x02 (chart sheet)
        let mut bs_payload = Vec::new();
        bs_payload.extend_from_slice(&0x1234u32.to_le_bytes()); // sheet offset
        bs_payload.push(0x02); // hsState
        bs_payload.push(0x02); // dt
        bs_payload.push(1); // cch
        bs_payload.push(0); // flags (compressed)
        bs_payload.push(b'A');

        let stream = [
            record(RECORD_BOUNDSHEET, &bs_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let codepage = parse_biff_codepage(&stream);
        let sheets = parse_biff_bound_sheets(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(
            sheets,
            vec![BoundSheetInfo {
                name: "A".to_string(),
                offset: 0x1234,
                hs_state: 0x02,
                sheet_visibility: Some(SheetVisibility::VeryHidden),
                dt: 0x02,
                sheet_type: Some(BoundSheetType::Chart),
            }]
        );
    }

    #[test]
    fn font_strips_embedded_nuls() {
        // FONT record with a compressed 8-bit name containing an embedded NUL byte.
        //
        // BIFF8 `ShortXLUnicodeString`: [cch=3][flags=0][A][\\0][B]
        let mut payload = Vec::new();
        payload.extend_from_slice(&200u16.to_le_bytes()); // height_twips (10pt)
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        payload.extend_from_slice(&0u16.to_le_bytes()); // icv
        payload.extend_from_slice(&400u16.to_le_bytes()); // weight
        payload.extend_from_slice(&0u16.to_le_bytes()); // sss
        payload.push(0); // underline
        payload.push(0); // family
        payload.push(0); // charset
        payload.push(0); // reserved
        payload.push(3); // cch
        payload.push(0); // flags (compressed)
        payload.extend_from_slice(&[b'A', 0x00, b'B']);

        let font = parse_biff_font_record(&payload, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(font.name, "AB");
    }

    #[test]
    fn decodes_boundsheet_names_using_codepage() {
        // CODEPAGE=1251 (Windows Cyrillic).
        let r_codepage = record(RECORD_CODEPAGE, &1251u16.to_le_bytes());

        // BoundSheet8 with a compressed 8-bit name (fHighByte=0).
        let mut bs_payload = Vec::new();
        bs_payload.extend_from_slice(&0x1234u32.to_le_bytes()); // sheet offset
        bs_payload.extend_from_slice(&[0, 0]); // visibility/type
        bs_payload.push(1); // cch
        bs_payload.push(0); // flags (compressed)
        bs_payload.push(0x80); // "Ђ" in cp1251
        let r_bs = record(RECORD_BOUNDSHEET, &bs_payload);

        let stream = [r_codepage, r_bs, record(records::RECORD_EOF, &[])].concat();
        let codepage = parse_biff_codepage(&stream);
        let sheets = parse_biff_bound_sheets(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(
            sheets,
            vec![BoundSheetInfo {
                name: "Ђ".to_string(),
                offset: 0x1234,
                hs_state: 0,
                sheet_visibility: Some(SheetVisibility::Visible),
                dt: 0,
                sheet_type: Some(BoundSheetType::Worksheet),
            }]
        );
    }

    #[test]
    fn decodes_boundsheet_names_using_codepage_1252_currency_symbol() {
        // CODEPAGE=1252 (Windows Western).
        let r_codepage = record(RECORD_CODEPAGE, &1252u16.to_le_bytes());

        // BoundSheet8 with a compressed 8-bit name (fHighByte=0): 0xA3 => '£' in cp1252.
        let mut bs_payload = Vec::new();
        bs_payload.extend_from_slice(&0x1234u32.to_le_bytes()); // sheet offset
        bs_payload.extend_from_slice(&[0, 0]); // visibility/type
        bs_payload.push(1); // cch
        bs_payload.push(0); // flags (compressed)
        bs_payload.push(0xA3); // "£" in cp1252
        let r_bs = record(RECORD_BOUNDSHEET, &bs_payload);

        let stream = [r_codepage, r_bs, record(records::RECORD_EOF, &[])].concat();
        let codepage = parse_biff_codepage(&stream);
        let sheets = parse_biff_bound_sheets(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(
            sheets,
            vec![BoundSheetInfo {
                name: "£".to_string(),
                offset: 0x1234,
                hs_state: 0,
                sheet_visibility: Some(SheetVisibility::Visible),
                dt: 0,
                sheet_type: Some(BoundSheetType::Worksheet),
            }]
        );
    }

    #[test]
    fn codepage_defaults_to_1252_when_missing() {
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        assert_eq!(parse_biff_codepage(&stream), 1252);
    }

    #[test]
    fn codepage_scan_stops_at_next_bof() {
        // CODEPAGE after the next BOF should be ignored.
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_CODEPAGE, &1251u16.to_le_bytes()),
        ]
        .concat();
        assert_eq!(parse_biff_codepage(&stream), 1252);
    }

    #[test]
    fn codepage_uses_first_record() {
        let stream = [
            record(RECORD_CODEPAGE, &1251u16.to_le_bytes()),
            record(RECORD_CODEPAGE, &1252u16.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        assert_eq!(parse_biff_codepage(&stream), 1251);
    }

    #[test]
    fn boundsheet_scan_stops_at_next_bof_without_eof() {
        // CODEPAGE=1251 (Windows Cyrillic).
        let r_codepage = record(RECORD_CODEPAGE, &1251u16.to_le_bytes());

        // BoundSheet8 with a compressed 8-bit name (fHighByte=0).
        let mut bs_payload = Vec::new();
        bs_payload.extend_from_slice(&0x1234u32.to_le_bytes()); // sheet offset
        bs_payload.extend_from_slice(&[0, 0]); // visibility/type
        bs_payload.push(1); // cch
        bs_payload.push(0); // flags (compressed)
        bs_payload.push(0x80); // "Ђ" in cp1251
        let r_bs = record(RECORD_BOUNDSHEET, &bs_payload);

        // BOF for the next substream (worksheet).
        let r_sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // No EOF record; should still stop at the worksheet BOF.
        let stream = [r_codepage, r_bs, r_sheet_bof].concat();
        let codepage = parse_biff_codepage(&stream);
        let sheets = parse_biff_bound_sheets(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(
            sheets,
            vec![BoundSheetInfo {
                name: "Ђ".to_string(),
                offset: 0x1234,
                hs_state: 0,
                sheet_visibility: Some(SheetVisibility::Visible),
                dt: 0,
                sheet_type: Some(BoundSheetType::Worksheet),
            }]
        );
    }

    #[test]
    fn globals_scan_stops_at_next_bof_without_eof() {
        let r_bof_globals = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);
        // CODEPAGE=1251 (Windows Cyrillic).
        let r_codepage = record(RECORD_CODEPAGE, &1251u16.to_le_bytes());

        // FORMAT id=200, code = byte 0x80 in cp1251 => "Ђ".
        let mut fmt_payload = Vec::new();
        fmt_payload.extend_from_slice(&200u16.to_le_bytes());
        fmt_payload.extend_from_slice(&1u16.to_le_bytes()); // cch
        fmt_payload.push(0); // flags (compressed)
        fmt_payload.push(0x80); // "Ђ" in cp1251
        let r_fmt = record(RECORD_FORMAT_BIFF8, &fmt_payload);

        let mut xf_payload = vec![0u8; 20];
        xf_payload[2..4].copy_from_slice(&200u16.to_le_bytes());
        let r_xf = record(RECORD_XF, &xf_payload);

        // BOF for the next substream (worksheet).
        let r_sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // A 1904 record and another CODEPAGE after the worksheet BOF should be ignored.
        let r_1904_after = record(RECORD_1904, &[1, 0]);
        let r_codepage_after = record(RECORD_CODEPAGE, &1252u16.to_le_bytes());

        // No EOF for globals; parser should stop at the worksheet BOF.
        let stream = [
            r_bof_globals,
            r_codepage,
            r_fmt,
            r_xf,
            r_sheet_bof,
            r_1904_after,
            r_codepage_after,
        ]
        .concat();

        let codepage = parse_biff_codepage(&stream);
        let globals =
            parse_biff_workbook_globals(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(globals.date_system, DateSystem::Excel1900);
        assert_eq!(globals.xf_count(), 1);
        assert_eq!(globals.resolve_number_format_code(0).as_deref(), Some("Ђ"));
    }

    #[test]
    fn globals_missing_eof_returns_partial_with_warning() {
        let r_bof_globals = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);
        let r_1904 = record(RECORD_1904, &[1, 0]);

        let mut xf_payload = vec![0u8; 20];
        xf_payload[2..4].copy_from_slice(&14u16.to_le_bytes()); // built-in date format
        let r_xf = record(RECORD_XF, &xf_payload);

        // No EOF record and no subsequent BOF; parser should return partial globals with a warning.
        let stream = [r_bof_globals, r_1904, r_xf].concat();
        let codepage = parse_biff_codepage(&stream);
        let globals =
            parse_biff_workbook_globals(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(globals.date_system, DateSystem::Excel1904);
        assert_eq!(globals.xf_count(), 1);
        assert!(
            globals.warnings.iter().any(|w| w.contains("missing EOF")),
            "expected missing-EOF warning, got {:?}",
            globals.warnings
        );
        assert_eq!(
            globals.resolve_number_format_code(0).as_deref(),
            Some("m/d/yyyy")
        );
    }

    #[test]
    fn globals_scan_stops_on_malformed_record_and_warns() {
        let r_bof_globals = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);
        let r_1904 = record(RECORD_1904, &[1, 0]);

        // Truncated record: declares 4 bytes but only provides 2.
        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x1234u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[1, 2]);

        let stream = [r_bof_globals, r_1904, truncated].concat();
        let codepage = parse_biff_codepage(&stream);
        let globals =
            parse_biff_workbook_globals(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(globals.date_system, DateSystem::Excel1904);
        assert!(
            globals
                .warnings
                .iter()
                .any(|w| w.contains("malformed BIFF record")),
            "expected malformed-record warning, got {:?}",
            globals.warnings
        );
        assert!(
            globals.warnings.iter().any(|w| w.contains("missing EOF")),
            "expected missing-EOF warning, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn globals_warnings_are_capped() {
        // Build a workbook-global stream containing many malformed records that each produce a
        // warning. Without a cap, this would grow `BiffWorkbookGlobals.warnings` without bound.
        let mut stream = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // PROTECT records with a payload shorter than 2 bytes produce a warning.
        for _ in 0..(MAX_GLOBAL_WARNINGS + 50) {
            stream.extend_from_slice(&record(RECORD_PROTECT, &[]));
        }

        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let globals =
            parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(globals.warnings.len(), MAX_GLOBAL_WARNINGS + 1);
        assert_eq!(
            globals.warnings.last().map(String::as_str),
            Some(GLOBAL_WARNINGS_SUPPRESSED_MSG)
        );
    }

    #[test]
    fn parses_workbook_protection_records() {
        let stream = [
            record(RECORD_PROTECT, &1u16.to_le_bytes()),
            record(RECORD_WINDOWPROTECT, &1u16.to_le_bytes()),
            record(RECORD_PASSWORD, &0x83AFu16.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(globals.workbook_protection.lock_structure, true);
        assert_eq!(globals.workbook_protection.lock_windows, true);
        assert_eq!(globals.workbook_protection.password_hash, Some(0x83AF));
        assert!(
            globals.warnings.is_empty(),
            "expected no warnings, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn workbook_protection_password_hash_zero_is_none() {
        let stream = [
            record(RECORD_PROTECT, &1u16.to_le_bytes()),
            record(RECORD_WINDOWPROTECT, &1u16.to_le_bytes()),
            // Hash value 0 indicates "no password" in Excel's legacy protection scheme.
            record(RECORD_PASSWORD, &0u16.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(globals.workbook_protection.lock_structure, true);
        assert_eq!(globals.workbook_protection.lock_windows, true);
        assert_eq!(globals.workbook_protection.password_hash, None);
        assert!(
            globals.warnings.is_empty(),
            "expected no warnings, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn globals_does_not_flag_filepass_without_bof() {
        // A stream that contains a FILEPASS record but does not start with a BOF record should not
        // be treated as an encrypted workbook stream.
        let stream = [
            record(0x0001, &[0xAA]),
            record(records::RECORD_FILEPASS, &[0x00, 0x00]),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert!(!globals.is_encrypted);
    }

    #[test]
    fn globals_flags_filepass_when_bof_present() {
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(records::RECORD_FILEPASS, &[0x00, 0x00]),
        ]
        .concat();

        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert!(globals.is_encrypted);
    }

    #[test]
    fn window1_all_zero_is_ignored_for_window_geometry() {
        // Some `.xls` writers emit an all-zero WINDOW1 record. We treat this as missing window
        // geometry metadata so we don't persist a meaningless 0x0 window.
        let stream = [
            record(RECORD_WINDOW1, &[0u8; 18]),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(globals.workbook_window, None);
        assert!(
            globals.warnings.is_empty(),
            "expected no warnings, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn window1_warns_on_truncated_state_flags_but_imports_geometry() {
        // WINDOW1 with only the first 8 bytes (x/y/width/height). The parser should warn about the
        // missing grbit/state and active-tab fields, but still import geometry.
        let mut payload = Vec::new();
        payload.extend_from_slice(&10i16.to_le_bytes()); // xWn
        payload.extend_from_slice(&20i16.to_le_bytes()); // yWn
        payload.extend_from_slice(&30u16.to_le_bytes()); // dxWn
        payload.extend_from_slice(&40u16.to_le_bytes()); // dyWn

        let stream = [
            record(RECORD_WINDOW1, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            globals.workbook_window,
            Some(WorkbookWindow {
                x: Some(10),
                y: Some(20),
                width: Some(30),
                height: Some(40),
                state: None
            })
        );
        assert!(
            globals
                .warnings
                .iter()
                .any(|w| w.contains("WINDOW1 record too short to read window state flags")),
            "expected WINDOW1-state warning, got {:?}",
            globals.warnings
        );
        assert!(
            globals
                .warnings
                .iter()
                .any(|w| w.contains("WINDOW1 record too short to read active tab index")),
            "expected WINDOW1-active-tab warning, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn window1_hidden_maps_to_minimized_state() {
        // fHidden corresponds to Excel's View -> Window -> Hide. We treat it as a minimized window
        // state because `formula_model` does not currently distinguish hidden vs minimized.
        let mut payload = [0u8; 18];
        payload[0..2].copy_from_slice(&111i16.to_le_bytes()); // xWn
        payload[2..4].copy_from_slice(&222i16.to_le_bytes()); // yWn
        payload[4..6].copy_from_slice(&333u16.to_le_bytes()); // dxWn
        payload[6..8].copy_from_slice(&444u16.to_le_bytes()); // dyWn
        payload[8..10].copy_from_slice(&WINDOW1_GRBIT_HIDDEN.to_le_bytes()); // grbit

        let stream = [
            record(RECORD_WINDOW1, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            globals.workbook_window,
            Some(WorkbookWindow {
                x: Some(111),
                y: Some(222),
                width: Some(333),
                height: Some(444),
                state: Some(WorkbookWindowState::Minimized)
            })
        );
        assert!(
            globals.warnings.is_empty(),
            "expected no warnings, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn window1_maximized_maps_to_maximized_state() {
        let mut payload = [0u8; 18];
        payload[0..2].copy_from_slice(&100i16.to_le_bytes()); // xWn
        payload[2..4].copy_from_slice(&200i16.to_le_bytes()); // yWn
        payload[4..6].copy_from_slice(&300u16.to_le_bytes()); // dxWn
        payload[6..8].copy_from_slice(&400u16.to_le_bytes()); // dyWn
        payload[8..10].copy_from_slice(&WINDOW1_GRBIT_MAXIMIZED.to_le_bytes()); // grbit

        let stream = [
            record(RECORD_WINDOW1, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            globals.workbook_window,
            Some(WorkbookWindow {
                x: Some(100),
                y: Some(200),
                width: Some(300),
                height: Some(400),
                state: Some(WorkbookWindowState::Maximized)
            })
        );
        assert!(
            globals.warnings.is_empty(),
            "expected no warnings, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn window1_iconic_maps_to_minimized_state() {
        let mut payload = [0u8; 18];
        payload[0..2].copy_from_slice(&10i16.to_le_bytes()); // xWn
        payload[2..4].copy_from_slice(&20i16.to_le_bytes()); // yWn
        payload[4..6].copy_from_slice(&30u16.to_le_bytes()); // dxWn
        payload[6..8].copy_from_slice(&40u16.to_le_bytes()); // dyWn
        payload[8..10].copy_from_slice(&WINDOW1_GRBIT_ICONIC.to_le_bytes()); // grbit

        let stream = [
            record(RECORD_WINDOW1, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            globals.workbook_window,
            Some(WorkbookWindow {
                x: Some(10),
                y: Some(20),
                width: Some(30),
                height: Some(40),
                state: Some(WorkbookWindowState::Minimized)
            })
        );
        assert!(
            globals.warnings.is_empty(),
            "expected no warnings, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn window1_iconic_overrides_maximized() {
        // If both iconic and maximized bits are set, treat the window as minimized. This matches the
        // parser precedence and is the safest fallback when flags are inconsistent.
        let mut payload = [0u8; 18];
        payload[0..2].copy_from_slice(&1i16.to_le_bytes()); // xWn
        payload[2..4].copy_from_slice(&2i16.to_le_bytes()); // yWn
        payload[4..6].copy_from_slice(&3u16.to_le_bytes()); // dxWn
        payload[6..8].copy_from_slice(&4u16.to_le_bytes()); // dyWn
        payload[8..10]
            .copy_from_slice(&(WINDOW1_GRBIT_ICONIC | WINDOW1_GRBIT_MAXIMIZED).to_le_bytes()); // grbit

        let stream = [
            record(RECORD_WINDOW1, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            globals.workbook_window,
            Some(WorkbookWindow {
                x: Some(1),
                y: Some(2),
                width: Some(3),
                height: Some(4),
                state: Some(WorkbookWindowState::Minimized)
            })
        );
        assert!(
            globals.warnings.is_empty(),
            "expected no warnings, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn window1_no_state_bits_maps_to_normal_state() {
        // Ensure we still record Normal state when geometry is present (so the record isn't treated
        // as empty metadata).
        let mut payload = [0u8; 18];
        payload[0..2].copy_from_slice(&5i16.to_le_bytes()); // xWn
        payload[2..4].copy_from_slice(&6i16.to_le_bytes()); // yWn
        payload[4..6].copy_from_slice(&7u16.to_le_bytes()); // dxWn
        payload[6..8].copy_from_slice(&8u16.to_le_bytes()); // dyWn
        payload[8..10].copy_from_slice(&0u16.to_le_bytes()); // grbit

        let stream = [
            record(RECORD_WINDOW1, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            globals.workbook_window,
            Some(WorkbookWindow {
                x: Some(5),
                y: Some(6),
                width: Some(7),
                height: Some(8),
                state: Some(WorkbookWindowState::Normal)
            })
        );
        assert!(
            globals.warnings.is_empty(),
            "expected no warnings, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn window1_tabs_bit_does_not_map_to_maximized_state() {
        // Real-world `.xls` files often set WINDOW1.grbit = 0x0038 (hscroll|vscroll|tabs).
        // Ensure the `tabs` UI bit (0x0020) is not interpreted as `fMaximized`.
        let mut payload = [0u8; 18];
        payload[0..2].copy_from_slice(&9i16.to_le_bytes()); // xWn
        payload[2..4].copy_from_slice(&10i16.to_le_bytes()); // yWn
        payload[4..6].copy_from_slice(&11u16.to_le_bytes()); // dxWn
        payload[6..8].copy_from_slice(&12u16.to_le_bytes()); // dyWn
        payload[8..10].copy_from_slice(&0x0020u16.to_le_bytes()); // grbit (tabs UI bit)

        let stream = [
            record(RECORD_WINDOW1, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            globals.workbook_window,
            Some(WorkbookWindow {
                x: Some(9),
                y: Some(10),
                width: Some(11),
                height: Some(12),
                state: Some(WorkbookWindowState::Normal),
            })
        );
    }

    #[test]
    fn workbook_protection_warns_on_truncated_protect_but_continues() {
        // PROTECT record with a 1-byte payload (too short for u16).
        let stream = [
            record(RECORD_PROTECT, &[1]),
            record(RECORD_WINDOWPROTECT, &1u16.to_le_bytes()),
            record(RECORD_PASSWORD, &0x83AFu16.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(globals.workbook_protection.lock_structure, false);
        assert_eq!(globals.workbook_protection.lock_windows, true);
        assert_eq!(globals.workbook_protection.password_hash, Some(0x83AF));
        assert!(
            globals
                .warnings
                .iter()
                .any(|w| w.contains("truncated PROTECT record")),
            "expected truncated-PROTECT warning, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn workbook_protection_warns_on_truncated_windowprotect_but_continues() {
        // WINDOWPROTECT record with a 1-byte payload (too short for u16).
        let stream = [
            record(RECORD_WINDOWPROTECT, &[1]),
            record(RECORD_PROTECT, &1u16.to_le_bytes()),
            record(RECORD_PASSWORD, &0x83AFu16.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(globals.workbook_protection.lock_structure, true);
        assert_eq!(globals.workbook_protection.lock_windows, false);
        assert_eq!(globals.workbook_protection.password_hash, Some(0x83AF));
        assert!(
            globals
                .warnings
                .iter()
                .any(|w| w.contains("truncated WINDOWPROTECT record")),
            "expected truncated-WINDOWPROTECT warning, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn workbook_protection_warns_on_truncated_password_but_continues() {
        // PASSWORD record with a 1-byte payload (too short for u16).
        let stream = [
            record(RECORD_PASSWORD, &[0xAF]),
            record(RECORD_PROTECT, &1u16.to_le_bytes()),
            record(RECORD_WINDOWPROTECT, &1u16.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(globals.workbook_protection.lock_structure, true);
        assert_eq!(globals.workbook_protection.lock_windows, true);
        assert_eq!(globals.workbook_protection.password_hash, None);
        assert!(
            globals
                .warnings
                .iter()
                .any(|w| w.contains("truncated PASSWORD record")),
            "expected truncated-PASSWORD warning, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn parses_globals_date_system_formats_and_xfs_biff8() {
        // 1904 record payload: f1904 = 1.
        let r_1904 = record(RECORD_1904, &[1, 0]);

        // FORMAT record: id=164, code="0.00" as XLUnicodeString (compressed).
        let mut fmt_payload = Vec::new();
        fmt_payload.extend_from_slice(&164u16.to_le_bytes());
        fmt_payload.extend_from_slice(&4u16.to_le_bytes()); // cch
        fmt_payload.push(0); // flags (compressed)
        fmt_payload.extend_from_slice(b"0.00");
        let r_fmt = record(RECORD_FORMAT_BIFF8, &fmt_payload);

        // XF record referencing numFmtId=164, cell xf (fStyle=0).
        let mut xf_payload = vec![0u8; 20];
        xf_payload[2..4].copy_from_slice(&164u16.to_le_bytes());
        xf_payload[4..6].copy_from_slice(&0u16.to_le_bytes());
        let r_xf = record(RECORD_XF, &xf_payload);

        let r_eof = record(records::RECORD_EOF, &[]);

        let mut stream = Vec::new();
        stream.extend_from_slice(&r_1904);
        stream.extend_from_slice(&r_fmt);
        stream.extend_from_slice(&r_xf);
        stream.extend_from_slice(&r_eof);

        let codepage = parse_biff_codepage(&stream);
        let globals =
            parse_biff_workbook_globals(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(globals.date_system, DateSystem::Excel1904);
        assert_eq!(
            globals.resolve_number_format_code(0).as_deref(),
            Some("0.00")
        );
    }

    #[test]
    fn resolves_builtins_and_placeholders() {
        let r_1900 = record(RECORD_1904, &[0, 0]);

        // Two XF records: one built-in (14), one unknown (60), and one General (0).
        let mut xf14 = vec![0u8; 20];
        xf14[2..4].copy_from_slice(&14u16.to_le_bytes());
        let mut xf60 = vec![0u8; 20];
        xf60[2..4].copy_from_slice(&60u16.to_le_bytes());
        let mut xf0 = vec![0u8; 20];
        xf0[2..4].copy_from_slice(&0u16.to_le_bytes());

        let stream = [
            r_1900,
            record(RECORD_XF, &xf14),
            record(RECORD_XF, &xf60),
            record(RECORD_XF, &xf0),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let codepage = parse_biff_codepage(&stream);
        let globals =
            parse_biff_workbook_globals(&stream, BiffVersion::Biff8, codepage).expect("parse");
        assert_eq!(
            globals.resolve_number_format_code(0).as_deref(),
            Some("m/d/yyyy")
        );
        assert_eq!(
            globals.resolve_number_format_code(1).as_deref(),
            Some("__builtin_numFmtId:60")
        );
        assert_eq!(globals.resolve_number_format_code(2), None);
        assert_eq!(globals.resolve_number_format_code(99), None);
    }

    #[test]
    fn parses_biff5_format_strings_and_strips_nuls() {
        // FORMAT2 record: id=200, "0\\0.00" (embedded NUL) as short ANSI string.
        let mut fmt_payload = Vec::new();
        fmt_payload.extend_from_slice(&200u16.to_le_bytes());
        fmt_payload.push(5); // cch (including NUL)
        fmt_payload.extend_from_slice(b"0\0.00");
        let r_fmt = record(RECORD_FORMAT2_BIFF5, &fmt_payload);

        let mut xf_payload = vec![0u8; 16];
        xf_payload[2..4].copy_from_slice(&200u16.to_le_bytes());

        let stream = [
            r_fmt,
            record(RECORD_XF, &xf_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let codepage = parse_biff_codepage(&stream);
        let globals =
            parse_biff_workbook_globals(&stream, BiffVersion::Biff5, codepage).expect("parse");
        assert_eq!(
            globals.resolve_number_format_code(0).as_deref(),
            Some("0.00")
        );
    }

    #[test]
    fn parses_externsheet_across_continue_records_in_workbook_globals() {
        // EXTERNSHEET payload: cXTI=2 with two internal sheet entries.
        let mut payload = Vec::new();
        payload.extend_from_slice(&2u16.to_le_bytes()); // cXTI
        // Entry 0: supbook=0, itabFirst=1, itabLast=1
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        // Entry 1: supbook=0, itabFirst=2, itabLast=3
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&2u16.to_le_bytes());
        payload.extend_from_slice(&3u16.to_le_bytes());

        // Split so a u16 spans the EXTERNSHEET/CONTINUE boundary.
        let split = 2 + 6 + 1; // cXTI + first entry + 1 byte of second entry's iSupBook
        let first = &payload[..split];
        let second = &payload[split..];

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, first),
            record(records::RECORD_CONTINUE, second),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            globals.extern_sheets,
            vec![
                externsheet::ExternSheetEntry {
                    supbook: 0,
                    itab_first: 1,
                    itab_last: 1,
                },
                externsheet::ExternSheetEntry {
                    supbook: 0,
                    itab_first: 2,
                    itab_last: 3,
                },
            ]
        );
        assert!(
            globals.warnings.is_empty(),
            "expected no warnings, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn workbook_globals_warns_on_externsheet_external_supbook() {
        // EXTERNSHEET payload: one entry with iSupBook!=0 (external workbook/add-in).
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u16.to_le_bytes()); // cXTI
        payload.extend_from_slice(&2u16.to_le_bytes()); // iSupBook != 0 => external
        payload.extend_from_slice(&0u16.to_le_bytes()); // itabFirst
        payload.extend_from_slice(&0u16.to_le_bytes()); // itabLast

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(globals.extern_sheets.len(), 1);
        assert_eq!(globals.extern_sheets[0].supbook, 2);
        assert!(
            globals
                .warnings
                .iter()
                .any(|w| w.contains("external SupBook references")),
            "expected external-supbook warning, got {:?}",
            globals.warnings
        );
    }

    #[test]
    fn palette_count_is_capped_by_available_bytes() {
        // Declares 1000 colors but only provides a single `LongRGB` entry (4 bytes).
        let mut data = Vec::new();
        data.extend_from_slice(&1000u16.to_le_bytes());
        data.extend_from_slice(&[0x10, 0x20, 0x30, 0x00]);

        let palette = parse_biff_palette_record(&data).expect("palette");
        assert_eq!(palette.len(), 1);
        assert_eq!(palette[0], 0xFF102030);
    }
}
