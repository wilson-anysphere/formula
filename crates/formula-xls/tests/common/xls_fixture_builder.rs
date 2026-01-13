#![allow(dead_code)]

use std::io::{Cursor, Write};

use formula_model::{
    indexed_color_argb, EXCEL_MAX_COLS, XLNM_FILTER_DATABASE, XLNM_PRINT_AREA, XLNM_PRINT_TITLES,
};

// This fixture builder writes just enough BIFF8 to exercise the importer. Keep record ids and
// commonly-used BIFF constants named so the intent stays readable.
const RECORD_BOF: u16 = 0x0809;
const RECORD_EOF: u16 = 0x000A;
const RECORD_CODEPAGE: u16 = 0x0042;
const RECORD_DATEMODE: u16 = 0x0022;
const RECORD_PROTECT: u16 = 0x0012;
const RECORD_PASSWORD: u16 = 0x0013;
const RECORD_WINDOWPROTECT: u16 = 0x0019;
const RECORD_WINDOW1: u16 = 0x003D;
const RECORD_FILEPASS: u16 = 0x002F;
const RECORD_FONT: u16 = 0x0031;
const RECORD_CALCCOUNT: u16 = 0x000C;
const RECORD_CALCMODE: u16 = 0x000D;
const RECORD_PRECISION: u16 = 0x000E;
const RECORD_DELTA: u16 = 0x0010;
const RECORD_ITERATION: u16 = 0x0011;
const RECORD_PALETTE: u16 = 0x0092;
const RECORD_FORMAT: u16 = 0x041E;
const RECORD_CONTINUE: u16 = 0x003C;
const RECORD_NAME: u16 = 0x0018;
const RECORD_NOTE: u16 = 0x001C;
const RECORD_OBJ: u16 = 0x005D;
const RECORD_TXO: u16 = 0x01B6;
const RECORD_XF: u16 = 0x00E0;
const RECORD_BOUNDSHEET: u16 = 0x0085;
const RECORD_SUPBOOK: u16 = 0x01AE;
const RECORD_EXTERNNAME: u16 = 0x0023;
const RECORD_EXTERNSHEET: u16 = 0x0017;
const RECORD_SAVERECALC: u16 = 0x005F;
const RECORD_SHEETEXT: u16 = 0x0862;
const RECORD_FEATHEADR: u16 = 0x0867;
const RECORD_FEAT: u16 = 0x0868;
const RECORD_FEATHEADR11: u16 = 0x0870;
const RECORD_FEAT11: u16 = 0x0871;
const RECORD_WINDOW2: u16 = 0x023E;
const RECORD_SCL: u16 = 0x00A0;
const RECORD_SETUP: u16 = 0x00A1;
const RECORD_HPAGEBREAKS: u16 = 0x001B;
const RECORD_VPAGEBREAKS: u16 = 0x001A;
const RECORD_LEFTMARGIN: u16 = 0x0026;
const RECORD_RIGHTMARGIN: u16 = 0x0027;
const RECORD_TOPMARGIN: u16 = 0x0028;
const RECORD_BOTTOMMARGIN: u16 = 0x0029;
const RECORD_PANE: u16 = 0x0041;
const RECORD_SELECTION: u16 = 0x001D;
const RECORD_DIMENSIONS: u16 = 0x0200;
const RECORD_MERGEDCELLS: u16 = 0x00E5;
const RECORD_BLANK: u16 = 0x0201;
const RECORD_NUMBER: u16 = 0x0203;
const RECORD_FORMULA: u16 = 0x0006;
/// SHRFMLA [MS-XLS 2.4.277] stores a shared formula (rgce) for a range.
const RECORD_SHRFMLA: u16 = 0x04BC;
const RECORD_HLINK: u16 = 0x01B8;
const RECORD_AUTOFILTERINFO: u16 = 0x009D;
const RECORD_SORT: u16 = 0x0090;
const RECORD_FILTERMODE: u16 = 0x009B;
// Excel 2007+ may store newer filter semantics in BIFF8 via future records.
const RECORD_AUTOFILTER12: u16 = 0x087E;
const RECORD_WSBOOL: u16 = 0x0081;
const RECORD_HORIZONTALPAGEBREAKS: u16 = 0x001B;
const RECORD_VERTICALPAGEBREAKS: u16 = 0x001A;
const RECORD_ROW: u16 = 0x0208;
const RECORD_COLINFO: u16 = 0x007D;
const RECORD_OBJPROTECT: u16 = 0x0063;
const RECORD_SCENPROTECT: u16 = 0x00DD;

const ROW_OPTION_HIDDEN: u16 = 0x0020;
const ROW_OPTION_COLLAPSED: u16 = 0x1000;
const COLINFO_OPTION_HIDDEN: u16 = 0x0001;
const COLINFO_OPTION_COLLAPSED: u16 = 0x1000;

// WSBOOL options.
const WSBOOL_OPTION_FIT_TO_PAGE: u16 = 0x0100;

const BOF_VERSION_BIFF8: u16 = 0x0600;
const BOF_DT_WORKBOOK_GLOBALS: u16 = 0x0005;
const BOF_DT_WORKSHEET: u16 = 0x0010;

const XF_FLAG_LOCKED: u16 = 0x0001;
const XF_FLAG_STYLE: u16 = 0x0004;

const COLOR_AUTOMATIC: u16 = 0x7FFF;

/// Build a minimal BIFF8 `.xls` fixture containing a single sheet named `Formats`.
///
/// The goal is not to be a complete `.xls` writer; it's just enough BIFF8 + CFB
/// to exercise our importer with targeted style payloads (FORMAT/XF + BLANK).
pub fn build_number_format_fixture_xls(date_1904: bool) -> Vec<u8> {
    let workbook_stream = build_workbook_stream(date_1904, false);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build the BIFF workbook stream bytes for [`build_number_format_fixture_xls`] while injecting a
/// `FILEPASS` record into the workbook globals substream.
///
/// This is useful for testing post-decryption parsing: encrypted workbooks retain the `FILEPASS`
/// record header even after decrypting the remaining stream bytes.
pub fn build_number_format_workbook_stream_with_filepass(date_1904: bool) -> Vec<u8> {
    build_workbook_stream(date_1904, true)
}

/// Build a minimal BIFF8 `.xls` fixture whose workbook globals include a `FILEPASS` record.
///
/// The presence of `FILEPASS` indicates the workbook stream is encrypted/password-protected.
pub fn build_encrypted_filepass_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_encrypted_filepass_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing visible, hidden, and very hidden sheets.
///
/// This is used to validate that sheet visibility can be recovered from BIFF `BoundSheet8`
/// (`hsState`) metadata.
pub fn build_sheet_visibility_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_sheet_visibility_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that forces sheet-name sanitization via truncation and includes a
/// cross-sheet formula referencing the original (over-long) name.
///
/// This exercises the importer’s ability to rewrite formula sheet references after
/// sanitizing invalid sheet names.
pub fn build_formula_sheet_name_truncation_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_formula_sheet_name_truncation_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a malformed shared-formula pattern:
/// - Base cell contains a full `FORMULA.rgce` token stream (no `PtgExp`).
/// - Follower cell contains only `PtgExp` pointing at the base cell.
/// - The expected `SHRFMLA`/`ARRAY` definition record is intentionally missing.
///
/// The `.xls` importer should still recover the follower formula by materializing from the base
/// cell's `FORMULA.rgce`.
pub fn build_shared_formula_ptgexp_missing_shrfmla_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgexp_missing_shrfmla_workbook_stream();
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a worksheet formula that calamine drops from
/// `worksheet_formula()` output (due to unresolved shared-formula state), but can be decoded by our
/// BIFF `rgce` parser.
///
/// Our importer should fall back to decoding BIFF8 `FORMULA` records directly so formulas are still
/// imported.
pub fn build_calamine_formula_error_biff_fallback_fixture_xls() -> Vec<u8> {
    let xf_general = 16u16;
    let sheet_stream = build_spill_operator_formula_sheet_stream(xf_general);
    let workbook_stream = build_single_sheet_workbook_stream("Sheet1", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a shared formula over `B1:B2`:
/// - `B1`: `SUM(A1,1)`
/// - `B2`: `SUM(A2,1)` via `PtgExp` referencing `B1`
///
/// The shared formula body is stored in a `SHRFMLA` record and contains a `PtgFuncVar` token
/// (variable-arity function) to exercise decoding of its payload (argc + function id).
pub fn build_shared_formula_ptgfuncvar_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgfuncvar_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with a merged region (`A1:B1`) where only the
/// non-anchor cell (`B1`) has a formatted `BLANK` record.
///
/// This exercises the importer’s “apply styles to merged-cell anchors” behaviour:
/// formatting attached to any cell inside the merged region must be applied to
/// the anchor cell so it round-trips in our model.
pub fn build_merged_formatted_blank_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_merged_formatted_blank_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with a merged region (`A1:B1`) where both cells carry
/// formatted `BLANK` records, but the formats conflict.
///
/// The importer should prefer the anchor cell’s XF (`A1`) over non-anchor XF indices (`B1`),
/// matching the model’s “anchor cell owns merged region formatting” semantics.
pub fn build_merged_conflicting_blank_formats_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_merged_conflicting_blank_formats_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that references an out-of-range XF index in a cell record.
///
/// BIFF files in the wild can contain corrupt style indices. The importer should skip those
/// assignments but surface a single aggregated warning per sheet rather than failing or
/// spamming logs.
pub fn build_out_of_range_xf_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_out_of_range_xf_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with a merged region (`A1:C1`) where the anchor cell has no
/// record, but two non-anchor cells (`B1` and `C1`) contain conflicting formatted `BLANK` records.
///
/// The importer should pick a deterministic format to apply to the merged-region anchor.
pub fn build_merged_non_anchor_conflicting_blank_formats_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_merged_non_anchor_conflicting_blank_formats_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with a merged region (`A1:B1`) where the only `FORMULA` record is
/// stored on the non-anchor cell (`B1`).
///
/// The importer should normalize formulas inside merged regions to the top-left anchor (`A1`),
/// matching the model semantics for values/styles/comments/hyperlinks.
pub fn build_merged_non_anchor_formula_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_merged_non_anchor_formula_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that references an out-of-range XF index in a cell record, where
/// the workbook contains **no** non-General number formats.
///
/// This ensures we still detect and warn on corrupt XF indices even when the workbook's
/// `XF/FORMAT` table doesn't contain any "interesting" number formats.
pub fn build_out_of_range_xf_no_formats_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_out_of_range_xf_no_formats_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that applies default formatting at the row and column level
/// (ROW/COLINFO `ixfe`) without any cell records referencing those XF indices.
pub fn build_row_col_style_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_row_col_style_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that stores a long custom number format split across a `CONTINUE`
/// record.
///
/// This exercises the importer’s handling of continued BIFF8 `FORMAT` records.
pub fn build_continued_format_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_continued_format_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with grouped rows/columns and collapsed outline groups.
///
/// This is used to validate that the importer preserves Excel outline metadata (levels,
/// collapsed summary rows/cols, and derived outline-hidden state).
pub fn build_outline_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_outline_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that uses an unknown/reserved built-in `numFmtId` value without
/// providing an explicit `FORMAT` record.
///
/// The importer should preserve the `numFmtId` via a placeholder number format string like
/// `__builtin_numFmtId:60`.
pub fn build_unknown_builtin_numfmtid_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_unknown_builtin_numfmtid_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing workbook- and sheet-scoped defined names.
///
/// The fixture includes:
/// - A workbook-scoped name referencing `Sheet1!$A$1` via `PtgRef3d`
/// - A sheet-scoped name (local scope / `itab`)
/// - A built-in print area name (`_xlnm.Print_Area`) using a union of two areas
///
/// It also includes additional NAME records to exercise `rgce` decoding paths (union operators,
/// function calls, missing args, hidden flags), and a minimal `EXTERNSHEET` table so 3D references
/// can be rendered.
pub fn build_defined_names_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_defined_names_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a defined name whose `NAME` record is split across
/// `CONTINUE` records.
///
/// The payload is intentionally split twice:
/// - within the `rgce` token stream
/// - within the description string
///
/// This exercises the importer’s handling of continued BIFF8 `NAME` records.
pub fn build_continued_name_record_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_continued_name_record_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing workbook-scoped defined names that reference sheets
/// requiring quoting (spaces, embedded quotes, reserved words), plus a 3D sheet span.
///
/// This is used to validate that our BIFF8 `rgce` decoder produces sheet-name prefixes that are:
/// - accepted by Excel conventions (proper quoting/escaping), and
/// - parseable by `formula-engine`.
pub fn build_defined_names_quoting_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_defined_names_quoting_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing workbook-scoped defined names that reference an
/// *external* workbook via `SUPBOOK`/`EXTERNSHEET`.
///
/// This validates our best-effort rendering of external 3D references into Excel-style text like:
/// - `'[Book1.xlsx]SheetA'!$A$1`
/// - `'[Book1.xlsx]SheetA:SheetC'!$A$1`
pub fn build_defined_names_external_workbook_refs_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_defined_names_external_workbook_refs_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a minimal BIFF8 `.xls` fixture containing a single workbook-scoped defined name.
///
/// This fixture is intended to validate the importer’s calamine `Reader::defined_names()` fallback
/// path (i.e. when BIFF workbook parsing is unavailable).
pub fn build_defined_name_calamine_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_defined_name_calamine_workbook_stream_with_sheet_name("Sheet1");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where the sheet name is invalid and will be sanitized by the
/// importer, but a calamine-surfaced defined name still references the original name.
///
/// This is used to verify that calamine fallback defined-name formulas are rewritten after
/// sheet-name sanitization, matching the cell formula rewrite behavior.
pub fn build_defined_name_sheet_name_sanitization_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_defined_name_calamine_workbook_stream_with_sheet_name("Bad:Name");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where a worksheet name is invalid and will be sanitized by the
/// importer, but a workbook-scoped defined name (surfaced via calamine) still references the
/// original name.
///
/// This is used to validate that the `.xls` importer rewrites calamine-defined-name formulas after
/// sheet name sanitization, similar to how it rewrites cell formulas and internal hyperlinks.
pub fn build_defined_name_sheet_name_sanitization_calamine_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_defined_name_sheet_name_sanitization_calamine_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing workbook-scoped defined names that mimic print
/// settings:
/// - `_xlnm.Print_Area` referencing Sheet1,
/// - `_xlnm.Print_Titles` referencing Sheet2.
///
/// This is intended to validate the `.xls` importer’s ability to infer sheet-scoped print
/// settings from calamine-defined names when BIFF workbook parsing is unavailable.
pub fn build_print_settings_calamine_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_print_settings_calamine_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a minimal BIFF8 `.xls` fixture containing a single sheet named `Sheet1` where
/// `WSBOOL.fFitToPage=1` (fit-to-page enabled) but the worksheet substream omits the `SETUP` record.
///
/// Some `.xls` writers omit `SETUP` even when fit-to-page is enabled; the importer should preserve
/// the scaling intent as `Scaling::FitTo { width: 0, height: 0 }`.
pub fn build_fit_to_page_without_setup_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_fit_to_page_without_setup_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where the worksheet name is invalid and will be sanitized by the
/// importer, but at least one defined name still refers to the original (invalid) sheet name.
pub fn build_sanitized_sheet_name_defined_name_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_sanitized_sheet_name_defined_name_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where sanitizing one invalid worksheet name causes a collision with
/// another sheet's *original* name.
///
/// This is used to ensure defined-name sheet-reference rewriting does not mistakenly rewrite
/// already-sanitized BIFF-defined-name formulas (which are resolved by sheet index) into the
/// colliding sheet.
pub fn build_sanitized_sheet_name_defined_name_collision_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_sanitized_sheet_name_defined_name_collision_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a sheet-scoped built-in `_xlnm.Print_Area` defined name
/// whose `refers_to` string uses a **quoted** sheet name containing non-ASCII characters.
///
/// This validates that print settings parsing handles quoted UTF-8 sheet names (e.g.
/// `'Ünicode Name'!$A$1:$A$2`) correctly.
pub fn build_print_settings_unicode_sheet_name_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_print_settings_unicode_sheet_name_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing explicit manual page breaks in the worksheet substream
/// (`HORIZONTALPAGEBREAKS` / `VERTICALPAGEBREAKS`).
pub fn build_manual_page_breaks_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_manual_page_breaks_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a malformed `HORIZONTALPAGEBREAKS` record with an
/// unreasonable break count (`cbrk=0xFFFF`) but only a single entry worth of payload bytes.
///
/// This is used to regression-test that the importer caps loop counts based on record length and
/// spec limits, preventing pathological CPU loops on corrupt files.
pub fn build_page_break_cbrk_cap_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_page_break_cbrk_cap_sheet_stream(16);
    let workbook_stream = build_single_sheet_workbook_stream("PageBreaks", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

fn build_page_break_cbrk_cap_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1 (A)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Malformed `HORIZONTALPAGEBREAKS` payload:
    // [cbrk:u16][(rw:u16, colStart:u16, colEnd:u16) * cbrk]
    // Here, `cbrk` claims 65535 entries, but we only provide one.
    let mut breaks = Vec::<u8>::new();
    breaks.extend_from_slice(&0xFFFFu16.to_le_bytes());
    breaks.extend_from_slice(&1u16.to_le_bytes()); // rw=1 => break after row 0
    breaks.extend_from_slice(&0u16.to_le_bytes()); // colStart
    breaks.extend_from_slice(&255u16.to_le_bytes()); // colEnd
    push_record(&mut sheet, RECORD_HORIZONTALPAGEBREAKS, &breaks);

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

/// Build a BIFF8 `.xls` fixture that stores worksheet print settings (page setup, margins and
/// manual page breaks) in the worksheet substream.
pub fn build_sheet_print_settings_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_sheet_print_settings_sheet_stream(16);
    let workbook_stream = build_single_sheet_workbook_stream("Print", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

#[derive(Debug, Clone, Copy)]
enum PageSetupScalingMode {
    Percent,
    FitTo,
}

/// Build a BIFF8 `.xls` fixture containing worksheet page setup + margins + manual page breaks,
/// using percent scaling (`WSBOOL.fFitToPage=0`, `SETUP.iScale=85`).
pub fn build_page_setup_percent_scaling_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_page_setup_fixture_sheet_stream(16, PageSetupScalingMode::Percent);
    let workbook_stream = build_single_sheet_workbook_stream("Sheet1", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing worksheet page setup + margins + manual page breaks,
/// using fit-to scaling (`WSBOOL.fFitToPage=1`, `SETUP.iFitWidth=2`, `SETUP.iFitHeight=3`).
pub fn build_page_setup_fit_to_scaling_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_page_setup_fixture_sheet_stream(16, PageSetupScalingMode::FitTo);
    let workbook_stream = build_single_sheet_workbook_stream("Sheet1", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a worksheet `SETUP` record with a custom paper size
/// (`iPaperSize=0`, `fNoPls=0`).
pub fn build_custom_paper_size_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_custom_paper_size_sheet_stream(16);
    let workbook_stream = build_single_sheet_workbook_stream("Sheet1", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

fn build_page_setup_fixture_sheet_stream(xf_cell: u16, mode: PageSetupScalingMode) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 10) cols [0, 5) (large enough to cover our page breaks).
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&10u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&5u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    // WINDOW2 is required by some consumers; keep defaults.
    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // WSBOOL controls scaling mode (`fFitToPage`).
    let mut wsbool: u16 = 0x0C01;
    match mode {
        PageSetupScalingMode::Percent => wsbool &= !WSBOOL_OPTION_FIT_TO_PAGE,
        PageSetupScalingMode::FitTo => wsbool |= WSBOOL_OPTION_FIT_TO_PAGE,
    }
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool.to_le_bytes());

    // Page setup: Landscape + A4 + scaling + non-default header/footer margins.
    match mode {
        PageSetupScalingMode::Percent => push_record(
            &mut sheet,
            RECORD_SETUP,
            &setup_record(9, 85, 0, 0, true, 0.9, 1.0),
        ),
        PageSetupScalingMode::FitTo => push_record(
            &mut sheet,
            RECORD_SETUP,
            &setup_record(9, 100, 2, 3, true, 0.9, 1.0),
        ),
    }

    // Margins.
    push_record(&mut sheet, RECORD_LEFTMARGIN, &0.5f64.to_le_bytes());
    push_record(&mut sheet, RECORD_RIGHTMARGIN, &0.6f64.to_le_bytes());
    push_record(&mut sheet, RECORD_TOPMARGIN, &0.7f64.to_le_bytes());
    push_record(&mut sheet, RECORD_BOTTOMMARGIN, &0.8f64.to_le_bytes());

    // Manual page breaks.
    // Note: BIFF8 page breaks store the 0-based index of the first row/col *after* the break.
    // The importer converts these to the model’s “after which break occurs” form by subtracting 1.
    push_record(&mut sheet, RECORD_HPAGEBREAKS, &hpagebreaks_record(&[5]));
    push_record(&mut sheet, RECORD_VPAGEBREAKS, &vpagebreaks_record(&[3]));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 0.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

/// Build a BIFF8 `.xls` fixture with a single sheet whose BoundSheet name is invalid per Excel
/// rules (`Bad:Name`) and must be sanitized by the importer.
///
/// The worksheet substream includes:
/// - SETUP (paper size, orientation, scaling, header/footer margins)
/// - LEFT/RIGHT/TOP/BOTTOMMARGIN
/// - HORIZONTALPAGEBREAKS / VERTICALPAGEBREAKS
///
/// This fixture is used to ensure BIFF-derived print settings are applied to the sanitized sheet
/// name stored in the output workbook (`Bad_Name`), not the original BIFF BoundSheet name.
pub fn build_page_setup_sanitized_sheet_name_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_page_setup_sanitized_sheet_name_sheet_stream(16);
    let workbook_stream = build_single_sheet_workbook_stream("Bad:Name", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

fn build_page_setup_sanitized_sheet_name_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: keep large enough to cover our page break positions.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&6u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&4u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    // WINDOW2 is required by some consumers; keep defaults.
    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Margins (distinct values, in inches).
    push_record(&mut sheet, RECORD_LEFTMARGIN, &1.11f64.to_le_bytes());
    push_record(&mut sheet, RECORD_RIGHTMARGIN, &2.22f64.to_le_bytes());
    push_record(&mut sheet, RECORD_TOPMARGIN, &3.33f64.to_le_bytes());
    push_record(&mut sheet, RECORD_BOTTOMMARGIN, &4.44f64.to_le_bytes());

    // Page setup:
    // - A4 paper size (9)
    // - 123% scaling
    // - Landscape orientation (fPortrait=0)
    // - Non-default header/footer margins
    let mut setup = Vec::<u8>::new();
    setup.extend_from_slice(&9u16.to_le_bytes()); // iPaperSize: 9 = A4
    setup.extend_from_slice(&123u16.to_le_bytes()); // iScale: 123%
    setup.extend_from_slice(&0u16.to_le_bytes()); // iPageStart
    setup.extend_from_slice(&0u16.to_le_bytes()); // iFitWidth
    setup.extend_from_slice(&0u16.to_le_bytes()); // iFitHeight
    setup.extend_from_slice(&0u16.to_le_bytes()); // grbit: fPortrait=0 => landscape
    setup.extend_from_slice(&600u16.to_le_bytes()); // iRes
    setup.extend_from_slice(&600u16.to_le_bytes()); // iVRes
    setup.extend_from_slice(&0.55f64.to_le_bytes()); // numHdr
    setup.extend_from_slice(&0.66f64.to_le_bytes()); // numFtr
    setup.extend_from_slice(&1u16.to_le_bytes()); // iCopies
    push_record(&mut sheet, RECORD_SETUP, &setup);

    // Manual page breaks.
    // Note: BIFF8 page breaks store the 0-based index of the first row/col *after* the break.
    // The importer converts these to the model’s “after which break occurs” form by subtracting 1.
    push_record(&mut sheet, RECORD_HPAGEBREAKS, &hpagebreaks_record(&[2, 5]));
    push_record(&mut sheet, RECORD_VPAGEBREAKS, &vpagebreaks_record(&[3]));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_sheet_print_settings_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 5) cols [0, 3) => A1:C5.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1 (A..C)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    // WINDOW2 is required by some consumers; keep defaults.
    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // WSBOOL controls scaling mode. Enable fit-to-page so SETUP.iFit* fields are respected.
    let wsbool: u16 = 0x0C01 | WSBOOL_OPTION_FIT_TO_PAGE;
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool.to_le_bytes());

    // Margins.
    push_record(&mut sheet, RECORD_LEFTMARGIN, &1.1f64.to_le_bytes());
    push_record(&mut sheet, RECORD_RIGHTMARGIN, &1.2f64.to_le_bytes());
    push_record(&mut sheet, RECORD_TOPMARGIN, &1.3f64.to_le_bytes());
    push_record(&mut sheet, RECORD_BOTTOMMARGIN, &1.4f64.to_le_bytes());

    // Enable fit-to-page scaling so `SETUP.iFitWidth/iFitHeight` apply.
    push_record(&mut sheet, RECORD_WSBOOL, &WSBOOL_OPTION_FIT_TO_PAGE.to_le_bytes());

    // Page setup: Landscape + A4 + Fit to 2 pages wide by 3 tall + non-default header/footer
    // margins.
    push_record(
        &mut sheet,
        RECORD_SETUP,
        &setup_record(
            9,    // A4
            100,  // scale (ignored when fit-to is used)
            2,    // fit width
            3,    // fit height
            true, // landscape
            0.5,  // header margin
            0.6,  // footer margin
        ),
    );
    // Enable fit-to-page scaling (WSBOOL.fFitToPage=1).
    push_record(&mut sheet, RECORD_WSBOOL, &0x0100u16.to_le_bytes());

    // Manual page breaks.
    // Note: BIFF8 page breaks store the 0-based index of the first row/col *after* the break.
    // The importer converts these to the model’s “after which break occurs” form by subtracting 1.
    //
    // This fixture encodes breaks-after rows {2,4} as rw={3,5} and break-after col {1} as col={2}.
    push_record(&mut sheet, RECORD_HPAGEBREAKS, &hpagebreaks_record(&[3, 5]));
    push_record(&mut sheet, RECORD_VPAGEBREAKS, &vpagebreaks_record(&[2]));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_custom_paper_size_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: A1:A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Custom paper size (`iPaperSize=0`) with `fNoPls=0`.
    //
    // Populate other fields with defaults so the resulting PageSetup stays as close to the model
    // default as possible (so tests can focus on paper size handling).
    push_record(
        &mut sheet,
        RECORD_SETUP,
        &setup_record(
            0,    // iPaperSize (custom)
            100,  // iScale
            0,    // iFitWidth
            0,    // iFitHeight
            true, // set fPortrait bit (see importer semantics)
            0.3,  // header margin (default)
            0.3,  // footer margin (default)
        ),
    );

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn setup_record(
    paper_size: u16,
    scale: u16,
    fit_width: u16,
    fit_height: u16,
    landscape: bool,
    header_margin: f64,
    footer_margin: f64,
) -> Vec<u8> {
    // BIFF8 SETUP record payload.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&paper_size.to_le_bytes()); // iPaperSize
    out.extend_from_slice(&scale.to_le_bytes()); // iScale
    out.extend_from_slice(&0u16.to_le_bytes()); // iPageStart
    out.extend_from_slice(&fit_width.to_le_bytes()); // iFitWidth
    out.extend_from_slice(&fit_height.to_le_bytes()); // iFitHeight
    let mut grbit = 0u16;
    // SETUP.grbit fPortrait (0x0002): when set, the sheet is in portrait mode; when clear, landscape.
    if !landscape {
        grbit |= 0x0002;
    }
    out.extend_from_slice(&grbit.to_le_bytes()); // grbit
    out.extend_from_slice(&600u16.to_le_bytes()); // iRes
    out.extend_from_slice(&600u16.to_le_bytes()); // iVRes
    out.extend_from_slice(&header_margin.to_le_bytes()); // numHdr
    out.extend_from_slice(&footer_margin.to_le_bytes()); // numFtr
    out.extend_from_slice(&1u16.to_le_bytes()); // iCopies
    out
}

fn hpagebreaks_record(breaks: &[u16]) -> Vec<u8> {
    // HORIZONTALPAGEBREAKS payload:
    // [cbrk:u16][(rw:u16, colStart:u16, colEnd:u16) * cbrk]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&(breaks.len() as u16).to_le_bytes());
    for &rw in breaks {
        out.extend_from_slice(&rw.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // colStart
        out.extend_from_slice(&255u16.to_le_bytes()); // colEnd (BIFF8 max col for Excel 97-2003)
    }
    out
}

fn vpagebreaks_record(breaks: &[u16]) -> Vec<u8> {
    // VERTICALPAGEBREAKS payload:
    // [cbrk:u16][(col:u16, rwStart:u16, rwEnd:u16) * cbrk]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&(breaks.len() as u16).to_le_bytes());
    for &col in breaks {
        out.extend_from_slice(&col.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // rwStart
        out.extend_from_slice(&65535u16.to_le_bytes()); // rwEnd (BIFF8 max row for Excel 97-2003)
    }
    out
}

/// Build a BIFF8 `.xls` fixture containing two worksheets (`First`, `Second`) that each store
/// distinct page setup (paper size, orientation, scaling), margins, and manual page breaks.
///
/// This is used to validate that BIFF sheet-index mapping / per-sheet application logic attaches
/// page setup metadata to the correct sheet in multi-sheet workbooks.
pub fn build_page_setup_multisheet_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_page_setup_multisheet_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a minimal BIFF8 `.xls` fixture containing a single sheet named `Notes`
/// with a NOTE/OBJ/TXO comment anchored to `A1`.
pub fn build_note_comment_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_workbook_stream(NoteCommentSheetKind::SingleCell);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a minimal BIFF5 `.xls` fixture containing a single sheet named `NotesBiff5`
/// with a NOTE/OBJ/TXO comment anchored to `A1`.
///
/// This exercises BIFF5 comment parsing paths:
/// - NOTE author stored as an ANSI short string (no BIFF8 flags byte)
/// - TXO text stored as raw bytes in CONTINUE records (no per-fragment flags byte)
/// - decoding via the workbook `CODEPAGE` record (here: 1251 / Windows-1251)
pub fn build_note_comment_biff5_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_biff5_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF5 `.xls` fixture containing a NOTE/OBJ/TXO comment whose TXO text is split across
/// multiple `CONTINUE` records (no per-fragment flags bytes).
pub fn build_note_comment_biff5_split_across_continues_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_biff5_split_across_continues_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF5 `.xls` fixture containing a NOTE/OBJ/TXO comment whose TXO text is split across
/// multiple `CONTINUE` records and each fragment begins with a BIFF8-style 0/1 "high-byte" flag
/// byte.
pub fn build_note_comment_biff5_split_across_continues_with_flags_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_note_comment_biff5_split_across_continues_with_flags_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF5 `.xls` fixture containing a NOTE/OBJ/TXO comment whose TXO text is split across
/// `CONTINUE` records in the middle of a multibyte codepage character (Shift-JIS / codepage 932).
pub fn build_note_comment_biff5_split_across_continues_codepage_932_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_note_comment_biff5_split_across_continues_codepage_932_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF5 `.xls` fixture where the NOTE author string is stored as a BIFF8-style
/// `ShortXLUnicodeString` (length + flags byte) even though the workbook is BIFF5.
pub fn build_note_comment_biff5_author_biff8_short_string_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_biff5_author_biff8_short_string_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF5 `.xls` fixture where the NOTE author string is stored as a BIFF8-style
/// `XLUnicodeString` (16-bit length + flags byte) even though the workbook is BIFF5.
pub fn build_note_comment_biff5_author_biff8_unicode_string_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_biff5_author_biff8_unicode_string_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF5 `.xls` fixture containing a single sheet with a merged region (`A1:B1`).
///
/// The NOTE record is targeted at the non-anchor cell (`B1`), but the importer should
/// anchor the resulting model comment to `A1` per Excel merged-cell semantics.
pub fn build_note_comment_biff5_in_merged_region_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_biff5_in_merged_region_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF5 `.xls` fixture containing a NOTE/OBJ pair but **no** TXO payload.
///
/// The importer should emit a warning and skip creating a model comment.
pub fn build_note_comment_biff5_missing_txo_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_biff5_missing_txo_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a merged region (`A1:B1`).
///
/// The NOTE record is targeted at the non-anchor cell (`B1`), but the importer should
/// anchor the resulting model comment to `A1` per Excel merged-cell semantics.
pub fn build_note_comment_in_merged_region_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_note_comment_workbook_stream(NoteCommentSheetKind::MergedRegionNonAnchorNote);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment, using
/// `CODEPAGE=1251` and a compressed 8-bit TXO text payload.
///
/// In Windows-1251, byte `0xC0` maps to Cyrillic "А" (U+0410). This fixture ensures we decode
/// comment text bytes using the workbook `CODEPAGE` record rather than assuming 1252.
pub fn build_note_comment_codepage_1251_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_codepage_1251_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment, using
/// `CODEPAGE=1251` and a compressed 8-bit NOTE author string containing non-ASCII bytes.
///
/// In Windows-1251, byte `0xC0` maps to Cyrillic "А" (U+0410). This fixture ensures we decode
/// NOTE author bytes using the workbook `CODEPAGE` record (not Windows-1252).
pub fn build_note_comment_author_codepage_1251_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_author_codepage_1251_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// NOTE author string is stored as an `XLUnicodeString` (16-bit length) instead of the usual
/// `ShortXLUnicodeString` (8-bit length).
///
/// Some `.xls` producers store NOTE authors in this non-standard form; the importer should still
/// decode the author correctly.
pub fn build_note_comment_author_xl_unicode_string_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_author_xl_unicode_string_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// NOTE author string is stored as a BIFF5-style short ANSI string (length + bytes), without the
/// BIFF8 `ShortXLUnicodeString` option flags byte.
///
/// Some `.xls` producers appear to omit the BIFF8 flags byte; the importer should still recover
/// the author via best-effort decoding.
pub fn build_note_comment_author_missing_flags_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_author_missing_flags_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// TXO text continuation payload omits the BIFF8 flags byte and instead stores raw ANSI bytes
/// (BIFF5-style).
///
/// Some `.xls` producers appear to omit the flags byte; the importer should still recover the
/// comment text via best-effort decoding.
pub fn build_note_comment_txo_text_missing_flags_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_txo_text_missing_flags_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// TXO text is split across multiple `CONTINUE` records and the *second* text fragment omits the
/// BIFF8 flags byte.
///
/// Some `.xls` producers appear to include the flags byte only for the first fragment; the
/// importer should still recover the full comment text.
pub fn build_note_comment_txo_text_missing_flags_in_second_fragment_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_note_comment_txo_text_missing_flags_in_second_fragment_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// TXO header reports `cchText=0` but still includes a `cbRuns` value and a continued text payload.
///
/// Some `.xls` producers appear to zero out the `cchText` field while still writing the text into
/// the continuation area; the importer should still recover the full comment text.
pub fn build_note_comment_txo_cch_text_zero_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_txo_cch_text_zero_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// TXO header reports `cchText=0` and the TXO text is split across multiple `CONTINUE` records,
/// with the *second* fragment omitting the BIFF8 flags byte.
///
/// This exercises a combination of real-world corruption patterns:
/// - missing/zero `cchText` requiring us to infer the length from the continuation area
/// - missing BIFF8 flags byte on a subsequent fragment
pub fn build_note_comment_txo_cch_text_zero_missing_flags_in_second_fragment_fixture_xls() -> Vec<u8>
{
    let workbook_stream =
        build_note_comment_txo_cch_text_zero_missing_flags_in_second_fragment_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment, where the
/// TXO text is split across multiple `CONTINUE` records.
///
/// Each continued segment starts with the BIFF8 string option flags byte (0 for compressed 8-bit).
pub fn build_note_comment_split_across_continues_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_split_across_continues_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment, where the
/// TXO text is split across multiple `CONTINUE` records and switches encoding between fragments.
///
/// The first fragment is stored as compressed 8-bit, while the second fragment is stored as UTF-16LE
/// (`fHighByte=1`). This ensures our TXO parser respects the per-fragment option flags byte at the
/// start of each continued-string fragment.
pub fn build_note_comment_split_across_continues_mixed_encoding_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_note_comment_split_across_continues_mixed_encoding_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment, where the
/// TXO text is split across multiple `CONTINUE` records and uses a multibyte codepage
/// (`CODEPAGE=932` / Shift-JIS).
///
/// This fixture intentionally splits a single multibyte character (`"あ"` = `0x82 0xA0` in Shift-JIS)
/// across two `CONTINUE` fragments to ensure we buffer/decode compressed bytes across record
/// boundaries (rather than decoding each fragment independently).
pub fn build_note_comment_split_across_continues_codepage_932_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_split_across_continues_codepage_932_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment, where the
/// TXO text is split across multiple `CONTINUE` records and uses `CODEPAGE=65001` (UTF-8).
///
/// The fixture splits a 3-byte UTF-8 character (`"€"` = `0xE2 0x82 0xAC`) across separate CONTINUE
/// fragments to ensure we buffer/decode compressed bytes across record boundaries.
pub fn build_note_comment_split_across_continues_codepage_65001_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_note_comment_split_across_continues_codepage_65001_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// UTF-16LE TXO text payload is split mid-code-unit across `CONTINUE` records.
///
/// This is technically malformed BIFF (Unicode fragments should contain an even number of bytes),
/// but some files in the wild appear to split UTF-16LE code units across record boundaries. The
/// importer should still recover the intended character.
pub fn build_note_comment_split_utf16_code_unit_across_continues_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_note_comment_split_utf16_code_unit_across_continues_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a NOTE/OBJ pair but **missing** the associated TXO text
/// payload.
///
/// This is used to validate best-effort import behavior: the importer should emit a warning and
/// skip creating a comment when the NOTE record cannot be joined to a TXO text payload.
pub fn build_note_comment_missing_txo_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_missing_txo_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// TXO record header is empty/truncated.
///
/// The importer should still recover the text by falling back to decoding the `CONTINUE`
/// fragments and surface a warning.
pub fn build_note_comment_missing_txo_header_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_missing_txo_header_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// TXO record header is empty/truncated and the TXO text is split across multiple `CONTINUE`
/// records, with the *second* fragment omitting the BIFF8 flags byte.
pub fn build_note_comment_missing_txo_header_missing_flags_in_second_fragment_fixture_xls(
) -> Vec<u8> {
    let workbook_stream =
        build_note_comment_missing_txo_header_missing_flags_in_second_fragment_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// TXO header is truncated such that the `cbRuns` field is missing.
///
/// The `cchText` field is intentionally larger than the available text bytes, and a formatting-run
/// `CONTINUE` payload follows the text. The importer should **not** decode formatting run bytes as
/// text.
pub fn build_note_comment_truncated_txo_header_missing_cb_runs_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_truncated_txo_header_missing_cb_runs_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet with a NOTE/OBJ/TXO comment where the
/// TXO `cchText` field is stored at an alternate offset (4 instead of the spec-defined offset 6).
pub fn build_note_comment_txo_cch_text_offset_4_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_txo_cch_text_offset_4_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a NOTE/OBJ/TXO comment where the NOTE record's
/// `grbit`/`idObj` fields are effectively swapped.
///
/// Some `.xls` producers appear to place the drawing object id in the NOTE `grbit` field (offset 4)
/// rather than the spec-defined `idObj` field (offset 6). Our importer should still join the NOTE
/// to its TXO payload by trying both fields.
pub fn build_note_comment_note_obj_id_swapped_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_note_comment_note_obj_id_swapped_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single external hyperlink on `A1`.
///
/// This is used to ensure we preserve BIFF `HLINK` records when importing `.xls` workbooks.
pub fn build_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "Links",
        hlink_external_url(
            0,
            0,
            0,
            0,
            "https://example.com",
            "Example",
            "Example tooltip",
        ),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a URL hyperlink whose URL moniker contains an embedded
/// NUL character followed by trailing garbage.
///
/// The importer should truncate the URL at the first NUL for best-effort compatibility.
pub fn build_url_hyperlink_embedded_nul_fixture_xls() -> Vec<u8> {
    let url = "https://example.com\0ignored";
    let workbook_stream =
        build_hyperlink_workbook_stream("UrlNul", hlink_external_url(0, 0, 0, 0, url, "Example", "Tooltip"));

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing an AutoFilter range with an active filter state
/// (`FILTERMODE`).
///
/// This is used to ensure the importer:
/// - preserves the AutoFilter dropdown range (best-effort), and
/// - surfaces a warning that filtered rows / criteria are not preserved.
pub fn build_autofilter_filtermode_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_filtermode_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_filtermode_fixture_xls`], but with at least
/// one row marked hidden via a `ROW` record.
///
/// When `FILTERMODE` is present, the importer should not preserve filtered-row visibility as
/// user-hidden rows.
pub fn build_autofilter_filtermode_hidden_rows_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_filtermode_hidden_rows_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_filtermode_fixture_xls`], but without a
/// `_xlnm._FilterDatabase` defined name.
///
/// This exercises the importer's best-effort AutoFilter range inference from the worksheet
/// substream (DIMENSIONS + AUTOFILTERINFO) when the canonical FilterDatabase name is missing.
pub fn build_autofilter_filtermode_no_filterdatabase_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_filtermode_no_filterdatabase_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a sheet with `FILTERMODE` but no `AUTOFILTERINFO` and no
/// `_xlnm._FilterDatabase` defined name.
///
/// Some producers may omit `AUTOFILTERINFO`; we still treat FILTERMODE as an indication that the
/// sheet had an active filter state and should therefore surface the warning and preserve the
/// dropdown range best-effort from DIMENSIONS.
pub fn build_filtermode_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_filtermode_only_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a sheet with `AUTOFILTERINFO` but no `FILTERMODE` and no
/// `_xlnm._FilterDatabase` defined name.
///
/// This exercises best-effort AutoFilter range inference from the worksheet substream when the
/// sheet contains an AutoFilter object but no currently-active filter state.
pub fn build_autofilterinfo_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilterinfo_only_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where the `_xlnm._FilterDatabase` NAME record uses a `PtgAreaN`
/// token (relative area reference) rather than `PtgArea`.
///
/// Some producers encode the AutoFilter range using BIFF8 relative-reference ptgs. The importer
/// should still recover the correct dropdown range.
pub fn build_autofilter_filterdatabase_arean_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_filterdatabase_arean_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with workbook calculation settings set to non-default values.
///
/// This is used to verify BIFF `CALCMODE`/`ITERATION`/`CALCCOUNT`/`DELTA`/`PRECISION`/`SAVERECALC`
/// records are imported into [`formula_model::Workbook::calc_settings`].
pub fn build_calc_settings_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_calc_settings_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

fn build_autofilter_filtermode_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    // One "cell" XF for NUMBER records.
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Filtered");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`).
    //
    // Use a smaller filter range than the sheet's DIMENSIONS bounding box so tests can verify we
    // prefer `_FilterDatabase` (canonical) over DIMENSIONS-based heuristics.
    let filter_db_rgce = ptg_area(0, 2, 0, 1); // $A$1:$B$3
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_autofilter_filtermode_sheet_stream(xf_cell);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_filtermode_hidden_rows_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    // One "cell" XF for NUMBER records.
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "FilteredHiddenRows");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`).
    //
    // Use a smaller filter range than the sheet's DIMENSIONS bounding box so tests can verify we
    // prefer `_FilterDatabase` (canonical) over DIMENSIONS-based heuristics.
    let filter_db_rgce = ptg_area(0, 2, 0, 1); // $A$1:$B$3
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_autofilter_filtermode_hidden_rows_sheet_stream(xf_cell);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_filtermode_no_filterdatabase_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    // One "cell" XF for NUMBER records.
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "FilteredNoDb");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_autofilter_filtermode_no_filterdatabase_sheet_stream(xf_cell);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_filtermode_no_filterdatabase_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 5) cols [0, 2) => A1:B5.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: 2 columns (A..B).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &2u16.to_le_bytes());
    // FILTERMODE indicates an active filter state (filtered rows).
    push_record(&mut sheet, RECORD_FILTERMODE, &[]);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_filtermode_only_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    // One "cell" XF for NUMBER records.
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "FilterModeOnly");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_filtermode_only_sheet_stream(xf_cell);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_filtermode_only_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 4) cols [0, 3) => A1:C4.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&4u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1 (A..C)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // FILTERMODE indicates an active filter state (filtered rows).
    push_record(&mut sheet, RECORD_FILTERMODE, &[]);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_autofilterinfo_only_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    // One "cell" XF for NUMBER records.
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "AutoFilterInfoOnly");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_autofilterinfo_only_sheet_stream(xf_cell);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilterinfo_only_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 4) cols [0, 4) => A1:D4.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&4u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&4u16.to_le_bytes()); // last col + 1 (A..D)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: 2 columns (A..B). No FILTERMODE present.
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &2u16.to_le_bytes());

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_autofilter_filterdatabase_arean_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    // One "cell" XF for NUMBER records.
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "AreaNFilterDb");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in id 0x0D) scoped to the sheet (`itab=1`), but encoded as a
    // relative-area token (PtgAreaN) rather than PtgArea. This should still decode to `A1:B3`.
    //
    // PtgAreaN token (ref class): [ptg=0x2D][rwFirst][rwLast][colFirst][colLast].
    // Row/col are relative when the corresponding bits are set in the col fields.
    let mut filter_db_rgce = Vec::<u8>::new();
    filter_db_rgce.push(0x2D); // PtgAreaN
    filter_db_rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst (offset 0)
    filter_db_rgce.extend_from_slice(&2u16.to_le_bytes()); // rwLast (offset 2 => row 3)
    let col_first_field: u16 = 0xC000 | 0; // A, row+col relative
    let col_last_field: u16 = 0xC000 | 1; // B, row+col relative
    filter_db_rgce.extend_from_slice(&col_first_field.to_le_bytes());
    filter_db_rgce.extend_from_slice(&col_last_field.to_le_bytes());
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_autofilter_filterdatabase_arean_sheet_stream(xf_cell);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_filterdatabase_arean_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 10) cols [0, 4) => A1:D10.
    // This intentionally exceeds the `_FilterDatabase` range so we can verify we prefer the defined
    // name over DIMENSIONS-based inference.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&10u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&4u16.to_le_bytes()); // last col + 1 (A..D)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: 4 columns (A..D). This will cause DIMENSIONS-based inference to yield A1:D10.
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &4u16.to_le_bytes());

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_autofilter_filtermode_hidden_rows_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 5) cols [0, 2) => A1:B5.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Mark row 2 (1-based) as hidden via the ROW record.
    // When FILTERMODE is present, the importer should *not* preserve this as a user-hidden row.
    push_record(&mut sheet, RECORD_ROW, &row_record(1, true, 0, false));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: 2 columns (A..B).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &2u16.to_le_bytes());
    // FILTERMODE indicates an active filter state (filtered rows).
    push_record(&mut sheet, RECORD_FILTERMODE, &[]);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_autofilter_filtermode_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 5) cols [0, 2) => A1:B5.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: 2 columns (A..B).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &2u16.to_le_bytes());
    // FILTERMODE indicates an active filter state (filtered rows).
    push_record(&mut sheet, RECORD_FILTERMODE, &[]);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

/// Build a BIFF8 `.xls` fixture with workbook and worksheet protection enabled.
///
/// This fixture includes:
/// - Workbook globals: `PROTECT`, `WINDOWPROTECT`, `PASSWORD`
/// - Worksheet: `PROTECT`, `PASSWORD` (plus object/scenario protection records)
pub fn build_protection_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_protection_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with workbook and worksheet protection records that include
/// truncated payloads.
///
/// This exercises the importer's best-effort behaviour: truncated record payloads should surface
/// as `ImportWarning`s, but parsing should continue and later valid records should still be
/// imported.
pub fn build_protection_truncated_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_protection_truncated_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with worksheet protection enabled and the richer "allow" flags
/// populated via BIFF8 FEAT/FEATHEADR records.
pub fn build_sheet_protection_allow_flags_fixture_xls() -> Vec<u8> {
    let sheet_stream = build_sheet_protection_allow_flags_sheet_stream(false);
    let workbook_stream =
        build_single_sheet_workbook_stream("ProtectedAllowFlags", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with worksheet protection enabled and the richer allow flags stored
/// in a `FEAT` record that is split across a `CONTINUE` record.
///
/// This exercises best-effort reassembly of continued FEAT records in the worksheet substream.
pub fn build_sheet_protection_allow_flags_feat_continued_fixture_xls() -> Vec<u8> {
    let sheet_stream = build_sheet_protection_allow_flags_feat_continued_sheet_stream();
    let workbook_stream =
        build_single_sheet_workbook_stream("ProtectedAllowFlagsFeatCont", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture like [`build_sheet_protection_allow_flags_fixture_xls`], but with a
/// deliberately malformed FEAT record that should produce a warning while still importing the final
/// allow flags.
pub fn build_sheet_protection_allow_flags_malformed_fixture_xls() -> Vec<u8> {
    let sheet_stream = build_sheet_protection_allow_flags_sheet_stream(true);
    let workbook_stream =
        build_single_sheet_workbook_stream("ProtectedAllowFlagsMalformed", &sheet_stream, 1252);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with a custom sheet tab color.
pub fn build_tab_color_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_tab_color_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where the sheet tab color is stored as an indexed palette color
/// (XColorType=1) and resolved through a workbook `PALETTE` record.
pub fn build_tab_color_indexed_palette_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_tab_color_indexed_palette_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that exercises workbook/worksheet view state import:
///
/// - Workbook WINDOW1 selects the second sheet tab.
/// - Second worksheet contains SCL (zoom), PANE (frozen first row/col), and SELECTION (active cell).
pub fn build_view_state_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_view_state_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that includes a non-default workbook window geometry/state in the
/// workbook-global `WINDOW1` record.
pub fn build_workbook_window_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_workbook_window_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with workbook window geometry/state populated in the WINDOW1 record.
///
/// This is used to validate `Workbook.view.window` import from BIFF.
pub fn build_window_geometry_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_window_geometry_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture with workbook window geometry/state populated in the WINDOW1 record,
/// using `WINDOW1.fHidden` to represent a hidden workbook window.
///
/// This is used to validate best-effort mapping of hidden windows to `WorkbookWindowState::Minimized`.
pub fn build_window_hidden_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_window_hidden_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that contains sheet-scoped built-in defined names (`NAME` records).
///
/// This is used to validate that the importer maps BIFF8 built-in name ids to the expected
/// Excel-visible `_xlnm.*` defined name strings, preserves the hidden flag, and imports the
/// correct sheet scope.
pub fn build_defined_names_builtins_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_defined_names_builtins_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture like [`build_defined_names_builtins_fixture_xls`], but with a
/// deliberate mismatch between the `NAME.chKey` byte and the built-in name id stored in
/// `NAME.rgchName`.
///
/// `chKey` is documented as a keyboard shortcut; when both fields are populated Excel appears to
/// prefer the built-in id stored in `rgchName` and treat `chKey` as a shortcut.
pub fn build_defined_names_builtins_chkey_mismatch_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_defined_names_builtins_chkey_mismatch_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a minimal BIFF8 `.xls` fixture containing a single sheet named `Filter` with a
/// sheet-scoped `_xlnm._FilterDatabase` defined name referencing `$A$1:$C$5`.
///
/// This exercises AutoFilter range import from legacy `.xls` files.
pub fn build_autofilter_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet named `FilterSort` with a
/// sheet-scoped `_xlnm._FilterDatabase` defined name referencing `$A$1:$C$5`, plus a BIFF8 `SORT`
/// record describing an AutoFilter sort state.
///
/// This exercises import of `SheetAutoFilter.sort_state` from legacy `.xls` files.
pub fn build_autofilter_sort_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_sort_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a minimal BIFF8 `.xls` fixture containing a single sheet with an AutoFilter range and a
/// BIFF8 Future Record Type `AutoFilter12` record.
///
/// Excel 2007+ can store filter criteria in legacy `.xls` files using `AutoFilter12` records. This
/// fixture is hand-crafted to ensure our importer does not panic when such records are present and
/// can recover at least one filter column when possible.
pub fn build_autofilter12_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter12_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a minimal BIFF8 `.xls` fixture containing a single sheet named `Filter` with a
/// **workbook-scoped** `_xlnm._FilterDatabase` defined name whose formula references `Filter!$A$1:$C$5`.
///
/// This fixture is intended to validate the importer’s calamine `Reader::defined_names()` fallback
/// path (i.e. when BIFF workbook parsing is unavailable and name scope metadata is missing).
pub fn build_autofilter_calamine_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_calamine_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing two sheets:
/// - `Calamine`: `_xlnm._FilterDatabase` stored as a *regular* workbook-scoped NAME (so calamine
///   surfaces it via `Reader::defined_names()`), and
/// - `Builtin`: `_xlnm._FilterDatabase` stored as a *built-in* sheet-scoped NAME (which calamine
///   may omit or mis-decode).
///
/// This is used to validate that, when BIFF workbook-stream parsing is unavailable, we can still
/// recover *missing* AutoFilter ranges from the workbook stream even if calamine surfaces some
/// FilterDatabase names.
pub fn build_autofilter_mixed_calamine_and_builtin_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_mixed_calamine_and_builtin_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture that exercises rich style import (FONT/XF/PALETTE).
///
/// The fixture contains a single sheet named `Styles` with a value cell (`A1`) that references
/// a rich XF record using:
/// - a non-default font (name/bold/italic/underline/strike/color)
/// - fill pattern + colors
/// - borders
/// - alignment
/// - protection flags
/// - a built-in number format
pub fn build_rich_styles_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_rich_styles_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet named `AutoFilter` with a
/// workbook-scoped `_xlnm._FilterDatabase` defined name.
///
/// Some `.xls` files in the wild store the FilterDatabase name with workbook scope (`itab==0`) and
/// reference the target sheet via a 3D token (`PtgArea3d`) that requires `EXTERNSHEET` resolution.
pub fn build_autofilter_workbook_scope_externsheet_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_workbook_scope_externsheet_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing two sheets:
/// - `Unqualified`: contains `AUTOFILTERINFO` so the importer infers an AutoFilter range from the
///   worksheet substream (DIMENSIONS + AUTOFILTERINFO), and
/// - `Other`: contains no AutoFilter records.
///
/// The workbook contains a *workbook-scoped* built-in `_xlnm._FilterDatabase` name whose formula is
/// an unqualified 2D range (e.g. `=$A$1:$B$3`).
///
/// Some `.xls` writers emit this form even for multi-sheet workbooks. This fixture validates the
/// importer's heuristic that, when the sheet scope is otherwise unknown, we attach such a
/// FilterDatabase range to the only sheet that already has AutoFilter metadata.
pub fn build_autofilter_workbook_scope_unqualified_multisheet_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_workbook_scope_unqualified_multisheet_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single sheet named `AutoFilter` with a
/// workbook-scoped `_FilterDatabase` defined name encoded as a *normal* (non-built-in) NAME string.
///
/// This fixture is specifically encoded so calamine surfaces the defined name via
/// `Reader::defined_names()`, exercising the importer's calamine-defined-name fallback path (when
/// BIFF parsing is unavailable).
pub fn build_autofilter_calamine_filterdatabase_alias_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_calamine_filterdatabase_alias_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

fn build_autofilter_workbook_scope_externsheet_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table. Keep the usual 16 style XFs so BIFF consumers stay happy.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "AutoFilter");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // Minimal EXTERNSHEET table with a single internal sheet entry.
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // Workbook-scoped _FilterDatabase name: Sheet0!$A$1:$C$5 (hidden).
    let filter_db_rgce = ptg_area3d(0, 0, 4, 0, 2);
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 0, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet_offset = globals.len();
    let sheet = build_autofilter_workbook_scope_externsheet_sheet_stream();

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_workbook_scope_unqualified_multisheet_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table. Keep the usual 16 style XFs so BIFF consumers stay happy.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Two worksheets.
    let boundsheet1_start = globals.len();
    let mut boundsheet1 = Vec::<u8>::new();
    boundsheet1.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet1.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet1, "Unqualified");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet1);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    let boundsheet2_start = globals.len();
    let mut boundsheet2 = Vec::<u8>::new();
    boundsheet2.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet2.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet2, "Other");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet2);
    let boundsheet2_offset_pos = boundsheet2_start + 4;

    // Workbook-scoped built-in `_FilterDatabase` name (hidden), but with a 2D PtgArea token that
    // does not specify a sheet.
    let filter_db_rgce = ptg_area(0, 2, 0, 1); // $A$1:$B$3
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 0, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet1_offset = globals.len();
    let sheet1 = build_autofilter_filterdatabase_arean_sheet_stream(xf_cell);
    let sheet2_offset = sheet1_offset + sheet1.len();
    let sheet2 = build_autofilter_sheet_stream_with_dimensions(xf_cell, 1, 1);

    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet1);
    globals.extend_from_slice(&sheet2);
    globals
}

fn build_autofilter_calamine_filterdatabase_alias_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table. Keep the usual 16 style XFs so BIFF consumers stay happy.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "AutoFilter");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // Minimal SUPBOOK entry for internal workbook references (calamine-compatible encoding).
    let supbook = {
        let mut data = Vec::<u8>::new();
        data.extend_from_slice(&1u16.to_le_bytes()); // ctab (sheet count)
        data.extend_from_slice(&1u16.to_le_bytes()); // cch
        data.push(0); // flags (compressed)
        data.push(0); // virtPath = "\0" (internal workbook marker)
        data
    };
    push_record(&mut globals, RECORD_SUPBOOK, &supbook);

    // Minimal EXTERNSHEET table with a single internal sheet entry.
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // Workbook-scoped `_FilterDatabase` name: AutoFilter!$A$1:$C$5.
    let filter_db_rgce = ptg_area3d(0, 0, 4, 0, 2);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record_calamine_compat("_FilterDatabase", &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet_offset = globals.len();
    let sheet = build_autofilter_calamine_filterdatabase_alias_sheet_stream(xf_general);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_workbook_scope_externsheet_sheet_stream() -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 5) cols [0, 3) (A..C), matching the AutoFilter range.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());
    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_autofilter_calamine_filterdatabase_alias_sheet_stream(xf_general: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 5) cols [0, 3) (A..C), matching the AutoFilter range.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the sheet has at least one cell so calamine reports a non-empty range.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 0.0),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_autofilter_mixed_calamine_and_builtin_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table (style XFs + one cell XF).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Two worksheets.
    let boundsheet1_start = globals.len();
    let mut boundsheet1 = Vec::<u8>::new();
    boundsheet1.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet1.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet1, "Calamine");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet1);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    let boundsheet2_start = globals.len();
    let mut boundsheet2 = Vec::<u8>::new();
    boundsheet2.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet2.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet2, "Builtin");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet2);
    let boundsheet2_offset_pos = boundsheet2_start + 4;

    // External reference tables so calamine can decode 3D references in the NAME formula stream.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0), (1, 1)]),
    );

    // Workbook-scoped string name `_xlnm._FilterDatabase` referencing Calamine!$A$1:$C$5.
    let calamine_rgce = ptg_area3d(0, 0, 4, 0, 2);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record(XLNM_FILTER_DATABASE, 0, false, None, &calamine_rgce),
    );

    // Built-in `_FilterDatabase` name scoped to the second sheet (itab=2): Builtin!$A$1:$B$3.
    let builtin_rgce = ptg_area(0, 2, 0, 1);
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 2, 0x0D, &builtin_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet substreams -------------------------------------------------------
    let sheet1_offset = globals.len();
    let sheet1 = build_autofilter_sheet_stream_with_dimensions(xf_cell, 5, 3); // A1:C5
    let sheet2_offset = sheet1_offset + sheet1.len();
    let sheet2 = build_autofilter_sheet_stream_with_dimensions(xf_cell, 3, 2); // A1:B3

    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet1);
    globals.extend_from_slice(&sheet2);
    globals
}

fn build_autofilter_sheet_stream_with_dimensions(
    xf_cell: u16,
    last_row_plus1: u32,
    last_col_plus1: u16,
) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, last_row_plus1) cols [0, last_col_plus1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&last_row_plus1.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&last_col_plus1.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_sheet_visibility_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: 16 style XFs + 1 default cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Three worksheets with differing BoundSheet `hsState` values.
    let bs_visible_start = globals.len();
    let mut bs_visible = Vec::<u8>::new();
    bs_visible.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    bs_visible.push(0x00); // hsState = visible
    bs_visible.push(0x00); // dt = worksheet
    write_short_unicode_string(&mut bs_visible, "Visible");
    push_record(&mut globals, RECORD_BOUNDSHEET, &bs_visible);
    let bs_visible_offset_pos = bs_visible_start + 4;

    let bs_hidden_start = globals.len();
    let mut bs_hidden = Vec::<u8>::new();
    bs_hidden.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    bs_hidden.push(0x01); // hsState = hidden
    bs_hidden.push(0x00); // dt = worksheet
    write_short_unicode_string(&mut bs_hidden, "Hidden");
    push_record(&mut globals, RECORD_BOUNDSHEET, &bs_hidden);
    let bs_hidden_offset_pos = bs_hidden_start + 4;

    let bs_very_hidden_start = globals.len();
    let mut bs_very_hidden = Vec::<u8>::new();
    bs_very_hidden.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    bs_very_hidden.push(0x02); // hsState = very hidden
    bs_very_hidden.push(0x00); // dt = worksheet
    write_short_unicode_string(&mut bs_very_hidden, "VeryHidden");
    push_record(&mut globals, RECORD_BOUNDSHEET, &bs_very_hidden);
    let bs_very_hidden_offset_pos = bs_very_hidden_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet1_offset = globals.len();
    let sheet1 = build_autofilter_sheet_stream_with_dimensions(xf_cell, 1, 1);
    let sheet2_offset = sheet1_offset + sheet1.len();
    let sheet2 = build_autofilter_sheet_stream_with_dimensions(xf_cell, 1, 1);
    let sheet3_offset = sheet2_offset + sheet2.len();
    let sheet3 = build_autofilter_sheet_stream_with_dimensions(xf_cell, 1, 1);

    // Patch BoundSheet offsets.
    globals[bs_visible_offset_pos..bs_visible_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals[bs_hidden_offset_pos..bs_hidden_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());
    globals[bs_very_hidden_offset_pos..bs_very_hidden_offset_pos + 4]
        .copy_from_slice(&(sheet3_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet1);
    globals.extend_from_slice(&sheet2);
    globals.extend_from_slice(&sheet3);
    globals
}

fn build_workbook_stream(date_1904: bool, include_filepass: bool) -> Vec<u8> {
    // -- Globals -----------------------------------------------------------------
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    if include_filepass {
        // FILEPASS indicates the workbook stream is encrypted/password-protected.
        //
        // In real encrypted workbooks, bytes after this record are encrypted. In our fixtures we
        // keep the stream plaintext to simulate a post-decryption buffer that still contains the
        // FILEPASS header.
        push_record(&mut globals, RECORD_FILEPASS, &[]);
    }
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(
        &mut globals,
        RECORD_DATEMODE,
        &(u16::from(date_1904)).to_le_bytes(),
    ); // DATEMODE: 1900 vs 1904 system
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1

    // Minimal font table.
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Custom formats. Excel typically allocates custom IDs >= 0x00A4.
    const FMT_CURRENCY: u16 = 0x00A4;
    const FMT_DATE: u16 = 0x00A5;
    push_record(
        &mut globals,
        RECORD_FORMAT,
        &format_record(FMT_CURRENCY, "$#,##0.00"),
    );
    push_record(
        &mut globals,
        RECORD_FORMAT,
        &format_record(FMT_DATE, "m/d/yy"),
    );

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Cell XFs.
    let xf_currency = 16u16;
    let xf_percent = 17u16;
    let xf_date = 18u16;
    let xf_time = 19u16;
    let xf_duration = 20u16;

    push_record(&mut globals, RECORD_XF, &xf_record(0, FMT_CURRENCY, false));
    push_record(&mut globals, RECORD_XF, &xf_record(0, 10, false)); // 0.00% (built-in)
    push_record(&mut globals, RECORD_XF, &xf_record(0, FMT_DATE, false));
    push_record(&mut globals, RECORD_XF, &xf_record(0, 21, false)); // h:mm:ss (built-in)
    push_record(&mut globals, RECORD_XF, &xf_record(0, 46, false)); // [h]:mm:ss (built-in)

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Formats");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_sheet_stream(xf_currency, xf_percent, xf_date, xf_time, xf_duration);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    // Append sheet substream.
    globals.extend_from_slice(&sheet);
    globals
}

fn build_encrypted_filepass_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();
    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    // FILEPASS indicates the workbook stream is encrypted/password-protected.
    // The payload layout depends on the encryption scheme; any bytes are fine for our fixture.
    push_record(&mut globals, RECORD_FILEPASS, &[]);
    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals
    globals
}

#[derive(Clone, Copy, Debug)]
enum NoteCommentSheetKind {
    SingleCell,
    MergedRegionNonAnchorNote,
}

fn build_note_comment_workbook_stream(kind: NoteCommentSheetKind) -> Vec<u8> {
    let (sheet_name, sheet_stream) = match kind {
        NoteCommentSheetKind::SingleCell => ("Notes", build_note_comment_sheet_stream(false)),
        NoteCommentSheetKind::MergedRegionNonAnchorNote => {
            ("MergedNotes", build_note_comment_sheet_stream(true))
        }
    };

    build_single_sheet_workbook_stream(sheet_name, &sheet_stream, 1252)
}

fn build_note_comment_biff5_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream_biff5(
        "NotesBiff5",
        &build_note_comment_biff5_sheet_stream(),
        1251,
    )
}

fn build_note_comment_biff5_split_across_continues_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream_biff5(
        "NotesBiff5Split",
        &build_note_comment_biff5_split_across_continues_sheet_stream(false),
        1251,
    )
}

fn build_note_comment_biff5_split_across_continues_with_flags_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream_biff5(
        "NotesBiff5SplitFlags",
        &build_note_comment_biff5_split_across_continues_sheet_stream(true),
        1251,
    )
}

fn build_note_comment_biff5_split_across_continues_codepage_932_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream_biff5(
        "NotesBiff5SplitCp932",
        &build_note_comment_biff5_split_across_continues_codepage_932_sheet_stream(),
        932,
    )
}

fn build_note_comment_biff5_author_biff8_short_string_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream_biff5(
        "NotesBiff5AuthorBiff8",
        &build_note_comment_biff5_author_biff8_short_string_sheet_stream(),
        1251,
    )
}

fn build_note_comment_biff5_author_biff8_unicode_string_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream_biff5(
        "NotesBiff5AuthorBiff8Unicode",
        &build_note_comment_biff5_author_biff8_unicode_string_sheet_stream(),
        1252,
    )
}

fn build_note_comment_biff5_in_merged_region_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream_biff5(
        "MergedNotesBiff5",
        &build_note_comment_biff5_in_merged_region_sheet_stream(),
        1252,
    )
}

fn build_note_comment_biff5_missing_txo_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream_biff5(
        "NotesBiff5MissingTxo",
        &build_note_comment_biff5_missing_txo_sheet_stream(),
        1252,
    )
}

fn build_note_comment_codepage_1251_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesCp1251",
        &build_note_comment_codepage_1251_sheet_stream(),
        1251,
    )
}

fn build_note_comment_author_codepage_1251_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesAuthorCp1251",
        &build_note_comment_author_codepage_1251_sheet_stream(),
        1251,
    )
}

fn build_note_comment_author_xl_unicode_string_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesAuthorXlUnicode",
        &build_note_comment_author_xl_unicode_string_sheet_stream(),
        1252,
    )
}

fn build_note_comment_author_missing_flags_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesAuthorNoFlags",
        &build_note_comment_author_missing_flags_sheet_stream(),
        1252,
    )
}

fn build_note_comment_txo_text_missing_flags_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesTxoTextNoFlags",
        &build_note_comment_txo_text_missing_flags_sheet_stream(),
        1252,
    )
}

fn build_note_comment_txo_text_missing_flags_in_second_fragment_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesTxoTextNoFlagsMid",
        &build_note_comment_txo_text_missing_flags_in_second_fragment_sheet_stream(),
        1252,
    )
}

fn build_note_comment_txo_cch_text_zero_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesTxoCchZero",
        &build_note_comment_txo_cch_text_zero_sheet_stream(),
        1252,
    )
}

fn build_note_comment_txo_cch_text_zero_missing_flags_in_second_fragment_workbook_stream() -> Vec<u8>
{
    build_single_sheet_workbook_stream(
        "NotesTxoCchZeroNoFlagsMid",
        &build_note_comment_txo_cch_text_zero_missing_flags_in_second_fragment_sheet_stream(),
        1252,
    )
}

fn build_note_comment_split_across_continues_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesSplit",
        &build_note_comment_split_across_continues_sheet_stream(),
        1252,
    )
}

fn build_note_comment_split_across_continues_mixed_encoding_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesSplitMixed",
        &build_note_comment_split_across_continues_mixed_encoding_sheet_stream(),
        1252,
    )
}

fn build_note_comment_split_across_continues_codepage_932_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesSplitCp932",
        &build_note_comment_split_across_continues_codepage_932_sheet_stream(),
        932,
    )
}

fn build_note_comment_split_across_continues_codepage_65001_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesSplitUtf8",
        &build_note_comment_split_across_continues_codepage_65001_sheet_stream(),
        65001,
    )
}

fn build_note_comment_split_utf16_code_unit_across_continues_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesSplitUtf16Odd",
        &build_note_comment_split_utf16_code_unit_across_continues_sheet_stream(),
        1252,
    )
}

fn build_note_comment_missing_txo_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesMissingTxo",
        &build_note_comment_missing_txo_sheet_stream(),
        1252,
    )
}

fn build_note_comment_missing_txo_header_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesMissingTxoHeader",
        &build_note_comment_missing_txo_header_sheet_stream(),
        1252,
    )
}

fn build_note_comment_missing_txo_header_missing_flags_in_second_fragment_workbook_stream(
) -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesMissingTxoHeaderNoFlagsMid",
        &build_note_comment_missing_txo_header_missing_flags_in_second_fragment_sheet_stream(),
        1252,
    )
}

fn build_note_comment_truncated_txo_header_missing_cb_runs_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesTxoHeaderNoCbRuns",
        &build_note_comment_truncated_txo_header_missing_cb_runs_sheet_stream(),
        1252,
    )
}

fn build_note_comment_txo_cch_text_offset_4_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesTxoCchOffset4",
        &build_note_comment_txo_cch_text_offset_4_sheet_stream(),
        1252,
    )
}

fn build_note_comment_note_obj_id_swapped_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "NotesObjIdSwapped",
        &build_note_comment_note_obj_id_swapped_sheet_stream(),
        1252,
    )
}

fn build_single_sheet_workbook_stream(
    sheet_name: &str,
    sheet_stream: &[u8],
    codepage: u16,
) -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &codepage.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One default cell XF (General).
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, sheet_name);
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet_offset = globals.len();
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(sheet_stream);
    globals
}

fn build_single_sheet_workbook_stream_biff5(
    sheet_name: &str,
    sheet_stream: &[u8],
    codepage: u16,
) -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(
        &mut globals,
        // Note: BIFF5 fixtures use BOF record id 0x0809 with a BIFF5 version (0x0500) in the BOF
        // payload because calamine accepts this layout for minimal test workbooks.
        RECORD_BOF,
        &bof_biff5(BOF_DT_WORKBOOK_GLOBALS),
    );
    push_record(&mut globals, RECORD_CODEPAGE, &codepage.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font_biff5("Arial"));

    // Many readers expect at least 16 style XFs before cell XFs. BIFF5 XF records are 16 bytes.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record_biff5(0, 0, true));
    }
    // One default cell XF (General).
    push_record(&mut globals, RECORD_XF, &xf_record_biff5(0, 0, false));

    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_ansi_string(&mut boundsheet, sheet_name);
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet_offset = globals.len();
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(sheet_stream);
    globals
}

fn build_note_comment_sheet_stream(include_merged_region: bool) -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello from note";

    let segments: [&[u8]; 1] = [TEXT.as_bytes()];
    build_note_comment_sheet_stream_with_compressed_txo(
        include_merged_region,
        OBJECT_ID,
        AUTHOR,
        &segments,
    )
}

fn build_note_comment_biff5_sheet_stream() -> Vec<u8> {
    // In Windows-1251, 0xC0 maps to Cyrillic "А" (U+0410).
    let author_bytes = [0xC0u8];
    let text_bytes = [b'H', b'i', b' ', 0xC0u8];
    let segments: [&[u8]; 1] = [&text_bytes];
    build_note_comment_biff5_sheet_stream_with_ansi_txo(&author_bytes, &segments, false, false)
}

fn build_note_comment_biff5_split_across_continues_sheet_stream(prefix_flags: bool) -> Vec<u8> {
    // In Windows-1251, 0xC0 maps to Cyrillic "А" (U+0410).
    let author_bytes = [0xC0u8];
    let part1 = [b'H', b'i', b' '];
    let part2 = [0xC0u8];
    let segments: [&[u8]; 2] = [&part1, &part2];
    build_note_comment_biff5_sheet_stream_with_ansi_txo(
        &author_bytes,
        &segments,
        prefix_flags,
        false,
    )
}

fn build_note_comment_biff5_split_across_continues_codepage_932_sheet_stream() -> Vec<u8> {
    // In Shift-JIS (codepage 932), '\u{3042}' ('あ') is encoded as 0x82 0xA0. Split across two
    // `CONTINUE` records as 0x82 + 0xA0 so we exercise decoding across record boundaries.
    let author_bytes = [0x82u8, 0xA0u8];
    let part1 = [0x82u8];
    let part2 = [0xA0u8];
    let segments: [&[u8]; 2] = [&part1, &part2];
    build_note_comment_biff5_sheet_stream_with_ansi_txo(&author_bytes, &segments, false, false)
}

fn build_note_comment_biff5_sheet_stream_with_ansi_txo(
    author_bytes: &[u8],
    text_segments: &[&[u8]],
    prefix_flags: bool,
    author_is_biff8_short_string: bool,
) -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    // The workbook globals above create 16 style XFs + 1 cell XF, so the first usable
    // cell XF index is 16.
    const XF_GENERAL_CELL: u16 = 16;

    let cch_text: u16 = text_segments
        .iter()
        .map(|seg| seg.len())
        .sum::<usize>()
        .try_into()
        .expect("comment text too long for u16 length");

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof_biff5(BOF_DT_WORKSHEET));

    // DIMENSIONS (BIFF5): [rwMic:u16][rwMac:u16][colMic:u16][colMac:u16][reserved:u16]
    // rows [0, 1), cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    // WINDOW2 isn't required for comment parsing, but helps ensure readers create a worksheet view.
    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    let note = if author_is_biff8_short_string {
        note_record_biff5_author_biff8_short_string_bytes(0u16, 0u16, OBJECT_ID, author_bytes)
    } else {
        note_record_biff5_author_bytes(0u16, 0u16, OBJECT_ID, author_bytes)
    };
    push_record(&mut sheet, RECORD_NOTE, &note);
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header: cchText at offset 6, cbRuns at offset 12.
    let mut txo = [0u8; 18];
    txo[6..8].copy_from_slice(&cch_text.to_le_bytes());
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE records: BIFF5 typically stores raw bytes, but some producers appear to prefix
    // each fragment with a BIFF8-style 0/1 option byte. The parser should handle both.
    for &seg in text_segments {
        if prefix_flags {
            let mut cont = Vec::<u8>::with_capacity(1 + seg.len());
            cont.push(0); // compressed 8-bit fragment
            cont.extend_from_slice(seg);
            push_record(&mut sheet, RECORD_CONTINUE, &cont);
        } else {
            push_record(&mut sheet, RECORD_CONTINUE, seg);
        }
    }

    // Formatting runs continuation (dummy bytes).
    push_record(&mut sheet, RECORD_CONTINUE, &[0u8; 4]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_biff5_author_biff8_short_string_sheet_stream() -> Vec<u8> {
    // In Windows-1251, 0xC0 maps to Cyrillic "А" (U+0410). Encode the author as a BIFF8-style
    // ShortXLUnicodeString ([len][flags][chars]) even though the workbook is BIFF5.
    let author_bytes = [0xC0u8];
    let text_bytes = [b'H', b'i'];
    let segments: [&[u8]; 1] = [&text_bytes];
    build_note_comment_biff5_sheet_stream_with_ansi_txo(&author_bytes, &segments, false, true)
}

fn build_note_comment_biff5_author_biff8_unicode_string_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const XF_GENERAL_CELL: u16 = 16;

    let author = "\u{0410}"; // Cyrillic "А"
    let text_bytes = [b'H', b'i'];
    let cch_text: u16 = text_bytes
        .len()
        .try_into()
        .expect("comment text too long for u16 length");

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof_biff5(BOF_DT_WORKSHEET));

    // DIMENSIONS (BIFF5): rows [0, 1), cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    // NOTE author stored as BIFF8 XLUnicodeString (u16 length + flags byte).
    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record_biff5_author_biff8_unicode_string(0u16, 0u16, OBJECT_ID, author),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header: cchText at offset 6, cbRuns at offset 12.
    let mut txo = [0u8; 18];
    txo[6..8].copy_from_slice(&cch_text.to_le_bytes());
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(&mut sheet, RECORD_TXO, &txo);

    push_record(&mut sheet, RECORD_CONTINUE, &text_bytes);
    push_record(&mut sheet, RECORD_CONTINUE, &[0u8; 4]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_biff5_in_merged_region_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const XF_GENERAL_CELL: u16 = 16;
    const AUTHOR: &str = "Alice";
    const TEXT: &[u8] = b"Hello";

    let cch_text: u16 = TEXT
        .len()
        .try_into()
        .expect("comment text too long for u16 length");

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof_biff5(BOF_DT_WORKSHEET));

    // DIMENSIONS (BIFF5): rows [0, 1), cols [0, 2) to cover A1..B1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&2u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // MERGEDCELLS: 1 range, A1:B1.
    let mut merged = Vec::<u8>::new();
    merged.extend_from_slice(&1u16.to_le_bytes()); // cAreas
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwLast
    merged.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A)
    merged.extend_from_slice(&1u16.to_le_bytes()); // colLast (B)
    push_record(&mut sheet, RECORD_MERGEDCELLS, &merged);

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    // NOTE record targets B1 (non-anchor) while A1:B1 is merged.
    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record_biff5_author_bytes(0u16, 1u16, OBJECT_ID, AUTHOR.as_bytes()),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header: cchText at offset 6, cbRuns at offset 12.
    let mut txo = [0u8; 18];
    txo[6..8].copy_from_slice(&cch_text.to_le_bytes());
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE: raw bytes.
    push_record(&mut sheet, RECORD_CONTINUE, TEXT);
    // Formatting runs continuation (dummy bytes).
    push_record(&mut sheet, RECORD_CONTINUE, &[0u8; 4]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_biff5_missing_txo_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const XF_GENERAL_CELL: u16 = 16;
    const AUTHOR: &str = "Alice";

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof_biff5(BOF_DT_WORKSHEET));

    // DIMENSIONS (BIFF5): rows [0, 1), cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record_biff5_author_bytes(0u16, 0u16, OBJECT_ID, AUTHOR.as_bytes()),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // Intentionally omit TXO/CONTINUE records.
    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_codepage_1251_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";

    // In Windows-1251, 0xC0 maps to Cyrillic "А" (U+0410).
    let text = [0xC0u8];
    let segments: [&[u8]; 1] = [&text];
    build_note_comment_sheet_stream_with_compressed_txo(false, OBJECT_ID, AUTHOR, &segments)
}

fn build_note_comment_author_codepage_1251_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const XF_GENERAL_CELL: u16 = 16;

    // In Windows-1251, 0xC0 maps to Cyrillic "А" (U+0410).
    let author_bytes = [0xC0u8];
    let segments: [&[u8]; 1] = [b"Hello"];
    let cch_text: u16 = segments
        .iter()
        .map(|seg| seg.len())
        .sum::<usize>()
        .try_into()
        .expect("comment text too long for u16 length");

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    // NOTE author stored as raw cp1251 bytes.
    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record_author_bytes(0u16, 0u16, OBJECT_ID, &author_bytes),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));
    push_txo_logical_record_compressed_segments(&mut sheet, cch_text, &segments);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_author_xl_unicode_string_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    // NOTE author stored as an XLUnicodeString (u16 length) rather than a ShortXLUnicodeString.
    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record_with_xl_unicode_author(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));
    push_txo_logical_record(&mut sheet, TEXT);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_author_missing_flags_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    // NOTE author stored as a BIFF5-style short string (no flags byte).
    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record_author_bytes_without_flags(0u16, 0u16, OBJECT_ID, AUTHOR.as_bytes()),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));
    push_txo_logical_record(&mut sheet, TEXT);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_txo_text_missing_flags_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header: cchText at offset 6, cbRuns at offset 12.
    let mut txo = [0u8; 18];
    txo[6..8].copy_from_slice(&(TEXT.len() as u16).to_le_bytes());
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE: text bytes only (missing the BIFF8 flags byte).
    push_record(&mut sheet, RECORD_CONTINUE, TEXT.as_bytes());

    // Formatting runs continuation.
    push_record(&mut sheet, RECORD_CONTINUE, &[0u8; 4]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_txo_text_missing_flags_in_second_fragment_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header: cchText at offset 6, cbRuns at offset 12.
    let mut txo = [0u8; 18];
    txo[6..8].copy_from_slice(&(TEXT.len() as u16).to_le_bytes());
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE #1: BIFF8 flags byte + first two chars.
    let mut cont1 = Vec::<u8>::new();
    cont1.push(0); // flags: compressed 8-bit chars
    cont1.extend_from_slice(b"He");
    push_record(&mut sheet, RECORD_CONTINUE, &cont1);

    // CONTINUE #2: remaining chars *without* a flags byte.
    push_record(&mut sheet, RECORD_CONTINUE, b"llo");

    // Formatting runs continuation.
    push_record(&mut sheet, RECORD_CONTINUE, &[0u8; 4]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_txo_cch_text_zero_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header with `cchText=0` but a non-zero `cbRuns`. The importer should still decode the
    // text from the continuation area.
    let mut txo = [0u8; 18];
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE: [flags: u8][chars...]
    let mut cont = Vec::<u8>::new();
    cont.push(0); // flags: compressed 8-bit chars
    cont.extend_from_slice(TEXT.as_bytes());
    push_record(&mut sheet, RECORD_CONTINUE, &cont);

    // Formatting runs continuation.
    push_record(&mut sheet, RECORD_CONTINUE, &[0u8; 4]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_txo_cch_text_zero_missing_flags_in_second_fragment_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header with `cchText=0` but a non-zero `cbRuns`.
    let mut txo = [0u8; 18];
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE #1: BIFF8 flags byte + first two chars.
    let mut cont1 = Vec::<u8>::new();
    cont1.push(0); // flags: compressed 8-bit chars
    cont1.extend_from_slice(b"He");
    push_record(&mut sheet, RECORD_CONTINUE, &cont1);

    // CONTINUE #2: remaining chars *without* a flags byte.
    let text_bytes = TEXT.as_bytes();
    push_record(&mut sheet, RECORD_CONTINUE, &text_bytes[2..]);

    // Formatting runs continuation.
    push_record(&mut sheet, RECORD_CONTINUE, &[0u8; 4]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_split_across_continues_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";

    // "ABCDE" split as "AB" + "CDE" across two `CONTINUE` records.
    let segments: [&[u8]; 2] = [b"AB", b"CDE"];
    build_note_comment_sheet_stream_with_compressed_txo(false, OBJECT_ID, AUTHOR, &segments)
}

fn build_note_comment_split_across_continues_mixed_encoding_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const XF_GENERAL_CELL: u16 = 16;
    let cch_text: u16 = 5; // "ABCDE"

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header: cchText at offset 6, cbRuns at offset 12.
    let mut txo = [0u8; 18];
    txo[6..8].copy_from_slice(&cch_text.to_le_bytes());
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE #1: compressed bytes "AB"
    let mut cont1 = Vec::<u8>::new();
    cont1.push(0); // flags: compressed 8-bit chars
    cont1.extend_from_slice(b"AB");
    push_record(&mut sheet, RECORD_CONTINUE, &cont1);

    // CONTINUE #2: UTF-16LE bytes "CDE"
    let mut cont2 = Vec::<u8>::new();
    cont2.push(0x01); // flags: fHighByte=1 (UTF-16LE)
    cont2.extend_from_slice(&[b'C', 0x00, b'D', 0x00, b'E', 0x00]);
    push_record(&mut sheet, RECORD_CONTINUE, &cont2);

    // Formatting runs continuation.
    push_record(&mut sheet, RECORD_CONTINUE, &[0u8; 4]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_split_across_continues_codepage_932_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";

    // "あ" in Shift-JIS is 0x82 0xA0. Split across two `CONTINUE` records as 0x82 + 0xA0.
    let part1 = [0x82u8];
    let part2 = [0xA0u8];
    let segments: [&[u8]; 2] = [&part1, &part2];

    build_note_comment_sheet_stream_with_compressed_txo(false, OBJECT_ID, AUTHOR, &segments)
}

fn build_note_comment_split_across_continues_codepage_65001_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";

    // "€" is 0xE2 0x82 0xAC in UTF-8. Split across three `CONTINUE` records.
    let part1 = [0xE2u8];
    let part2 = [0x82u8];
    let part3 = [0xACu8];
    let segments: [&[u8]; 3] = [&part1, &part2, &part3];

    build_note_comment_sheet_stream_with_compressed_txo(false, OBJECT_ID, AUTHOR, &segments)
}

fn build_note_comment_split_utf16_code_unit_across_continues_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const XF_GENERAL_CELL: u16 = 16;
    const CCH_TEXT: u16 = 1; // one Unicode character

    // '€' (U+20AC) is 0xAC 0x20 in UTF-16LE. Split across two CONTINUE records as 0xAC + 0x20.
    let part1 = [0xACu8];
    let part2 = [0x20u8];

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header: cchText at offset 6, cbRuns=0 (no formatting runs).
    let mut txo = [0u8; 18];
    txo[6..8].copy_from_slice(&CCH_TEXT.to_le_bytes());
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE #1: flags=1 (UTF-16LE), then first byte of the code unit.
    let mut cont1 = Vec::<u8>::new();
    cont1.push(0x01);
    cont1.extend_from_slice(&part1);
    push_record(&mut sheet, RECORD_CONTINUE, &cont1);

    // CONTINUE #2: flags=1 (UTF-16LE), then second byte of the code unit.
    let mut cont2 = Vec::<u8>::new();
    cont2.push(0x01);
    cont2.extend_from_slice(&part2);
    push_record(&mut sheet, RECORD_CONTINUE, &cont2);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_missing_txo_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";

    // The workbook globals above create 16 style XFs + 1 cell XF, so the first usable
    // cell XF index is 16.
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    // NOTE + OBJ but no TXO: importer should warn and skip.
    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_missing_txo_header_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO record with an empty header, forcing best-effort fallback decoding from CONTINUE fragments.
    push_record(&mut sheet, RECORD_TXO, &[]);

    let mut cont = Vec::<u8>::new();
    cont.push(0); // flags: compressed 8-bit chars
    cont.extend_from_slice(TEXT.as_bytes());
    push_record(&mut sheet, RECORD_CONTINUE, &cont);

    // Formatting runs continuation. Unlike continued string fragments, this payload does **not**
    // have a leading flags byte. When the TXO header is missing, our fallback decoder should stop
    // before decoding these bytes as text.
    push_record(&mut sheet, RECORD_CONTINUE, &[0x00, 0x00, 0x01, 0x00]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_missing_txo_header_missing_flags_in_second_fragment_sheet_stream() -> Vec<u8>
{
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO record with an empty header, forcing best-effort fallback decoding from CONTINUE fragments.
    push_record(&mut sheet, RECORD_TXO, &[]);

    // CONTINUE #1: BIFF8 flags byte + first two chars.
    let mut cont1 = Vec::<u8>::new();
    cont1.push(0); // flags: compressed 8-bit chars
    cont1.extend_from_slice(b"He");
    push_record(&mut sheet, RECORD_CONTINUE, &cont1);

    // CONTINUE #2: remaining chars *without* a flags byte.
    let text_bytes = TEXT.as_bytes();
    push_record(&mut sheet, RECORD_CONTINUE, &text_bytes[2..]);

    // Formatting runs continuation. Unlike continued string fragments, this payload does **not**
    // have a leading flags byte. Our fallback decoder should stop before decoding these bytes as text.
    push_record(&mut sheet, RECORD_CONTINUE, &[0x00, 0x00, 0x01, 0x00]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_truncated_txo_header_missing_cb_runs_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header (truncated) containing only `cchText` at offset 6. `cbRuns` (offset 12) is
    // missing, so the parser must avoid decoding formatting runs via heuristics.
    let mut txo = [0u8; 8];
    txo[6..8].copy_from_slice(&10u16.to_le_bytes()); // cchText larger than actual text bytes
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE #1: compressed bytes "Hello"
    let mut cont1 = Vec::<u8>::new();
    cont1.push(0); // flags: compressed 8-bit chars
    cont1.extend_from_slice(b"Hello");
    push_record(&mut sheet, RECORD_CONTINUE, &cont1);

    // CONTINUE #2: formatting runs (no leading flags byte): [ich=0][ifnt=1]
    push_record(&mut sheet, RECORD_CONTINUE, &[0x00, 0x00, 0x01, 0x00]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_txo_cch_text_offset_4_sheet_stream() -> Vec<u8> {
    const OBJECT_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hi";
    const XF_GENERAL_CELL: u16 = 16;

    let cch_text: u16 = TEXT
        .len()
        .try_into()
        .expect("comment text too long for u16 length");

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(0u16, 0u16, OBJECT_ID, AUTHOR),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(OBJECT_ID));

    // TXO header (18 bytes) with `cchText` at offset 4 (non-standard) and `cbRuns` at offset 12.
    let mut txo = [0u8; 18];
    txo[4..6].copy_from_slice(&cch_text.to_le_bytes());
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(&mut sheet, RECORD_TXO, &txo);

    // CONTINUE: [flags: u8][chars...]
    let mut cont = Vec::<u8>::new();
    cont.push(0); // flags: compressed 8-bit chars
    cont.extend_from_slice(TEXT.as_bytes());
    push_record(&mut sheet, RECORD_CONTINUE, &cont);

    // Formatting runs continuation.
    push_record(&mut sheet, RECORD_CONTINUE, &[0u8; 4]);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_note_obj_id_swapped_sheet_stream() -> Vec<u8> {
    const OBJ_ID_IN_OBJ_TCO: u16 = 2;
    const NOTE_PRIMARY_OBJ_ID: u16 = 1;
    const AUTHOR: &str = "Alice";
    const TEXT: &str = "Hello from swapped obj id";
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&1u16.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Ensure the anchor cell exists in the calamine value grid.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    // NOTE record with mismatched candidate object ids:
    // - bytes[6..8] (primary) doesn't match any TXO payload
    // - bytes[4..6] (secondary) matches the OBJ/TXO id (simulating a swapped field ordering).
    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record_with_obj_id_candidates(
            0u16,
            0u16,
            /*secondary_obj_id=*/ OBJ_ID_IN_OBJ_TCO,
            /*primary_obj_id=*/ NOTE_PRIMARY_OBJ_ID,
            AUTHOR,
        ),
    );
    push_record(
        &mut sheet,
        RECORD_OBJ,
        &obj_record_with_ftcmo(OBJ_ID_IN_OBJ_TCO),
    );
    push_txo_logical_record(&mut sheet, TEXT);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_note_comment_sheet_stream_with_compressed_txo(
    include_merged_region: bool,
    object_id: u16,
    author: &str,
    text_segments: &[&[u8]],
) -> Vec<u8> {
    let (note_row, note_col) = if include_merged_region {
        // NOTE targets B1 (non-anchor) while A1:B1 is merged.
        (0u16, 1u16)
    } else {
        // NOTE targets A1.
        (0u16, 0u16)
    };

    // `cchText` in TXO is a u16 character count; our fixtures use BIFF8 compressed 8-bit text, so
    // bytes == chars and we can sum segment byte lengths.
    let cch_text: u16 = text_segments
        .iter()
        .map(|seg| seg.len())
        .sum::<usize>()
        .try_into()
        .expect("comment text too long for u16 length");

    // The workbook globals above create 16 style XFs + 1 cell XF, so the first usable
    // cell XF index is 16.
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1) or [0, 2) if we include B1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes());
    dims.extend_from_slice(&1u32.to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    dims.extend_from_slice(&(if include_merged_region { 2u16 } else { 1u16 }).to_le_bytes());
    dims.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    if include_merged_region {
        // MERGEDCELLS: 1 range, A1:B1.
        let mut merged = Vec::<u8>::new();
        merged.extend_from_slice(&1u16.to_le_bytes()); // cAreas
        merged.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        merged.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        merged.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A)
        merged.extend_from_slice(&1u16.to_le_bytes()); // colLast (B)
        push_record(&mut sheet, RECORD_MERGEDCELLS, &merged);
    }

    // Ensure the anchor cell exists in the calamine value grid (even though the comment could
    // conceptually be attached to an empty cell).
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, XF_GENERAL_CELL));

    // NOTE/OBJ/TXO trio that encodes the comment payload.
    push_record(
        &mut sheet,
        RECORD_NOTE,
        &note_record(note_row, note_col, object_id, author),
    );
    push_record(&mut sheet, RECORD_OBJ, &obj_record_with_ftcmo(object_id));
    push_txo_logical_record_compressed_segments(&mut sheet, cch_text, text_segments);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn note_record(row: u16, col: u16, object_id: u16, author: &str) -> Vec<u8> {
    // NOTE record (BIFF8): [rw: u16][col: u16][grbit: u16][idObj: u16][stAuthor]
    //
    // Some parsers differ on whether `idObj` precedes `grbit`; we keep this fixture robust
    // by writing the same stable value into both fields (object_id=1).
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&object_id.to_le_bytes()); // grbit (or idObj)
    out.extend_from_slice(&object_id.to_le_bytes()); // idObj (or grbit)
    write_short_unicode_string(&mut out, author);
    out
}

fn note_record_with_xl_unicode_author(row: u16, col: u16, object_id: u16, author: &str) -> Vec<u8> {
    // NOTE record (BIFF8): [rw: u16][col: u16][grbit: u16][idObj: u16][stAuthor]
    //
    // `stAuthor` should normally be a ShortXLUnicodeString, but some producers store an
    // XLUnicodeString (u16 length) instead.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&object_id.to_le_bytes()); // grbit (or idObj)
    out.extend_from_slice(&object_id.to_le_bytes()); // idObj (or grbit)
    write_unicode_string(&mut out, author);
    out
}

fn note_record_with_obj_id_candidates(
    row: u16,
    col: u16,
    secondary_obj_id: u16,
    primary_obj_id: u16,
    author: &str,
) -> Vec<u8> {
    // NOTE record (BIFF8): [rw: u16][col: u16][grbit: u16][idObj: u16][stAuthor]
    //
    // Our parser treats the field at offset 6 as the primary `obj_id` and the field at offset 4 as
    // the secondary candidate (some files appear to swap them).
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&secondary_obj_id.to_le_bytes()); // grbit (or idObj)
    out.extend_from_slice(&primary_obj_id.to_le_bytes()); // idObj (or grbit)
    write_short_unicode_string(&mut out, author);
    out
}

fn note_record_author_bytes(row: u16, col: u16, object_id: u16, author_bytes: &[u8]) -> Vec<u8> {
    // NOTE record (BIFF8): [rw: u16][col: u16][grbit: u16][idObj: u16][stAuthor]
    //
    // `stAuthor` is a ShortXLUnicodeString; for compressed strings the `chars` payload is raw
    // codepage bytes. This helper lets fixtures supply non-UTF8 bytes directly.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&object_id.to_le_bytes()); // grbit (or idObj)
    out.extend_from_slice(&object_id.to_le_bytes()); // idObj (or grbit)
    let len: u8 = author_bytes
        .len()
        .try_into()
        .expect("author string too long for u8 length");
    out.push(len);
    out.push(0); // compressed (8-bit)
    out.extend_from_slice(author_bytes);
    out
}

fn note_record_author_bytes_without_flags(
    row: u16,
    col: u16,
    object_id: u16,
    author_bytes: &[u8],
) -> Vec<u8> {
    // NOTE record (BIFF8): [rw: u16][col: u16][grbit: u16][idObj: u16][stAuthor]
    //
    // Nonstandard variant: `stAuthor` is stored as a BIFF5-style ANSI short string (no BIFF8
    // option flags byte):
    //   [cch: u8][rgb: u8 * cch]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&object_id.to_le_bytes()); // grbit (or idObj)
    out.extend_from_slice(&object_id.to_le_bytes()); // idObj (or grbit)
    write_short_ansi_bytes(&mut out, author_bytes);
    out
}

fn note_record_biff5_author_bytes(
    row: u16,
    col: u16,
    object_id: u16,
    author_bytes: &[u8],
) -> Vec<u8> {
    // NOTE record (BIFF5): [rw: u16][col: u16][grbit: u16][idObj: u16][stAuthor]
    //
    // `stAuthor` is stored as an ANSI short string:
    //   [cch: u8][rgb: u8 * cch]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&object_id.to_le_bytes()); // grbit (or idObj)
    out.extend_from_slice(&object_id.to_le_bytes()); // idObj (or grbit)
    write_short_ansi_bytes(&mut out, author_bytes);
    out
}

fn note_record_biff5_author_biff8_short_string_bytes(
    row: u16,
    col: u16,
    object_id: u16,
    author_bytes: &[u8],
) -> Vec<u8> {
    // NOTE record (BIFF5) with a non-standard BIFF8 ShortXLUnicodeString author encoding:
    //   [rw:u16][col:u16][grbit:u16][idObj:u16][cch:u8][flags:u8][chars...]
    //
    // Some `.xls` producers appear to use this encoding even in BIFF5 workbooks.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&object_id.to_le_bytes()); // grbit (or idObj)
    out.extend_from_slice(&object_id.to_le_bytes()); // idObj (or grbit)
    let len: u8 = author_bytes
        .len()
        .try_into()
        .expect("author string too long for u8 length");
    out.push(len);
    out.push(0); // flags: compressed 8-bit chars
    out.extend_from_slice(author_bytes);
    out
}

fn note_record_biff5_author_biff8_unicode_string(
    row: u16,
    col: u16,
    object_id: u16,
    author: &str,
) -> Vec<u8> {
    // NOTE record (BIFF5) with a non-standard BIFF8 XLUnicodeString author encoding:
    //   [rw:u16][col:u16][grbit:u16][idObj:u16][cch:u16][flags:u8][chars...]
    //
    // We encode the author as UTF-16LE (fHighByte=1) so it round-trips independently of workbook
    // codepage.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&object_id.to_le_bytes()); // grbit (or idObj)
    out.extend_from_slice(&object_id.to_le_bytes()); // idObj (or grbit)

    let utf16: Vec<u16> = author.encode_utf16().collect();
    let len: u16 = utf16
        .len()
        .try_into()
        .expect("author string too long for u16 length");
    out.extend_from_slice(&len.to_le_bytes());
    out.push(1); // flags: uncompressed UTF-16LE
    for ch in utf16 {
        out.extend_from_slice(&ch.to_le_bytes());
    }
    out
}

fn obj_record_with_ftcmo(object_id: u16) -> Vec<u8> {
    // OBJ record [MS-XLS 2.4.178] contains a series of subrecords.
    // For legacy NOTE comments we only need `ftCmo` (common object data) + `ftEnd`.
    const FT_CMO: u16 = 0x0015;
    const FT_END: u16 = 0x0000;

    // In BIFF8, `cb` is the size of the subrecord payload *excluding* the `ft/cb` header.
    const CMO_CB: u16 = 0x0012;
    const OBJECT_TYPE_NOTE: u16 = 0x0019;

    let mut out = Vec::<u8>::new();

    // ftCmo header
    out.extend_from_slice(&FT_CMO.to_le_bytes());
    out.extend_from_slice(&CMO_CB.to_le_bytes());
    // ftCmo payload (18 bytes)
    out.extend_from_slice(&OBJECT_TYPE_NOTE.to_le_bytes()); // ot
    out.extend_from_slice(&object_id.to_le_bytes()); // id
    out.extend_from_slice(&0u16.to_le_bytes()); // grbit
    out.extend_from_slice(&[0u8; 12]); // reserved

    // ftEnd subrecord
    out.extend_from_slice(&FT_END.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());

    out
}

fn push_txo_logical_record(out: &mut Vec<u8>, text: &str) {
    let cch_text: u16 = text
        .len()
        .try_into()
        .expect("comment text too long for u16 length");

    // TXO record payload (18 bytes).
    // Store `cchText` as u16 at offset 6, and `cbRuns` as u16 at offset 12.
    let mut txo = [0u8; 18];
    txo[6..8].copy_from_slice(&cch_text.to_le_bytes());
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(out, RECORD_TXO, &txo);

    // First CONTINUE record: BIFF8 continued-string form: [flags: u8][chars...]
    let mut cont_text = Vec::<u8>::new();
    cont_text.push(0); // flags: compressed 8-bit chars
    cont_text.extend_from_slice(text.as_bytes());
    push_record(out, RECORD_CONTINUE, &cont_text);

    // Second CONTINUE record: formatting runs. We use 4 bytes of dummy data.
    push_record(out, RECORD_CONTINUE, &[0u8; 4]);
}

fn push_txo_logical_record_compressed_segments(
    out: &mut Vec<u8>,
    cch_text: u16,
    segments: &[&[u8]],
) {
    // TXO record payload (18 bytes).
    // Store `cchText` as u16 at offset 6, and `cbRuns` as u16 at offset 12.
    let mut txo = [0u8; 18];
    txo[6..8].copy_from_slice(&cch_text.to_le_bytes());
    txo[12..14].copy_from_slice(&4u16.to_le_bytes()); // cbRuns
    push_record(out, RECORD_TXO, &txo);

    // One or more `CONTINUE` records containing the continued-string payload.
    // Each segment must begin with the BIFF8 string option flags byte.
    for seg in segments {
        let mut cont_text = Vec::<u8>::new();
        cont_text.push(0); // flags: compressed 8-bit chars
        cont_text.extend_from_slice(seg);
        push_record(out, RECORD_CONTINUE, &cont_text);
    }

    // Final `CONTINUE` record: formatting runs. We use 4 bytes of dummy data.
    push_record(out, RECORD_CONTINUE, &[0u8; 4]);
}

fn build_rich_styles_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
                                                           // Custom palette: indices start at 8.
                                                           // - 8 => red
                                                           // - 9 => green
    push_record(
        &mut globals,
        RECORD_PALETTE,
        &palette(&[(255, 0, 0), (0, 255, 0)]),
    );

    // Font table: default + styled.
    push_record(&mut globals, RECORD_FONT, &font("Arial"));
    push_record(
        &mut globals,
        RECORD_FONT,
        &font_with_options(FontOptions {
            name: "Courier New",
            height_twips: 200, // 10pt
            weight: 700,       // bold
            italic: true,
            underline: true,
            strike: true,
            color_idx: 8, // palette red
        }),
    );

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Two rich cell XFs:
    // - index 16: solid fill
    // - index 17: mediumGray pattern fill (exercises non-solid fill patterns)
    let xf_rich = 16u16;
    let xf_rich_medium_gray = 17u16;
    push_record(&mut globals, RECORD_XF, &xf_record_rich());
    push_record(
        &mut globals,
        RECORD_XF,
        &xf_record_rich_with_fill_pattern(2),
    );

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Styles");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_rich_styles_sheet_stream(xf_rich, xf_rich_medium_gray);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_rich_styles_sheet_stream(xf_rich: u16, xf_rich_medium_gray: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 2) (A..B)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // A1: number cell with rich style.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_rich, 0.5));
    // B1: number cell with a non-solid fill pattern.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 1, xf_rich_medium_gray, 0.25),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_formula_sheet_name_truncation_workbook_stream() -> Vec<u8> {
    // This workbook contains:
    // - Sheet 0: an over-long name (invalid; will be truncated on import), with a numeric A1.
    // - Sheet 1: `Ref`, with a formula in A1 that references Sheet0!A1.
    //
    // The important part is that the formula token stream encodes a 3D reference using an
    // EXTERNSHEET table entry, so calamine decodes it back into a sheet-qualified formula.
    let long_name = "A".repeat(40);
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records (workbook sheet list).
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in [long_name.as_str(), "Ref"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0 ------------------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 1.0));

    // -- Sheet 1 ------------------------------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_ref3d_formula_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_ptgexp_missing_shrfmla_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgexp_missing_shrfmla_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("Shared", &sheet_stream, 1252)
}

fn build_shared_formula_ptgexp_missing_shrfmla_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 2) cols [0, 2) => A1:B2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide at least one value cell so the sheet is not completely empty.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 0.0));

    // Base formula in B1: `A1+1` with relative row/col flags set (so it should become `A2+1`
    // when filled down).
    let rgce_base = vec![
        0x24, // PtgRef
        0x00, 0x00, // rw = 0
        0x00, 0xC0, // col = 0 | row_rel | col_rel
        0x1E, // PtgInt
        0x01, 0x00, // 1
        0x03, // PtgAdd
    ];
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &rgce_base),
    );

    // Follower formula in B2: `PtgExp` pointing at base cell B1 (row=0, col=1).
    let rgce_ptgexp = vec![0x01, 0x00, 0x00, 0x01, 0x00];
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &rgce_ptgexp),
    );

    // Intentionally omit SHRFMLA/ARRAY definition records.

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_ptgfuncvar_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgfuncvar_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedFormula", &sheet_stream, 1252)
}

fn build_shared_formula_ptgfuncvar_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [0, 2) => A1:B2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide numeric inputs in A1/A2 so the references are within the sheet's used range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // Base cell B1 formula: `SUM(A1,1)` encoded as a full `FORMULA.rgce` token stream.
    //
    // We keep the formula's cell reference relative (no `$`) so the expected text matches
    // `SUM(A1,1)` rather than `$A$1`.
    let full_rgce: Vec<u8> = vec![
        0x24, // PtgRef
        0x00, 0x00, // row = 0
        0x00, 0xC0, // col = 0 (A) + row_rel + col_rel
        0x1E, // PtgInt
        0x01, 0x00, // 1
        0x22, // PtgFuncVar
        0x02, // argc=2
        0x04, 0x00, // iftab=4 (SUM)
    ];
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &full_rgce),
    );

    // Shared formula rgce stored in SHRFMLA:
    //   PtgRefN(row_off=0,col_off=-1) + PtgInt(1) + PtgFuncVar(argc=2,iftab=SUM)
    let shared_rgce: Vec<u8> = {
        let mut v = Vec::new();

        // PtgRefN: [ptg=0x2C][rw:u16][col:u16]
        // - rw stores row offset when ROW_RELATIVE bit is set.
        // - col stores 14-bit signed col offset when COL_RELATIVE bit is set.
        // Use row_off=0, col_off=-1 relative to the formula cell.
        v.push(0x2C);
        v.extend_from_slice(&0u16.to_le_bytes()); // row_off = 0
        v.extend_from_slice(&0xFFFFu16.to_le_bytes()); // col_off = -1 (14-bit), row+col relative

        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());

        v.push(0x22); // PtgFuncVar
        v.push(2); // argc
        v.extend_from_slice(&0x0004u16.to_le_bytes()); // SUM
        v
    };

    // SHRFMLA record defining shared rgce for the range B1:B2.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &shared_rgce),
    );

    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));

    // Follower cell B2: PtgExp(B1).
    let ptgexp = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_merged_formatted_blank_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One cell XF: built-in percent format (numFmtId=10 => "0.00%").
    let xf_percent = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 10, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "MergedFmt");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_merged_formatted_blank_sheet_stream(xf_percent);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_merged_conflicting_blank_formats_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Two cell XFs with different built-in number formats.
    let xf_percent = 16u16;
    let xf_duration = 17u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 10, false)); // 0.00% (built-in)
    push_record(&mut globals, RECORD_XF, &xf_record(0, 46, false)); // [h]:mm:ss (built-in)

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "MergedFmtConflict");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_merged_conflicting_blank_formats_sheet_stream(xf_percent, xf_duration);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_out_of_range_xf_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One cell XF (valid), so the importer enables XF scanning with an "interesting" mask.
    let xf_percent = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 10, false)); // 0.00% (built-in)

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "OutOfRangeXF");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    let sheet_offset = globals.len();
    let sheet = build_out_of_range_xf_sheet_stream(xf_percent);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_merged_non_anchor_conflicting_blank_formats_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Two cell XFs with different built-in number formats.
    let xf_percent = 16u16;
    let xf_duration = 17u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 10, false)); // 0.00% (built-in)
    push_record(&mut globals, RECORD_XF, &xf_record(0, 46, false)); // [h]:mm:ss (built-in)

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "MergedFmtNonAnchorConflict");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet =
        build_merged_non_anchor_conflicting_blank_formats_sheet_stream(xf_percent, xf_duration);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_merged_non_anchor_formula_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF (required by some readers).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "MergedFormula");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_merged_non_anchor_formula_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_out_of_range_xf_no_formats_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // No cell XFs and no custom FORMAT records -> all XFs are General (mask is all-false).

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "OutOfRangeXFNoFormats");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    let sheet_offset = globals.len();
    let sheet = build_out_of_range_xf_no_formats_sheet_stream();

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_row_col_style_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One default (General) cell XF plus two non-default XFs referenced only by ROW/COLINFO.
    let xf_cell_general = 16u16;
    let xf_row_percent = 17u16;
    let xf_col_duration = 18u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false)); // General (default)
    push_record(&mut globals, RECORD_XF, &xf_record(0, 10, false)); // 0.00% (built-in)
    push_record(&mut globals, RECORD_XF, &xf_record(0, 46, false)); // [h]:mm:ss (built-in)

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "RowColStyles");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_row_col_style_sheet_stream(xf_cell_general, xf_row_percent, xf_col_duration);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_continued_format_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // FORMAT record split across a CONTINUE record.
    let format_string = "yyyy-mm-dd hh:mm:ss";
    const FMT_CONT: u16 = 0x00A4;
    let split_at = 10usize;
    let mut fmt_part1 = Vec::new();
    fmt_part1.extend_from_slice(&FMT_CONT.to_le_bytes());
    fmt_part1.extend_from_slice(&(format_string.len() as u16).to_le_bytes()); // cch
    fmt_part1.push(0); // flags (compressed)
    fmt_part1.extend_from_slice(&format_string.as_bytes()[..split_at]);
    push_record(&mut globals, RECORD_FORMAT, &fmt_part1);

    let mut fmt_cont = Vec::new();
    fmt_cont.push(0); // continued segment is also compressed
    fmt_cont.extend_from_slice(&format_string.as_bytes()[split_at..]);
    push_record(&mut globals, RECORD_CONTINUE, &fmt_cont); // CONTINUE

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    let xf_cont = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, FMT_CONT, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "ContinuedFmt");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    let sheet_offset = globals.len();
    let sheet = build_continued_format_sheet_stream(xf_cont);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_unknown_builtin_numfmtid_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Unknown/reserved built-in numFmtId (not in OOXML 0-49 table).
    const UNKNOWN_NUM_FMT_ID: u16 = 60;
    let xf_unknown = 16u16;
    push_record(
        &mut globals,
        RECORD_XF,
        &xf_record(0, UNKNOWN_NUM_FMT_ID, false),
    );

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "UnknownBuiltinFmt");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    let sheet_offset = globals.len();
    let sheet = build_unknown_builtin_numfmtid_sheet_stream(xf_unknown);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_defined_names_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Two worksheets.
    let boundsheet1_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    let boundsheet2_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Sheet2");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet2_offset_pos = boundsheet2_start + 4;

    // Minimal EXTERNSHEET table with two internal sheet entries.
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0), (1, 1)]),
    );

    // NAME records.
    // 1) Workbook-scoped name referencing Sheet1!$A$1 via PtgRef3d.
    let global_rgce = ptg_ref3d(0, 0, 0);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("GlobalName", 0, false, None, &global_rgce),
    );

    // 2) Workbook-scoped name: ZedName -> Sheet1!$B$1
    let zed_rgce = ptg_ref3d(0, 0, 1);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("ZedName", 0, false, None, &zed_rgce),
    );

    // 3) Sheet-scoped name: LocalName (scope Sheet1) -> Sheet1!$A$1
    let local_rgce = ptg_ref3d(0, 0, 0);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record(
            "LocalName",
            1,
            false,
            Some("Local description"),
            &local_rgce,
        ),
    );

    // 4) Hidden workbook-scoped name: HiddenName -> Sheet1!$A$1:$B$2
    let hidden_rgce = ptg_area3d(0, 0, 1, 0, 1);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("HiddenName", 0, true, None, &hidden_rgce),
    );

    // 5) Union to exercise PtgUnion decoding.
    let union_rgce = [ptg_ref3d(0, 0, 0), ptg_ref3d(0, 0, 1), vec![0x10u8]].concat();
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("UnionName", 0, false, None, &union_rgce),
    );

    // 6) Function call to exercise PtgFuncVar decoding: MyName -> SUM(Sheet1!$A$1:$A$3)
    let mut sum_rgce = ptg_area3d(0, 0, 2, 0, 0);
    sum_rgce.extend_from_slice(&[0x22u8, 0x01, 0x04, 0x00]); // PtgFuncVar argc=1 iftab=4 (SUM)
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("MyName", 0, false, None, &sum_rgce),
    );

    // 7) Fixed-arity function to exercise PtgFunc decoding: AbsName -> ABS(1)
    let abs_rgce = vec![0x1E, 0x01, 0x00, 0x21, 0x18, 0x00]; // PtgInt 1; PtgFunc iftab=24 (ABS)
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("AbsName", 0, false, None, &abs_rgce),
    );

    // 8) Union inside a function argument must be parenthesized: UnionFunc -> SUM((Sheet1!$A$1,Sheet1!$B$1))
    let mut union_func_rgce = [ptg_ref3d(0, 0, 0), ptg_ref3d(0, 0, 1), vec![0x10u8]].concat();
    union_func_rgce.extend_from_slice(&[0x22u8, 0x01, 0x04, 0x00]); // SUM
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("UnionFunc", 0, false, None, &union_func_rgce),
    );

    // 9) Missing argument slot: MissingArgName -> IF(,1,2)
    let miss_rgce = vec![
        0x16u8, // PtgMissArg
        0x1E, 0x01, 0x00, // PtgInt 1
        0x1E, 0x02, 0x00, // PtgInt 2
        0x22, 0x03, 0x01, 0x00, // PtgFuncVar argc=3 iftab=1 (IF)
    ];
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("MissingArgName", 0, false, None, &miss_rgce),
    );

    // 10) Built-in print area (`_xlnm.Print_Area`) on Sheet1 using a union of two 3D areas (hidden).
    // Excel stores this as a built-in NAME record (fBuiltin=1, rgchName=builtin_id).
    let print_area_rgce = [
        ptg_area3d(0, 0, 1, 0, 1),
        ptg_area3d(0, 3, 4, 3, 4),
        vec![0x10u8],
    ]
    .concat();
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x06, &print_area_rgce),
    );

    // 11) Reference to another defined name via PtgName: RefToGlobal -> GlobalName
    let ref_to_global_rgce = ptg_name(1);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("RefToGlobal", 0, false, None, &ref_to_global_rgce),
    );

    // 12) Reference to a sheet-scoped name via PtgName: RefToLocal -> Sheet1!LocalName
    let ref_to_local_rgce = ptg_name(3);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("RefToLocal", 0, false, None, &ref_to_local_rgce),
    );

    // 13) Duplicate workbook-scoped name (should be skipped by the importer with a warning).
    let dup_rgce = ptg_ref3d(0, 0, 2); // Sheet1!$C$1
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("GlobalName", 0, false, None, &dup_rgce),
    );

    // 14) Invalid name (looks like a cell reference) should be skipped with a warning.
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("A1", 0, false, None, &global_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet substreams -------------------------------------------------------
    let sheet1_offset = globals.len();
    let sheet1 = build_empty_sheet_stream(xf_general);
    let sheet2_offset = sheet1_offset + sheet1.len();
    let sheet2 = build_empty_sheet_stream(xf_general);

    // Patch BoundSheet offsets.
    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet1);
    globals.extend_from_slice(&sheet2);

    globals
}

fn build_continued_name_record_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "DefinedNames");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // Minimal EXTERNSHEET table with a single internal sheet entry (itab 0).
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // A workbook-scoped defined name referencing DefinedNames!$A$1.
    let name = "MyContinuedName";
    let description = "This is a long description used to test continued NAME records.";
    let rgce = ptg_ref3d(0, 0, 0);

    // Split rgce after 3 bytes (mid token payload) and split description after 10 bytes of
    // character data.
    let (name_part1, cont1, cont2) =
        continued_name_record_fragments(name, 0, false, description, &rgce, 3, 10);
    push_record(&mut globals, RECORD_NAME, &name_part1);
    push_record(&mut globals, RECORD_CONTINUE, &cont1);
    push_record(&mut globals, RECORD_CONTINUE, &cont2);

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    // Include a formula cell that references the defined name so we exercise calamine's `PtgName`
    // decoding path (which depends on successfully ingesting the NAME record header).
    let sheet = build_name_reference_formula_sheet_stream(xf_general, 1);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_defined_names_quoting_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Three worksheets with names requiring quoting rules.
    let boundsheet1_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Sheet One");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    let boundsheet2_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "O'Brien");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet2_offset_pos = boundsheet2_start + 4;

    let boundsheet3_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "TRUE");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet3_offset_pos = boundsheet3_start + 4;

    // Minimal EXTERNSHEET table with:
    // - three internal sheet entries, plus
    // - one sheet span entry (Sheet One -> O'Brien).
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0), (1, 1), (2, 2), (0, 1)]),
    );

    // NAME records:
    // Workbook-scoped names that reference each sheet + a 3D span.
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("SpaceRef", 0, false, None, &ptg_ref3d(0, 0, 0)),
    ); // Sheet One!$A$1
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("QuoteRef", 0, false, None, &ptg_ref3d(1, 1, 1)),
    ); // O'Brien!$B$2
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("ReservedRef", 0, false, None, &ptg_ref3d(2, 2, 2)),
    ); // TRUE!$C$3 (must be quoted)
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("SpanRef", 0, false, None, &ptg_ref3d(3, 3, 3)),
    ); // Sheet One:O'Brien!$D$4

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet substreams -------------------------------------------------------
    let sheet1_offset = globals.len();
    let sheet1 = build_empty_sheet_stream(xf_general);
    let sheet2_offset = sheet1_offset + sheet1.len();
    let sheet2 = build_empty_sheet_stream(xf_general);
    let sheet3_offset = sheet2_offset + sheet2.len();
    let sheet3 = build_empty_sheet_stream(xf_general);

    // Patch BoundSheet offsets.
    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());
    globals[boundsheet3_offset_pos..boundsheet3_offset_pos + 4]
        .copy_from_slice(&(sheet3_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet1);
    globals.extend_from_slice(&sheet2);
    globals.extend_from_slice(&sheet3);

    globals
}

fn build_defined_names_external_workbook_refs_workbook_stream() -> Vec<u8> {
    // This workbook contains:
    // - One internal sheet (`Local`) so calamine considers the workbook valid.
    // - SUPBOOK[0]: internal workbook marker
    // - SUPBOOK[1]: external workbook `Book1.xlsx` with sheets SheetA/SheetB/SheetC
    // - EXTERNSHEET entries pointing at SUPBOOK[1]
    // - Defined names referencing those EXTERNSHEET entries via PtgRef3d
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet (internal).
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Local");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // External reference tables.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(1)); // internal workbook marker
    push_record(
        &mut globals,
        RECORD_SUPBOOK,
        &supbook_external("Book1.xlsx", &["SheetA", "SheetB", "SheetC"]),
    );

    // External names (EXTERNNAME) for PtgNameX references. These are attached to the preceding
    // SUPBOOK (the external workbook).
    push_record(
        &mut globals,
        RECORD_EXTERNNAME,
        &externname_record("ExtDefined"),
    );
    push_record(&mut globals, RECORD_EXTERNNAME, &externname_record("MyUdf"));

    // Additional external SUPBOOK without a corresponding EXTERNSHEET entry. Some writers appear to
    // store `PtgNameX.ixti` as a SUPBOOK index directly (when EXTERNSHEET is missing). Our decoder
    // treats missing EXTERNSHEET as a signal to interpret `ixti` as a SUPBOOK index.
    //
    // This SUPBOOK is only used for the PtgNameX workbook-scoped external-name test.
    push_record(
        &mut globals,
        RECORD_SUPBOOK,
        &supbook_external("Book2.xlsx", &[]),
    );
    push_record(
        &mut globals,
        RECORD_EXTERNNAME,
        &externname_record("WBName"),
    );
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record_with_supbook(&[
            // ixti=0 => [Book1.xlsx]SheetA
            (1, 0, 0),
            // ixti=1 => [Book1.xlsx]SheetA:SheetC
            (1, 0, 2),
        ]),
    );

    // Defined names referencing external sheets via PtgRef3d.
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("ExtSingle", 0, false, None, &ptg_ref3d(0, 0, 0)),
    );
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("ExtSpan", 0, false, None, &ptg_ref3d(1, 0, 0)),
    );
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("ExtNameX", 0, false, None, &ptg_namex(0, 1)),
    );

    // User-defined function call via PtgNameX + PtgFuncVar(0x00FF): should render as `MyUdf(1)`.
    let udf_rgce = [
        vec![0x1E, 0x01, 0x00],
        ptg_namex(0, 2),
        vec![0x22, 0x02, 0xFF, 0x00],
    ]
    .concat();
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("ExtUdfCall", 0, false, None, &udf_rgce),
    );

    // External workbook-scoped name reference with missing EXTERNSHEET (ixti=2 => SUPBOOK index 2).
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("ExtNameXWb", 0, false, None, &ptg_namex(2, 1)),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_empty_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_defined_name_calamine_workbook_stream() -> Vec<u8> {
    build_defined_name_calamine_workbook_stream_with_sheet_name("Sheet1")
}

fn build_defined_name_calamine_workbook_stream_with_sheet_name(sheet_name: &str) -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, sheet_name);
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // Minimal SUPBOOK entry for internal workbook references.
    let supbook = {
        let mut data = Vec::<u8>::new();
        data.extend_from_slice(&1u16.to_le_bytes()); // ctab (sheet count)
        data.extend_from_slice(&1u16.to_le_bytes()); // cch
        data.push(0); // flags (compressed)
        data.push(0); // virtPath = "\0" (internal workbook marker)
        data
    };
    push_record(&mut globals, RECORD_SUPBOOK, &supbook);

    // Minimal EXTERNSHEET table with a single internal sheet entry.
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // One workbook-scoped defined name: TestName -> <sheet_name>!$A$1:$A$1.
    let rgce = ptg_area3d(0, 0, 0, 0, 0);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record_calamine_compat("TestName", &rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_empty_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_defined_name_sheet_name_sanitization_calamine_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet with an invalid name (will sanitize to `Bad_Name`).
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Bad:Name");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // Minimal SUPBOOK entry for internal workbook references (calamine-compatible encoding).
    let supbook = {
        let mut data = Vec::<u8>::new();
        data.extend_from_slice(&1u16.to_le_bytes()); // ctab (sheet count)
        data.extend_from_slice(&1u16.to_le_bytes()); // cch
        data.push(0); // flags (compressed)
        data.push(0); // virtPath = "\0" (internal workbook marker)
        data
    };
    push_record(&mut globals, RECORD_SUPBOOK, &supbook);

    // Minimal EXTERNSHEET table with a single internal sheet entry.
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // Workbook-scoped defined name: BadRef -> Bad:Name!$A$1:$A$1
    let rgce = ptg_area3d(0, 0, 0, 0, 0);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record_calamine_compat("BadRef", &rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_empty_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_print_settings_calamine_workbook_stream() -> Vec<u8> {
    // Similar to `build_defined_names_builtins_workbook_stream`, but encodes print built-ins as
    // regular defined name strings so calamine surfaces them via `Reader::defined_names()`.
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Two worksheets.
    let boundsheet1_start = globals.len();
    let mut boundsheet1 = Vec::<u8>::new();
    boundsheet1.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet1.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet1, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet1);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    let boundsheet2_start = globals.len();
    let mut boundsheet2 = Vec::<u8>::new();
    boundsheet2.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet2.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet2, "Sheet2");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet2);
    let boundsheet2_offset_pos = boundsheet2_start + 4;

    // External reference tables so calamine can decode 3D references in the NAME formula stream.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0), (1, 1)]),
    );

    // `_xlnm.Print_Area` for Sheet1: Sheet1!$A$1:$A$2.
    //
    // Note: Although Excel can store multiple print areas using the union operator (`,` /
    // `PtgUnion`), calamine's `.xls` defined-name decoder does not currently preserve all union
    // operands reliably. Keep this fixture to a single rectangular range so it remains stable.
    let print_area_rgce = ptg_area3d(0, 0, 1, 0, 0);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record_calamine_compat(XLNM_PRINT_AREA, &print_area_rgce),
    );

    // `_xlnm.Print_Titles` for Sheet2: Sheet2!$1:$1 (repeat first row).
    let print_titles_rgce = ptg_area3d(1, 0, 0, 0, 0x00FF);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record_calamine_compat(XLNM_PRINT_TITLES, &print_titles_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet substreams -------------------------------------------------------
    let sheet1_offset = globals.len();
    let sheet1 = build_empty_sheet_stream(xf_general);
    let sheet2_offset = sheet1_offset + sheet1.len();
    let sheet2 = build_empty_sheet_stream(xf_general);

    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet1);
    globals.extend_from_slice(&sheet2);

    globals
}

fn build_fit_to_page_without_setup_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_fit_to_page_without_setup_sheet_stream(xf_general);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);

    globals
}

fn build_sanitized_sheet_name_defined_name_collision_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Sheet 0: invalid name, will sanitize from `Bad/Name` -> `Bad_Name`.
    let boundsheet0_start = globals.len();
    let mut boundsheet0 = Vec::<u8>::new();
    boundsheet0.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet0.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet0, "Bad/Name");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet0);
    let boundsheet0_offset_pos = boundsheet0_start + 4;

    // Sheet 1: already has the sanitized base name, will dedupe to `Bad_Name (2)`.
    let boundsheet1_start = globals.len();
    let mut boundsheet1 = Vec::<u8>::new();
    boundsheet1.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet1.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet1, "Bad_Name");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet1);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    // Minimal EXTERNSHEET table with one internal sheet entry (sheet 0).
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // One workbook-scoped name that refers to the invalid sheet name (sheet 0).
    // MyRange -> 'Bad/Name'!$A$1
    let rgce = ptg_ref3d(0, 0, 0);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("MyRange", 0, false, None, &rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet substreams ------------------------------------------------------
    let sheet0_offset = globals.len();
    let sheet0 = build_empty_sheet_stream(xf_general);
    globals[boundsheet0_offset_pos..boundsheet0_offset_pos + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet0);

    let sheet1_offset = globals.len();
    let sheet1 = build_empty_sheet_stream(xf_general);
    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet1);

    globals
}

fn build_sanitized_sheet_name_defined_name_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet with an invalid name (contains `/`), which the importer will sanitize.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Bad/Name");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // Minimal EXTERNSHEET table with one internal sheet entry.
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // One workbook-scoped name that refers to the (invalid) sheet name.
    // MyRange -> 'Bad/Name'!$A$1
    let rgce = ptg_ref3d(0, 0, 0);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("MyRange", 0, false, None, &rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet substream --------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_empty_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_print_settings_unicode_sheet_name_workbook_stream() -> Vec<u8> {
    let sheet_name = "Ünicode Name";

    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF (required by some readers).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, sheet_name);
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // Minimal EXTERNSHEET table with one internal sheet entry so we can encode 3D references.
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // Print_Area on the sheet: 'Ünicode Name'!$A$1:$A$2,'Ünicode Name'!$C$1:$C$2 (hidden).
    let print_area_rgce = [
        ptg_area3d(0, 0, 1, 0, 0),
        ptg_area3d(0, 0, 1, 2, 2),
        vec![0x10], // PtgUnion
    ]
    .concat();
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x06, &print_area_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_empty_sheet_stream(xf_general);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);

    globals
}

fn build_manual_page_breaks_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF (required by some readers).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_manual_page_breaks_sheet_stream(xf_general);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);

    globals
}

fn build_manual_page_breaks_sheet_stream(xf_general: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 0.0),
    );

    // HORIZONTALPAGEBREAKS: two unique breaks + one duplicate.
    // `row` is 0-based index of the first row below the break:
    // - row=2 => break after row 1 (0-based).
    // - row=5 => break after row 4 (0-based).
    let mut horizontal = Vec::<u8>::new();
    let horizontal_rows = [2u16, 2u16, 5u16];
    horizontal.extend_from_slice(&(horizontal_rows.len() as u16).to_le_bytes());
    for row in horizontal_rows {
        horizontal.extend_from_slice(&row.to_le_bytes());
        horizontal.extend_from_slice(&0u16.to_le_bytes()); // colStart
        horizontal.extend_from_slice(&255u16.to_le_bytes()); // colEnd
    }
    push_record(&mut sheet, RECORD_HORIZONTALPAGEBREAKS, &horizontal);

    // VERTICALPAGEBREAKS: two unique breaks + one duplicate.
    // `col` is 0-based index of the first column to the right of the break:
    // - col=3 => break after col 2 (0-based).
    // - col=10 => break after col 9 (0-based).
    let mut vertical = Vec::<u8>::new();
    let vertical_cols = [3u16, 3u16, 10u16];
    vertical.extend_from_slice(&(vertical_cols.len() as u16).to_le_bytes());
    for col in vertical_cols {
        vertical.extend_from_slice(&col.to_le_bytes());
        vertical.extend_from_slice(&0u16.to_le_bytes()); // rowStart
        vertical.extend_from_slice(&0u16.to_le_bytes()); // rowEnd
    }
    push_record(&mut sheet, RECORD_VERTICALPAGEBREAKS, &vertical);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

#[derive(Clone, Copy)]
struct PageSetupFixtureSheet {
    paper_size: u16,
    landscape: bool,
    scale_percent: u16,
    header_margin: f64,
    footer_margin: f64,
    left_margin: f64,
    right_margin: f64,
    top_margin: f64,
    bottom_margin: f64,
    row_break_after: u16,
    col_break_after: u16,
    cell_value: f64,
}

fn build_page_setup_sheet_stream(xf_cell: u16, cfg: PageSetupFixtureSheet) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: ensure the configured break positions are in range.
    let last_row_plus1 = u32::from(cfg.row_break_after).saturating_add(2).max(1);
    let last_col_plus1 = cfg.col_break_after.saturating_add(2).max(1);
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&last_row_plus1.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&last_col_plus1.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Margins.
    push_record(&mut sheet, RECORD_LEFTMARGIN, &cfg.left_margin.to_le_bytes());
    push_record(&mut sheet, RECORD_RIGHTMARGIN, &cfg.right_margin.to_le_bytes());
    push_record(&mut sheet, RECORD_TOPMARGIN, &cfg.top_margin.to_le_bytes());
    push_record(
        &mut sheet,
        RECORD_BOTTOMMARGIN,
        &cfg.bottom_margin.to_le_bytes(),
    );

    // Page setup.
    push_record(
        &mut sheet,
        RECORD_SETUP,
        &setup_record(
            cfg.paper_size,
            cfg.scale_percent,
            0,
            0,
            cfg.landscape,
            cfg.header_margin,
            cfg.footer_margin,
        ),
    );

    // Manual page breaks.
    push_record(
        &mut sheet,
        RECORD_HPAGEBREAKS,
        // BIFF8 stores the 0-based index of the first row *after* the break.
        &hpagebreaks_record(&[cfg.row_break_after.saturating_add(1)]),
    );
    push_record(
        &mut sheet,
        RECORD_VPAGEBREAKS,
        // BIFF8 stores the 0-based index of the first column *after* the break.
        &vpagebreaks_record(&[cfg.col_break_after.saturating_add(1)]),
    );

    // A1: a single cell so calamine returns a non-empty range.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_cell, cfg.cell_value),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_page_setup_multisheet_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF (required by some readers).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Two worksheets.
    let boundsheet1_start = globals.len();
    let mut boundsheet1 = Vec::<u8>::new();
    boundsheet1.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet1.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet1, "First");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet1);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    let boundsheet2_start = globals.len();
    let mut boundsheet2 = Vec::<u8>::new();
    boundsheet2.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet2.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet2, "Second");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet2);
    let boundsheet2_offset_pos = boundsheet2_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet substreams -------------------------------------------------------
    let sheet1_offset = globals.len();
    let sheet1 = build_page_setup_sheet_stream(
        xf_general,
        PageSetupFixtureSheet {
            paper_size: 9,
            landscape: true,
            scale_percent: 80,
            header_margin: 0.125,
            footer_margin: 0.875,
            left_margin: 0.5,
            right_margin: 1.0,
            top_margin: 1.5,
            bottom_margin: 2.0,
            row_break_after: 4,
            col_break_after: 1,
            cell_value: 1.0,
        },
    );
    let sheet2_offset = sheet1_offset + sheet1.len();
    let sheet2 = build_page_setup_sheet_stream(
        xf_general,
        PageSetupFixtureSheet {
            paper_size: 1,
            landscape: false,
            scale_percent: 120,
            header_margin: 0.25,
            footer_margin: 0.75,
            left_margin: 0.375,
            right_margin: 0.625,
            top_margin: 1.125,
            bottom_margin: 1.875,
            row_break_after: 9,
            col_break_after: 3,
            cell_value: 2.0,
        },
    );

    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet1);
    globals.extend_from_slice(&sheet2);

    globals
}

fn build_empty_sheet_stream(xf_general: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 0.0),
    );
    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_fit_to_page_without_setup_sheet_stream(xf_general: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // WSBOOL with fFitToPage bit set (bit 8) but no SETUP record.
    //
    // We keep the other bits consistent with the `build_outline_sheet_stream` fixture so the
    // value resembles Excel output.
    let wsbool: u16 = 0x0C01 | 0x0100;
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool.to_le_bytes());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 0.0),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_name_reference_formula_sheet_stream(xf_cell: u16, name_index: u32) -> Vec<u8> {
    // Single-sheet stream containing one formula cell (A1) whose formula is a `PtgName` reference
    // to a workbook defined name.
    //
    // `name_index` is one-based (PtgName.iname).
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) (A1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    let mut rgce = Vec::<u8>::new();
    rgce.push(0x23); // PtgName
    rgce.extend_from_slice(&name_index.to_le_bytes());

    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 0, xf_cell, 0.0, &rgce),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn externsheet_record(entries: &[(u16, u16)]) -> Vec<u8> {
    // Convenience wrapper for internal-workbook XTI entries (iSupBook=0).
    let entries: Vec<(u16, u16, u16)> = entries
        .iter()
        .copied()
        .map(|(itab_first, itab_last)| (0u16, itab_first, itab_last))
        .collect();
    externsheet_record_with_supbook(&entries)
}

fn externsheet_record_with_supbook(entries: &[(u16, u16, u16)]) -> Vec<u8> {
    // EXTERNSHEET payload: [cXTI: u16][rgXTI: cXTI * 6 bytes]
    // Each XTI: [iSupBook: u16][itabFirst: u16][itabLast: u16]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&(entries.len() as u16).to_le_bytes());
    for &(i_sup_book, itab_first, itab_last) in entries {
        out.extend_from_slice(&i_sup_book.to_le_bytes());
        out.extend_from_slice(&itab_first.to_le_bytes());
        out.extend_from_slice(&itab_last.to_le_bytes());
    }
    out
}

fn name_record(
    name: &str,
    itab: u16,
    hidden: bool,
    description: Option<&str>,
    rgce: &[u8],
) -> Vec<u8> {
    // NAME record payload (BIFF8) header:
    // [grbit: u16][chKey: u8][cch: u8][cce: u16][ixals: u16][itab: u16]
    // [cchCustMenu: u8][cchDescription: u8][cchHelpTopic: u8][cchStatusText: u8]
    let mut out = Vec::<u8>::new();

    let mut grbit: u16 = 0;
    if hidden {
        grbit |= 0x0001; // fHidden
    }
    out.extend_from_slice(&grbit.to_le_bytes());
    out.push(0); // chKey

    let cch: u8 = name
        .len()
        .try_into()
        .expect("defined name too long for u8 length");
    out.push(cch);

    let cce: u16 = rgce.len().try_into().expect("rgce too long for u16 length");
    out.extend_from_slice(&cce.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // ixals
    out.extend_from_slice(&itab.to_le_bytes());

    out.push(0); // cchCustMenu

    let desc_len: u8 = description
        .map(|s| s.len().try_into().expect("description too long"))
        .unwrap_or(0);
    out.push(desc_len); // cchDescription
    out.push(0); // cchHelpTopic
    out.push(0); // cchStatusText

    // Name string (XLUnicodeStringNoCch).
    write_unicode_string_no_cch(&mut out, name);

    // Formula token stream.
    out.extend_from_slice(rgce);

    // Optional strings.
    if let Some(desc) = description {
        write_unicode_string_no_cch(&mut out, desc);
    }

    out
}

fn name_record_calamine_compat(name: &str, rgce: &[u8]) -> Vec<u8> {
    // Like `name_record`, but encodes the name string in a way calamine's `.xls` defined-name
    // parser accepts (avoids embedded NULs in the returned name).
    //
    // This is intentionally scoped to a single fixture used to test the calamine fallback path.
    let mut out = Vec::<u8>::new();

    out.extend_from_slice(&0u16.to_le_bytes()); // grbit
    out.push(0); // chKey
                 // Calamine's `.xls` NAME parser currently assumes BIFF8 NAME strings are "uncompressed"
                 // (`fHighByte=1`) but still reads only `cch` bytes of payload. For odd-length ASCII names this
                 // truncates the final byte. Work around this by padding odd-length names with a trailing NUL
                 // byte and incrementing `cch` so calamine sees the full string (our importer strips NULs).
    let mut cch: usize = name.len();
    let pad_nul = cch % 2 == 1;
    if pad_nul {
        cch = cch.saturating_add(1);
    }
    let cch: u8 = cch.try_into().expect("defined name too long for u8 length");
    out.push(cch);
    let cce: u16 = rgce.len().try_into().expect("rgce too long for u16 length");
    out.extend_from_slice(&cce.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // ixals
    out.extend_from_slice(&0u16.to_le_bytes()); // itab (0 = workbook scoped)
    out.push(0); // cchCustMenu
    out.push(0); // cchDescription
    out.push(0); // cchHelpTopic
    out.push(0); // cchStatusText

    // XLUnicodeStringNoCch with `fHighByte=1` (uncompressed). For calamine compatibility, we
    // intentionally emit a single byte per character.
    out.push(1);
    out.extend_from_slice(name.as_bytes());
    if pad_nul {
        out.push(0);
    }

    out.extend_from_slice(rgce);
    out
}

fn continued_name_record_fragments(
    name: &str,
    itab: u16,
    hidden: bool,
    description: &str,
    rgce: &[u8],
    split_rgce_at: usize,
    split_description_at: usize,
) -> (Vec<u8>, Vec<u8>, Vec<u8>) {
    // NAME record payload (BIFF8) header:
    // [grbit: u16][chKey: u8][cch: u8][cce: u16][ixals: u16][itab: u16]
    // [cchCustMenu: u8][cchDescription: u8][cchHelpTopic: u8][cchStatusText: u8]
    let mut header = Vec::<u8>::new();

    let mut grbit: u16 = 0;
    if hidden {
        grbit |= 0x0001; // fHidden
    }
    header.extend_from_slice(&grbit.to_le_bytes());
    header.push(0); // chKey

    let cch: u8 = name
        .len()
        .try_into()
        .expect("defined name too long for u8 length");
    header.push(cch);

    let cce: u16 = rgce.len().try_into().expect("rgce too long for u16 length");
    header.extend_from_slice(&cce.to_le_bytes());
    header.extend_from_slice(&0u16.to_le_bytes()); // ixals
    header.extend_from_slice(&itab.to_le_bytes());

    header.push(0); // cchCustMenu

    let desc_len: u8 = description
        .len()
        .try_into()
        .expect("description too long for u8 length");
    header.push(desc_len); // cchDescription
    header.push(0); // cchHelpTopic
    header.push(0); // cchStatusText

    let mut part1 = Vec::<u8>::new();
    part1.extend_from_slice(&header);
    // Name string (XLUnicodeStringNoCch).
    write_unicode_string_no_cch(&mut part1, name);

    let split_rgce_at = split_rgce_at.min(rgce.len());
    part1.extend_from_slice(&rgce[..split_rgce_at]);

    let mut cont1 = Vec::<u8>::new();
    cont1.extend_from_slice(&rgce[split_rgce_at..]);

    // Description string (XLUnicodeStringNoCch) begins in the first CONTINUE record, then is split
    // across another CONTINUE record.
    let desc_bytes = description.as_bytes();
    let split_description_at = split_description_at.min(desc_bytes.len());

    // Initial string fragment includes the XLUnicodeStringNoCch flags byte.
    cont1.push(0); // flags (compressed)
    cont1.extend_from_slice(&desc_bytes[..split_description_at]);

    // Continuation fragment begins with continued-segment option flags (fHighByte).
    let mut cont2 = Vec::<u8>::new();
    cont2.push(0); // continued segment option flags (compressed)
    cont2.extend_from_slice(&desc_bytes[split_description_at..]);

    (part1, cont1, cont2)
}

fn write_unicode_string_no_cch(out: &mut Vec<u8>, s: &str) {
    // BIFF8 XLUnicodeStringNoCch: [flags: u8][chars]
    // We only emit compressed (8-bit) strings in fixtures.
    out.push(0); // flags (fHighByte=0)
    out.extend_from_slice(s.as_bytes());
}

fn ptg_exp(row: u16, col: u16) -> Vec<u8> {
    // PtgExp (0x01) payload: [rw: u16][col: u16]
    let mut out = Vec::<u8>::new();
    out.push(0x01);
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out
}

fn ptg_ref3d(ixti: u16, row: u16, col: u16) -> Vec<u8> {
    // PtgRef3d (0x3A) payload: [ixti: u16][row: u16][col: u16]
    let mut out = Vec::<u8>::new();
    out.push(0x3A);
    out.extend_from_slice(&ixti.to_le_bytes());
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out
}

fn ptg_area3d(ixti: u16, row1: u16, row2: u16, col1: u16, col2: u16) -> Vec<u8> {
    // PtgArea3d (0x3B) payload: [ixti: u16][rowFirst: u16][rowLast: u16][colFirst: u16][colLast: u16]
    let mut out = Vec::<u8>::new();
    out.push(0x3B);
    out.extend_from_slice(&ixti.to_le_bytes());
    out.extend_from_slice(&row1.to_le_bytes());
    out.extend_from_slice(&row2.to_le_bytes());
    out.extend_from_slice(&col1.to_le_bytes());
    out.extend_from_slice(&col2.to_le_bytes());
    out
}

fn ptg_name(name_id: u32) -> Vec<u8> {
    // PtgName (0x23) payload: [name_id: u32][reserved: u16]
    let mut out = Vec::<u8>::new();
    out.push(0x23);
    out.extend_from_slice(&name_id.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    out
}

fn ptg_namex(ixti: u16, iname: u16) -> Vec<u8> {
    // PtgNameX (0x39) payload: [ixti: u16][iname: u16][reserved: u16]
    let mut out = Vec::<u8>::new();
    out.push(0x39);
    out.extend_from_slice(&ixti.to_le_bytes());
    out.extend_from_slice(&iname.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    out
}
fn build_outline_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Outline");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_outline_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

/// Build a BIFF8 workbook stream with a single worksheet containing a single HLINK record.
fn build_hyperlink_workbook_stream(sheet_name: &str, hlink: Vec<u8>) -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: 16 style XFs + one cell XF for the A1 cell value.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, sheet_name);
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet_offset = globals.len();
    let sheet = build_hyperlink_sheet_stream(xf_cell, hlink);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_formula_sheet_name_sanitization_workbook_stream() -> Vec<u8> {
    // This workbook contains:
    // - Sheet 0: `Bad:Name` (invalid; will be sanitized to `Bad_Name` on import), with a numeric A1.
    // - Sheet 1: `Bad_Name` (already valid, but will be renamed to `Bad_Name (2)` due to a name
    //   collision after sheet 0 sanitization).
    // - Sheet 2: `Ref`, with a formula in A1 that references `Bad:Name!A1`.
    //
    // The important part is that the formula token stream encodes a 3D reference using an
    // EXTERNSHEET table entry, so calamine decodes it back into a sheet-qualified formula.
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records (workbook sheet list).
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["Bad:Name", "Bad_Name", "Ref"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    // - SUPBOOK: one internal workbook entry (marker name = 0x01)
    // - EXTERNSHEET: one mapping for `Bad:Name` (sheet index 0)
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(3));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0 ------------------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 1.0));

    // -- Sheet 1 ------------------------------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 2.0));

    // -- Sheet 2 ------------------------------------------------------------------
    let sheet2_offset = globals.len();
    globals[boundsheet_offset_positions[2]..boundsheet_offset_positions[2] + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_ref3d_formula_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_sheet_name_sanitization_workbook_stream() -> Vec<u8> {
    // This workbook contains:
    // - Sheet 0: `Bad:Name` (invalid; will be sanitized to `Bad_Name` on import).
    // - Sheet 1: `Ref`, with a **shared formula** in A1:A2 whose rgce references `Bad:Name!A1`
    //   via `PtgRef3d`. The second cell (A2) uses `PtgExp` and must be resolved through the
    //   `SHRFMLA` record.
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records.
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["Bad:Name", "Ref"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    // - SUPBOOK: internal workbook marker
    // - EXTERNSHEET: one mapping for `Bad:Name` (sheet index 0)
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0 ------------------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 1.0));

    // -- Sheet 1 ------------------------------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_ref3d_shrfmla_sheet_stream(xf_cell));

    globals
}

fn build_calc_settings_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252

    // Workbook calculation settings (non-default values).
    // - CALCMODE: 0 = manual.
    push_record(&mut globals, RECORD_CALCMODE, &0u16.to_le_bytes());
    // - ITERATION: 1 = iterative calc enabled.
    push_record(&mut globals, RECORD_ITERATION, &1u16.to_le_bytes());
    // - CALCCOUNT: max iterations.
    push_record(&mut globals, RECORD_CALCCOUNT, &7u16.to_le_bytes());
    // - DELTA: max change.
    push_record(&mut globals, RECORD_DELTA, &0.01f64.to_le_bytes());
    // - PRECISION: 0 = precision as displayed (not full precision).
    push_record(&mut globals, RECORD_PRECISION, &0u16.to_le_bytes());
    // - SAVERECALC: 0 = don't recalc before saving.
    push_record(&mut globals, RECORD_SAVERECALC, &0u16.to_le_bytes());

    // Remaining required/standard globals.
    push_record(&mut globals, RECORD_DATEMODE, &0u16.to_le_bytes()); // DATEMODE: 1900 system
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "CalcSettings");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    let sheet_offset = globals.len();
    let sheet = build_calc_settings_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_protection_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());

    // Workbook protection (structure + windows) with a legacy password hash.
    push_record(&mut globals, RECORD_PROTECT, &1u16.to_le_bytes()); // lock structure
    push_record(&mut globals, RECORD_WINDOWPROTECT, &1u16.to_le_bytes()); // lock windows
    push_record(&mut globals, RECORD_PASSWORD, &0x83AFu16.to_le_bytes()); // "password" hash

    // Remaining required/standard globals.
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Protected");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet_offset = globals.len();
    let sheet = build_protection_sheet_stream();

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_protection_truncated_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());

    // Workbook protection records with intentionally short payloads first, followed by valid
    // records. This should yield warnings but still import the final values.
    push_record(&mut globals, RECORD_PROTECT, &[1]); // truncated (expected u16)
    push_record(&mut globals, RECORD_PROTECT, &1u16.to_le_bytes()); // lock structure
    push_record(&mut globals, RECORD_WINDOWPROTECT, &[1]); // truncated (expected u16)
    push_record(&mut globals, RECORD_WINDOWPROTECT, &1u16.to_le_bytes()); // lock windows
    push_record(&mut globals, RECORD_PASSWORD, &[0xAF]); // truncated (expected u16)
    push_record(&mut globals, RECORD_PASSWORD, &0x83AFu16.to_le_bytes()); // "password" hash

    // Remaining required/standard globals.
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "ProtectedTruncated");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet_offset = globals.len();
    let sheet = build_protection_truncated_sheet_stream();

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_tab_color_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // Keep a minimal style XF table so readers tolerate the file even if it contains no cells.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "TabColor");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // SHEETEXT: store a tab color as an XColor RGB payload.
    // The importer converts this to an OOXML-style ARGB string.
    push_record(
        &mut globals,
        RECORD_SHEETEXT,
        &sheetext_record_rgb(0x11, 0x22, 0x33),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_tab_color_sheet_stream();

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_tab_color_indexed_palette_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // Emit a PALETTE record that overrides the first palette entry (index 8) to 0x112233.
    push_record(
        &mut globals,
        RECORD_PALETTE,
        &palette_record_with_override(8, 0x11, 0x22, 0x33),
    );

    // Keep a minimal style XF table so readers tolerate the file even if it contains no cells.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "TabColorIndexed");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // SHEETEXT: store a tab color as an indexed XColor referring to palette index 8.
    push_record(&mut globals, RECORD_SHEETEXT, &sheetext_record_indexed(8));

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_tab_color_sheet_stream();

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_hyperlink_sheet_stream(xf_cell: u16, hlink: Vec<u8>) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: NUMBER record so calamine reports at least one used cell.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_HLINK, &hlink);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_simple_number_sheet_stream(xf_cell: u16, v: f64) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: NUMBER record.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, v));

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_simple_ref3d_formula_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: FORMULA record referencing the first sheet's A1 (ixti=0, row=0, col=0).
    let rgce = ptg_ref3d(0, 0, 0);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 0, xf_cell, 0.0, &rgce),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_ref3d_shrfmla_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [0, 1) (A1:A2)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Shared formula anchor: A1 formula token stream is PtgExp(A1), followed by SHRFMLA
    // containing the shared rgce.
    let ptgexp = ptg_exp(0, 0);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 0, xf_cell, 0.0, &ptgexp),
    );

    // Shared rgce: reference Sheet0!A1 via ixti=0, and encode A1 as relative (no '$') so the
    // decoded text is `Bad_Name!A1`.
    let col_rel_flags: u16 = 0xC000; // rowRel + colRel bits set; col index = 0 (A)
    let shared_rgce = ptg_ref3d(0, 0, col_rel_flags);
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 0, 0, &shared_rgce),
    );

    // Second cell in the shared range: A2 formula record containing PtgExp(A1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 0, xf_cell, 0.0, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_spill_operator_formula_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 2) (A..B)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: NUMBER record so the shared-formula reference points at an existing cell.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // B1: FORMULA record containing a shared-formula token (`PtgExp`) without an accompanying
    // `SHRFMLA` record. Calamine may drop such formulas (returning an empty formula range), while
    // our BIFF rgce decoder renders a parseable placeholder (`#UNKNOWN!`).
    //
    // PtgExp payload: [rwFirst: u16][colFirst: u16] (shared-formula reference).
    let rgce = [0x01u8, 0x00, 0x00, 0x00, 0x00].to_vec();
    let cached_numeric: [u8; 8] = 0.0f64.to_le_bytes();
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_raw_value(0, 1, xf_cell, cached_numeric, &rgce),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_calc_settings_sheet_stream(xf: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) (A1 only).
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // A1: NUMBER record (value doesn't matter; ensures calamine sees a used cell).
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf, 42.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_protection_sheet_stream() -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Worksheet protection with a legacy password hash.
    push_record(&mut sheet, RECORD_PROTECT, &1u16.to_le_bytes());
    push_record(&mut sheet, RECORD_PASSWORD, &0xCBEBu16.to_le_bytes()); // "test" hash
                                                                        // Allow editing objects and scenarios while protection is enabled.
                                                                        // This verifies we correctly map BIFF's "is protected" flags to our model's "is allowed"
                                                                        // booleans.
    push_record(&mut sheet, RECORD_OBJPROTECT, &0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_SCENPROTECT, &0u16.to_le_bytes());

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_protection_truncated_sheet_stream() -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Worksheet protection records with intentionally short payloads first, followed by valid
    // ones. This should yield warnings but still import the final values.
    push_record(&mut sheet, RECORD_PROTECT, &[1]); // truncated
    push_record(&mut sheet, RECORD_PROTECT, &1u16.to_le_bytes());
    push_record(&mut sheet, RECORD_PASSWORD, &[0xEB]); // truncated
    push_record(&mut sheet, RECORD_PASSWORD, &0xCBEBu16.to_le_bytes()); // "test" hash
    push_record(&mut sheet, RECORD_OBJPROTECT, &[0]); // truncated
    push_record(&mut sheet, RECORD_OBJPROTECT, &0u16.to_le_bytes()); // allow objects
    push_record(&mut sheet, RECORD_SCENPROTECT, &[0]); // truncated
    push_record(&mut sheet, RECORD_SCENPROTECT, &0u16.to_le_bytes()); // allow scenarios

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_sheet_protection_allow_flags_sheet_stream(include_malformed_feat: bool) -> Vec<u8> {
    // Enable sheet protection and populate additional allow flags via FEAT/FEATHEADR records.
    //
    // The allow-flag mask here is intentionally non-trivial so tests can validate multiple fields.
    // Bit layout (best-effort, matches importer):
    //   bit0  select_locked_cells
    //   bit1  select_unlocked_cells
    //   bit2  format_cells
    //   bit3  format_columns
    //   bit4  format_rows
    //   bit5  insert_columns
    //   bit6  insert_rows
    //   bit7  insert_hyperlinks
    //   bit8  delete_columns
    //   bit9  delete_rows
    //   bit10 sort
    //   bit11 auto_filter
    //   bit12 pivot_tables
    //
    // Mask chosen:
    // - select_locked_cells = false
    // - select_unlocked_cells = true
    // - format_cells = true
    // - format_columns = true
    // - insert_columns = true
    // - insert_hyperlinks = true
    // - delete_rows = true
    // - sort = true
    // - auto_filter = true
    let allow_mask: u16 = 0x0EAE;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Basic worksheet protection records.
    push_record(&mut sheet, RECORD_PROTECT, &1u16.to_le_bytes());
    push_record(&mut sheet, RECORD_PASSWORD, &0xCBEBu16.to_le_bytes()); // "test" hash
                                                                        // Allow editing objects and scenarios while protection is enabled.
    push_record(&mut sheet, RECORD_OBJPROTECT, &0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_SCENPROTECT, &0u16.to_le_bytes());

    // Enhanced allow flags. Include both FEATHEADR (header data) and FEAT (feat data) variants;
    // Excel typically emits these as part of its shared-feature plumbing.
    push_record(
        &mut sheet,
        RECORD_FEATHEADR,
        &feat_hdr_record_sheet_protection_allow_mask(allow_mask),
    );

    if include_malformed_feat {
        push_record(
            &mut sheet,
            RECORD_FEAT,
            &feat_record_sheet_protection_allow_mask_malformed(allow_mask),
        );
    }

    push_record(
        &mut sheet,
        RECORD_FEAT,
        &feat_record_sheet_protection_allow_mask(allow_mask),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_sheet_protection_allow_flags_feat_continued_sheet_stream() -> Vec<u8> {
    // Same allow-flag mask as `build_sheet_protection_allow_flags_sheet_stream`, but store it in a
    // continued FEAT record (without FEATHEADR) so the importer must reassemble across `CONTINUE`.
    let allow_mask: u16 = 0x0EAE;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Basic worksheet protection records.
    push_record(&mut sheet, RECORD_PROTECT, &1u16.to_le_bytes());
    push_record(&mut sheet, RECORD_PASSWORD, &0xCBEBu16.to_le_bytes()); // "test" hash
    push_record(&mut sheet, RECORD_OBJPROTECT, &0u16.to_le_bytes());
    push_record(&mut sheet, RECORD_SCENPROTECT, &0u16.to_le_bytes());

    // FEAT record split across CONTINUE: write everything up to (but excluding) the allow-mask
    // bytes, then write the remaining allow-mask bytes in a continuation record.
    let feat_payload = feat_record_sheet_protection_allow_mask(allow_mask);
    let split = feat_payload.len().saturating_sub(2);
    push_record(&mut sheet, RECORD_FEAT, &feat_payload[..split]);
    push_record(&mut sheet, RECORD_CONTINUE, &feat_payload[split..]);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_tab_color_sheet_stream() -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2
    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

/// Build a BIFF8 `.xls` fixture containing a single `mailto:` hyperlink on `A1`.
pub fn build_mailto_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "Mail",
        hlink_external_url(
            0,
            0,
            0,
            0,
            "mailto:test@example.com",
            "Email",
            "Email tooltip",
        ),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single file hyperlink (`file:`) on `A1` backed by a
/// `CLSID_FILE_MONIKER` payload.
pub fn build_file_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "File",
        hlink_file_moniker(0, 0, 0, 0, r"C:\foo\bar.txt", "Local file", "File tooltip"),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single UNC file hyperlink (`file:`) on `A1` backed by
/// a `CLSID_FILE_MONIKER` payload.
pub fn build_unc_file_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "UNC",
        hlink_file_moniker(
            0,
            0,
            0,
            0,
            r"\\server\share\file.xlsx",
            "UNC file",
            "UNC tooltip",
        ),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single file hyperlink (`file:`) on `A1` whose
/// `CLSID_FILE_MONIKER` payload includes a Unicode path extension.
///
/// The ANSI segment in this fixture uses the workbook codepage (Windows-1252) but carries UTF-8
/// bytes for the non-ASCII filename. This ensures the importer prefers the Unicode tail when
/// present.
pub fn build_unicode_file_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "Unicode",
        hlink_file_moniker(
            0,
            0,
            0,
            0,
            r"C:\foo\日本.txt",
            "Unicode file",
            "Unicode tooltip",
        ),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a malformed file hyperlink whose FileMoniker declares
/// an ANSI path length larger than the available bytes.
///
/// The importer should not crash; it should emit a warning and skip the hyperlink.
pub fn build_malformed_file_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "FileBad",
        hlink_file_moniker_malformed_ansi_len(0, 0, 0, 0, 500),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a single internal hyperlink on `A1`.
pub fn build_internal_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "Internal",
        hlink_internal_location(0, 0, 0, 0, "Internal!B2", "Go to B2", "Internal tooltip"),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where the worksheet name is invalid and will be sanitized by the
/// importer, but an internal hyperlink still references the original name.
pub fn build_sanitized_sheet_name_internal_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "A:B",
        hlink_internal_location(0, 0, 0, 0, "A:B!B2", "Go to B2", "Internal tooltip"),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where a sheet name is invalid and will be sanitized by the
/// importer, but a cross-sheet formula still references the original name.
///
/// This is used to verify that the `.xls` importer rewrites formulas after sheet name
/// sanitization (similar to how internal hyperlinks are already rewritten).
pub fn build_formula_sheet_name_sanitization_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_formula_sheet_name_sanitization_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where a sheet name is invalid and will be sanitized by the
/// importer, but a **shared formula** (SHRFMLA + PtgExp) still references the original sheet via a
/// 3D token (`PtgRef3d`).
///
/// This is used to validate BIFF-decoded shared formula rendering after sheet-name sanitization.
pub fn build_shared_formula_sheet_name_sanitization_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_sheet_name_sanitization_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a merged region (`A1:B1`) and a hyperlink anchored to
/// a single cell within the merged region.
///
/// Excel treats merged cells as a single clickable area, so the importer should expand the
/// hyperlink anchor to cover the full merged region.
pub fn build_merged_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_merged_hyperlink_workbook_stream();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing an HLINK record whose anchor range is out of Excel
/// bounds (column >= EXCEL_MAX_COLS). The importer should ignore the hyperlink.
pub fn build_out_of_bounds_hyperlink_fixture_xls() -> Vec<u8> {
    let oob_col: u16 = EXCEL_MAX_COLS as u16;
    let workbook_stream = build_hyperlink_workbook_stream(
        "OOB",
        hlink_external_url(
            0,
            0,
            oob_col,
            oob_col,
            "https://example.com",
            "Example",
            "Tooltip",
        ),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing an external hyperlink with a `textMark` (location)
/// sub-address. The importer should preserve this by appending it as a `#fragment` to the URL.
pub fn build_external_hyperlink_with_location_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "ExternalLoc",
        hlink_external_url_with_location(
            0,
            0,
            0,
            0,
            "https://example.com",
            "#Section1",
            "Example",
            "Example tooltip",
        ),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing a file hyperlink (FileMoniker) with a location
/// sub-address. The importer should decode this as an external hyperlink and preserve the location
/// as a `#fragment`.
pub fn build_file_hyperlink_with_location_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "FileLink",
        hlink_file_moniker_with_location(
            0,
            0,
            0,
            0,
            r"C:\Temp\foo.txt",
            "#Sheet2!A1",
            "File",
            "File tooltip",
        ),
    );

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture where the HLINK payload is split across a `CONTINUE` record.
pub fn build_continued_hyperlink_fixture_xls() -> Vec<u8> {
    let url = format!("https://example.com/{}", "a".repeat(200));
    let hlink = hlink_external_url(0, 0, 0, 0, &url, "Example", "Example tooltip");

    // Split the HLINK record payload into two physical records: HLINK + CONTINUE.
    let split_at = hlink.len().min(64);
    let (first, rest) = hlink.split_at(split_at);

    // -- Globals -----------------------------------------------------------------
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: 16 style XFs + one cell XF for the A1 cell value.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Continued");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: NUMBER record so calamine reports at least one used cell.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // HLINK + CONTINUE fragments.
    push_record(&mut sheet, RECORD_HLINK, first);
    push_record(&mut sheet, RECORD_CONTINUE, rest);

    push_record(&mut sheet, RECORD_EOF, &[]);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream.write_all(&globals).expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

/// Build a BIFF8 `.xls` fixture containing hyperlink edge cases:
///
/// - A malformed hyperlink record (declared string length larger than available bytes) that should
///   yield an import warning and be skipped.
/// - A valid hyperlink split across `HLINK` + multiple `CONTINUE` records, with UTF-16 string data
///   split at fragment boundaries (including an odd-byte boundary within a UTF-16 code unit).
/// - Embedded NUL characters inside the display/tooltip/location strings; the importer should
///   truncate at the first NUL for best-effort compatibility.
pub fn build_hyperlink_edge_cases_fixture_xls() -> Vec<u8> {
    // Malformed hyperlink anchored at B1.
    let malformed = hlink_external_url_malformed_tooltip_len(
        0,
        0,
        1,
        1,
        "https://example.com",
        "Bad",
        "Bad tooltip",
        1000,
    );

    // Valid continued internal hyperlink anchored at A1, pointing to B2.
    let display = "Display\u{0}After";
    let location = "EdgeCases!B2\u{0}Ignored";
    let tooltip = "Tooltip\u{0}Ignored";
    let continued = hlink_internal_location(0, 0, 0, 0, location, display, tooltip);

    // Compute split points that fall inside the `location` string's UTF-16 payload:
    // - First split: 1 byte into the UTF-16 character data (between bytes of the first code unit).
    // - Second split: on a UTF-16 code unit boundary.
    let display_field_len = hyperlink_string_field_len(display);
    let location_field_start = 32usize + display_field_len;
    let location_utf16_start = location_field_start + 4;
    let split1 = location_utf16_start + 1;
    let split2 = location_utf16_start + 10;

    // -- Globals -----------------------------------------------------------------
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: 16 style XFs + one cell XF for the A1 cell value.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "EdgeCases");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [0, 2) => A1:B2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: NUMBER record so calamine reports at least one used cell.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // Malformed HLINK record at B1.
    push_record(&mut sheet, RECORD_HLINK, &malformed);

    // Continued HLINK record at A1.
    push_hlink_record_continued(&mut sheet, &continued, &[split1, split2]);

    push_record(&mut sheet, RECORD_EOF, &[]);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream.write_all(&globals).expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

fn build_merged_hyperlink_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: 16 style XFs + one cell XF for the A1 cell value.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "MergedLink");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet_offset = globals.len();
    let sheet = build_merged_hyperlink_sheet_stream(xf_cell);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_merged_hyperlink_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 2) (A..B)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // MERGEDCELLS: 1 range, A1:B1.
    let mut merged = Vec::<u8>::new();
    merged.extend_from_slice(&1u16.to_le_bytes()); // cAreas
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwLast
    merged.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A)
    merged.extend_from_slice(&1u16.to_le_bytes()); // colLast (B)
    push_record(&mut sheet, RECORD_MERGEDCELLS, &merged);

    // A1: NUMBER record so calamine reports at least one used cell.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // A1: HLINK record pointing to https://example.com. The anchor is a single cell even though
    // A1:B1 is merged.
    push_record(
        &mut sheet,
        RECORD_HLINK,
        &hlink_external_url(
            0,
            0,
            0,
            0,
            "https://example.com",
            "Example",
            "Example tooltip",
        ),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn hlink_internal_location(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    location: &str,
    display: &str,
    tooltip: &str,
) -> Vec<u8> {
    // HLINK record layout (BIFF8) [MS-XLS 2.4.110], matching the importer’s best-effort parser.
    //
    // ref8 (8) + guid (16) + streamVersion (4) + linkOpts (4) + variable data.
    const STREAM_VERSION: u32 = 2;
    const LINK_OPTS_HAS_LOCATION: u32 = 0x0000_0008;
    const LINK_OPTS_HAS_DISPLAY: u32 = 0x0000_0010;
    const LINK_OPTS_HAS_TOOLTIP: u32 = 0x0000_0020;

    let link_opts = LINK_OPTS_HAS_LOCATION | LINK_OPTS_HAS_DISPLAY | LINK_OPTS_HAS_TOOLTIP;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());

    out.extend_from_slice(&[0u8; 16]); // hyperlink GUID (unused)
    out.extend_from_slice(&STREAM_VERSION.to_le_bytes());
    out.extend_from_slice(&link_opts.to_le_bytes());

    write_hyperlink_string(&mut out, display);
    write_hyperlink_string(&mut out, location);
    write_hyperlink_string(&mut out, tooltip);

    out
}

fn hlink_external_url(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    url: &str,
    display: &str,
    tooltip: &str,
) -> Vec<u8> {
    // HLINK record layout (BIFF8) [MS-XLS 2.4.110], matching the importer’s best-effort parser.
    //
    // ref8 (8) + guid (16) + streamVersion (4) + linkOpts (4) + variable data.
    const STREAM_VERSION: u32 = 2;
    const LINK_OPTS_HAS_MONIKER: u32 = 0x0000_0001;
    const LINK_OPTS_HAS_DISPLAY: u32 = 0x0000_0010;
    const LINK_OPTS_HAS_TOOLTIP: u32 = 0x0000_0020;

    // URL moniker CLSID: 79EAC9E0-BAF9-11CE-8C82-00AA004BA90B (COM GUID little-endian fields).
    const CLSID_URL_MONIKER: [u8; 16] = [
        0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9,
        0x0B,
    ];

    let link_opts = LINK_OPTS_HAS_MONIKER | LINK_OPTS_HAS_DISPLAY | LINK_OPTS_HAS_TOOLTIP;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());

    out.extend_from_slice(&[0u8; 16]); // hyperlink GUID (unused)
    out.extend_from_slice(&STREAM_VERSION.to_le_bytes());
    out.extend_from_slice(&link_opts.to_le_bytes());

    write_hyperlink_string(&mut out, display);

    // URL moniker: CLSID + length (bytes) + UTF-16LE URL (NUL terminated).
    out.extend_from_slice(&CLSID_URL_MONIKER);
    let mut url_utf16: Vec<u16> = url.encode_utf16().collect();
    url_utf16.push(0); // NUL terminator
    let url_bytes_len: u32 = (url_utf16.len() * 2) as u32;
    out.extend_from_slice(&url_bytes_len.to_le_bytes());
    for code_unit in url_utf16 {
        out.extend_from_slice(&code_unit.to_le_bytes());
    }

    write_hyperlink_string(&mut out, tooltip);

    out
}

fn hlink_external_url_malformed_tooltip_len(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    url: &str,
    display: &str,
    tooltip: &str,
    tooltip_cch: u32,
) -> Vec<u8> {
    // Same as `hlink_external_url`, but writes an intentionally malformed tooltip HyperlinkString
    // length prefix so the importer sees a declared length larger than the available bytes.
    const STREAM_VERSION: u32 = 2;
    const LINK_OPTS_HAS_MONIKER: u32 = 0x0000_0001;
    const LINK_OPTS_HAS_DISPLAY: u32 = 0x0000_0010;
    const LINK_OPTS_HAS_TOOLTIP: u32 = 0x0000_0020;

    // URL moniker CLSID: 79EAC9E0-BAF9-11CE-8C82-00AA004BA90B (COM GUID little-endian fields).
    const CLSID_URL_MONIKER: [u8; 16] = [
        0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9,
        0x0B,
    ];

    let link_opts = LINK_OPTS_HAS_MONIKER | LINK_OPTS_HAS_DISPLAY | LINK_OPTS_HAS_TOOLTIP;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());

    out.extend_from_slice(&[0u8; 16]); // hyperlink GUID (unused)
    out.extend_from_slice(&STREAM_VERSION.to_le_bytes());
    out.extend_from_slice(&link_opts.to_le_bytes());

    write_hyperlink_string(&mut out, display);

    // URL moniker: CLSID + length (bytes) + UTF-16LE URL (NUL terminated).
    out.extend_from_slice(&CLSID_URL_MONIKER);
    let mut url_utf16: Vec<u16> = url.encode_utf16().collect();
    url_utf16.push(0); // NUL terminator
    let url_bytes_len: u32 = (url_utf16.len() * 2) as u32;
    out.extend_from_slice(&url_bytes_len.to_le_bytes());
    for code_unit in url_utf16 {
        out.extend_from_slice(&code_unit.to_le_bytes());
    }

    // Tooltip with malformed declared length.
    write_hyperlink_string_with_declared_cch(&mut out, tooltip, tooltip_cch);

    out
}

fn hlink_external_url_with_location(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    url: &str,
    location: &str,
    display: &str,
    tooltip: &str,
) -> Vec<u8> {
    const STREAM_VERSION: u32 = 2;
    const LINK_OPTS_HAS_MONIKER: u32 = 0x0000_0001;
    const LINK_OPTS_HAS_LOCATION: u32 = 0x0000_0008;
    const LINK_OPTS_HAS_DISPLAY: u32 = 0x0000_0010;
    const LINK_OPTS_HAS_TOOLTIP: u32 = 0x0000_0020;

    // URL moniker CLSID: 79EAC9E0-BAF9-11CE-8C82-00AA004BA90B (COM GUID little-endian fields).
    const CLSID_URL_MONIKER: [u8; 16] = [
        0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9,
        0x0B,
    ];

    let link_opts = LINK_OPTS_HAS_MONIKER
        | LINK_OPTS_HAS_LOCATION
        | LINK_OPTS_HAS_DISPLAY
        | LINK_OPTS_HAS_TOOLTIP;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());

    out.extend_from_slice(&[0u8; 16]); // hyperlink GUID (unused)
    out.extend_from_slice(&STREAM_VERSION.to_le_bytes());
    out.extend_from_slice(&link_opts.to_le_bytes());

    write_hyperlink_string(&mut out, display);

    // URL moniker: CLSID + length (bytes) + UTF-16LE URL (NUL terminated).
    out.extend_from_slice(&CLSID_URL_MONIKER);
    let mut url_utf16: Vec<u16> = url.encode_utf16().collect();
    url_utf16.push(0); // NUL terminator
    let url_bytes_len: u32 = (url_utf16.len() * 2) as u32;
    out.extend_from_slice(&url_bytes_len.to_le_bytes());
    for code_unit in url_utf16 {
        out.extend_from_slice(&code_unit.to_le_bytes());
    }

    write_hyperlink_string(&mut out, location);
    write_hyperlink_string(&mut out, tooltip);

    out
}

fn hlink_file_moniker(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    path: &str,
    display: &str,
    tooltip: &str,
) -> Vec<u8> {
    // HLINK record layout (BIFF8) [MS-XLS 2.4.110], matching the importer’s best-effort parser.
    //
    // ref8 (8) + guid (16) + streamVersion (4) + linkOpts (4) + variable data.
    const STREAM_VERSION: u32 = 2;
    const LINK_OPTS_HAS_MONIKER: u32 = 0x0000_0001;
    const LINK_OPTS_HAS_DISPLAY: u32 = 0x0000_0010;
    const LINK_OPTS_HAS_TOOLTIP: u32 = 0x0000_0020;

    // File moniker CLSID: 00000303-0000-0000-C000-000000000046 (COM GUID little-endian fields).
    const CLSID_FILE_MONIKER: [u8; 16] = [
        0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x46,
    ];

    let link_opts = LINK_OPTS_HAS_MONIKER | LINK_OPTS_HAS_DISPLAY | LINK_OPTS_HAS_TOOLTIP;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());

    out.extend_from_slice(&[0u8; 16]); // hyperlink GUID (unused)
    out.extend_from_slice(&STREAM_VERSION.to_le_bytes());
    out.extend_from_slice(&link_opts.to_le_bytes());

    write_hyperlink_string(&mut out, display);

    // File moniker payload:
    // [CLSID][cAnti:u32][ansiPath:cAnti bytes incl NUL][endServer:u16][version:u16][cbUnicode:u32][unicodePath bytes].
    out.extend_from_slice(&CLSID_FILE_MONIKER);
    let mut ansi_bytes = path.as_bytes().to_vec();
    ansi_bytes.push(0);
    out.extend_from_slice(&(ansi_bytes.len() as u32).to_le_bytes());
    out.extend_from_slice(&ansi_bytes);

    // endServer + version/reserved.
    out.extend_from_slice(&0u16.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());

    // Unicode path (UTF-16LE) including a terminating NUL.
    let mut u16s: Vec<u16> = path.encode_utf16().collect();
    u16s.push(0);
    let unicode_bytes_len: u32 = (u16s.len() * 2) as u32;
    out.extend_from_slice(&unicode_bytes_len.to_le_bytes());
    for code_unit in u16s {
        out.extend_from_slice(&code_unit.to_le_bytes());
    }
    write_hyperlink_string(&mut out, tooltip);

    out
}

fn hlink_file_moniker_malformed_ansi_len(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    declared_ansi_len: u32,
) -> Vec<u8> {
    // Like `hlink_file_moniker`, but declares an ANSI path length larger than the available bytes
    // in the FileMoniker payload.
    const STREAM_VERSION: u32 = 2;
    const LINK_OPTS_HAS_MONIKER: u32 = 0x0000_0001;
    const LINK_OPTS_HAS_DISPLAY: u32 = 0x0000_0010;
    const LINK_OPTS_HAS_TOOLTIP: u32 = 0x0000_0020;

    // File moniker CLSID: 00000303-0000-0000-C000-000000000046 (COM GUID little-endian fields).
    const CLSID_FILE_MONIKER: [u8; 16] = [
        0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x46,
    ];

    let link_opts = LINK_OPTS_HAS_MONIKER | LINK_OPTS_HAS_DISPLAY | LINK_OPTS_HAS_TOOLTIP;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());

    out.extend_from_slice(&[0u8; 16]); // hyperlink GUID (unused)
    out.extend_from_slice(&STREAM_VERSION.to_le_bytes());
    out.extend_from_slice(&link_opts.to_le_bytes());

    write_hyperlink_string(&mut out, "Bad file");

    out.extend_from_slice(&CLSID_FILE_MONIKER);
    out.extend_from_slice(&declared_ansi_len.to_le_bytes());
    // Provide fewer bytes than declared (just a tiny NUL-terminated prefix).
    out.extend_from_slice(b"C\0");

    out
}

fn hlink_file_moniker_with_location(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    path: &str,
    location: &str,
    display: &str,
    tooltip: &str,
) -> Vec<u8> {
    const STREAM_VERSION: u32 = 2;
    const LINK_OPTS_HAS_MONIKER: u32 = 0x0000_0001;
    const LINK_OPTS_HAS_LOCATION: u32 = 0x0000_0008;
    const LINK_OPTS_HAS_DISPLAY: u32 = 0x0000_0010;
    const LINK_OPTS_HAS_TOOLTIP: u32 = 0x0000_0020;

    // File moniker CLSID: 00000303-0000-0000-C000-000000000046 (COM GUID little-endian fields).
    const CLSID_FILE_MONIKER: [u8; 16] = [
        0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x46,
    ];

    let link_opts = LINK_OPTS_HAS_MONIKER
        | LINK_OPTS_HAS_LOCATION
        | LINK_OPTS_HAS_DISPLAY
        | LINK_OPTS_HAS_TOOLTIP;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());

    out.extend_from_slice(&[0u8; 16]); // hyperlink GUID (unused)
    out.extend_from_slice(&STREAM_VERSION.to_le_bytes());
    out.extend_from_slice(&link_opts.to_le_bytes());

    write_hyperlink_string(&mut out, display);

    // File moniker payload:
    // [CLSID][cAnti:u32][ansiPath:cAnti bytes incl NUL][endServer:u16][reserved:u16][cbUnicode:u32][unicodePath bytes].
    out.extend_from_slice(&CLSID_FILE_MONIKER);

    let mut ansi_bytes = path.as_bytes().to_vec();
    ansi_bytes.push(0);
    out.extend_from_slice(&(ansi_bytes.len() as u32).to_le_bytes());
    out.extend_from_slice(&ansi_bytes);
    out.extend_from_slice(&0u16.to_le_bytes()); // endServer (best-effort)
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved

    let mut unicode: Vec<u16> = path.encode_utf16().collect();
    unicode.push(0);
    let unicode_len: u32 = (unicode.len() * 2) as u32;
    out.extend_from_slice(&unicode_len.to_le_bytes());
    for cu in unicode {
        out.extend_from_slice(&cu.to_le_bytes());
    }

    write_hyperlink_string(&mut out, location);
    write_hyperlink_string(&mut out, tooltip);

    out
}

fn write_hyperlink_string(out: &mut Vec<u8>, s: &str) {
    // HyperlinkString: u32 cch + UTF-16LE (including trailing NUL).
    let mut u16s: Vec<u16> = s.encode_utf16().collect();
    u16s.push(0);
    out.extend_from_slice(&(u16s.len() as u32).to_le_bytes());
    for code_unit in u16s {
        out.extend_from_slice(&code_unit.to_le_bytes());
    }
}

fn write_hyperlink_string_with_declared_cch(out: &mut Vec<u8>, s: &str, declared_cch: u32) {
    // HyperlinkString: u32 cch + UTF-16LE (including trailing NUL).
    // Some corrupted files in the wild have inconsistent `cch` values; tests use this helper to
    // build such payloads.
    let mut u16s: Vec<u16> = s.encode_utf16().collect();
    u16s.push(0);
    out.extend_from_slice(&declared_cch.to_le_bytes());
    for code_unit in u16s {
        out.extend_from_slice(&code_unit.to_le_bytes());
    }
}

fn hyperlink_string_field_len(s: &str) -> usize {
    // Length in bytes written by `write_hyperlink_string`.
    let u16_len = s.encode_utf16().count() + 1; // trailing NUL
    4 + u16_len * 2
}

fn push_hlink_record_continued(out: &mut Vec<u8>, payload: &[u8], split_points: &[usize]) {
    // Emit an HLINK record split across one or more CONTINUE records. `split_points` are absolute
    // byte offsets into the logical record payload.
    if split_points.is_empty() {
        push_record(out, RECORD_HLINK, payload);
        return;
    }

    let mut start = 0usize;
    let mut first = true;
    for &split in split_points {
        let split = split.min(payload.len());
        if split <= start {
            continue;
        }
        let frag = &payload[start..split];
        if first {
            push_record(out, RECORD_HLINK, frag);
            first = false;
        } else {
            push_record(out, RECORD_CONTINUE, frag);
        }
        start = split;
    }

    let tail = &payload[start..];
    if first {
        push_record(out, RECORD_HLINK, tail);
    } else {
        push_record(out, RECORD_CONTINUE, tail);
    }
}

fn build_merged_formatted_blank_sheet_stream(xf_percent: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 2)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // MERGEDCELLS [MS-XLS 2.4.139]: 1 range, A1:B1.
    let mut merged = Vec::<u8>::new();
    merged.extend_from_slice(&1u16.to_le_bytes()); // cAreas
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwLast
    merged.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A)
    merged.extend_from_slice(&1u16.to_le_bytes()); // colLast (B)
    push_record(&mut sheet, RECORD_MERGEDCELLS, &merged);

    // B1: BLANK record with percent format.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 1, xf_percent));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_merged_non_anchor_formula_sheet_stream(xf_general: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 2)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // MERGEDCELLS [MS-XLS 2.4.139]: 1 range, A1:B1.
    let mut merged = Vec::<u8>::new();
    merged.extend_from_slice(&1u16.to_le_bytes()); // cAreas
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwLast
    merged.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A)
    merged.extend_from_slice(&1u16.to_le_bytes()); // colLast (B)
    push_record(&mut sheet, RECORD_MERGEDCELLS, &merged);

    // B1: FORMULA record (non-anchor). Use a simple parseable rgce token stream: `1+1`.
    let rgce: [u8; 7] = [0x1E, 0x01, 0x00, 0x1E, 0x01, 0x00, 0x03];
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_general, 2.0, &rgce),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_merged_conflicting_blank_formats_sheet_stream(
    xf_percent: u16,
    xf_duration: u16,
) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 2)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // MERGEDCELLS: 1 range, A1:B1.
    let mut merged = Vec::<u8>::new();
    merged.extend_from_slice(&1u16.to_le_bytes()); // cAreas
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwLast
    merged.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A)
    merged.extend_from_slice(&1u16.to_le_bytes()); // colLast (B)
    push_record(&mut sheet, RECORD_MERGEDCELLS, &merged);

    // A1: BLANK record (anchor) with percent format.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, xf_percent));
    // B1: BLANK record (non-anchor) with duration format.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 1, xf_duration));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_out_of_range_xf_sheet_stream(xf_percent: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 2) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // A1: valid percent style.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_percent, 0.5),
    );

    // A2: BLANK with an invalid/out-of-range XF index.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(1, 0, 5000));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_out_of_range_xf_no_formats_sheet_stream() -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // A1: BLANK with an invalid/out-of-range XF index.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 0, 5000));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_row_col_style_sheet_stream(xf_cell: u16, xf_row: u16, xf_col: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 2) cols [0, 3) (A..C, rows 1..2)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Row 2 (1-based): apply a row-level default format via ROW.ixfe.
    let mut row_payload = row_record(1, false, 0, false);
    row_payload[14..16].copy_from_slice(&xf_row.to_le_bytes());
    push_record(&mut sheet, RECORD_ROW, &row_payload);

    // Column C: apply a column-level default format via COLINFO.ixfe, without any other overrides.
    let mut col_payload = colinfo_record(2, 2, false, 0, false);
    col_payload[4..6].copy_from_slice(&0u16.to_le_bytes()); // cx: default width (no override)
    col_payload[6..8].copy_from_slice(&xf_col.to_le_bytes());
    push_record(&mut sheet, RECORD_COLINFO, &col_payload);

    // Provide at least one cell so calamine returns a non-empty range. Do not reference the
    // row/col XF indices here.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_continued_format_sheet_stream(xf: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // A1: number cell with the continued custom format.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf, 45123.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_unknown_builtin_numfmtid_sheet_stream(xf: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // A1: number cell with unknown built-in numFmtId.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf, 1234.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_outline_sheet_stream(xf: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 4) cols [0, 4) (A..D, rows 1..4)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&4u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&4u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // WSBOOL: keep Excel's default worksheet boolean options (matches real fixtures).
    // This ensures the importer correctly interprets `summaryBelow/summaryRight` from BIFF.
    push_record(&mut sheet, RECORD_WSBOOL, &0x0C01u16.to_le_bytes());

    // Outline rows:
    // - Rows 2-3 (1-based) are detail rows: outline level 1 and hidden (collapsed).
    // - Row 4 (1-based) is the collapsed summary row (level 0, collapsed).
    push_record(&mut sheet, RECORD_ROW, &row_record(1, true, 1, false));
    push_record(&mut sheet, RECORD_ROW, &row_record(2, true, 1, false));
    push_record(&mut sheet, RECORD_ROW, &row_record(3, false, 0, true));

    // Outline columns:
    // - Columns B-C (1-based) are detail columns: outline level 1 and hidden (collapsed).
    // - Column D (1-based) is the collapsed summary column.
    push_record(
        &mut sheet,
        RECORD_COLINFO,
        &colinfo_record(1, 2, true, 1, false),
    );
    push_record(
        &mut sheet,
        RECORD_COLINFO,
        &colinfo_record(3, 3, false, 0, true),
    );

    // A1: number cell with a General XF.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_merged_non_anchor_conflicting_blank_formats_sheet_stream(
    xf_percent: u16,
    xf_duration: u16,
) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 3) (A..C)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // MERGEDCELLS: 1 range, A1:C1.
    let mut merged = Vec::<u8>::new();
    merged.extend_from_slice(&1u16.to_le_bytes()); // cAreas
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
    merged.extend_from_slice(&0u16.to_le_bytes()); // rwLast
    merged.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A)
    merged.extend_from_slice(&2u16.to_le_bytes()); // colLast (C)
    push_record(&mut sheet, RECORD_MERGEDCELLS, &merged);

    // No A1 record. Two conflicting non-anchor BLANK records:
    // B1: percent format, C1: duration format.
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 1, xf_percent));
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(0, 2, xf_duration));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_sheet_stream(
    xf_currency: u16,
    xf_percent: u16,
    xf_date: u16,
    xf_time: u16,
    xf_duration: u16,
) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 6) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&6u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    // WINDOW2 is required by some consumers; keep defaults.
    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: currency
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_currency, 1234.56),
    );
    // A2: percent
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(1, 0, xf_percent, 0.1234),
    );
    // A3: date (serial)
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(2, 0, xf_date, 45123.0),
    );
    // A4: time (serial fraction)
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(3, 0, xf_time, 0.5));
    // A5: duration (serial days; 1.5 days = 36 hours)
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(4, 0, xf_duration, 1.5),
    );
    // A6: BLANK cell with non-General format (percent)
    push_record(&mut sheet, RECORD_BLANK, &blank_cell(5, 0, xf_percent));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u16).to_le_bytes());
    out.extend_from_slice(data);
}

fn bof(dt: u16) -> [u8; 16] {
    // BOF record payload (BIFF8).
    // [0..2]  BIFF version (0x0600)
    // [2..4]  stream type (dt)
    // Remaining fields are build/version metadata; keep stable defaults.
    let mut out = [0u8; 16];
    out[0..2].copy_from_slice(&BOF_VERSION_BIFF8.to_le_bytes());
    out[2..4].copy_from_slice(&dt.to_le_bytes());
    out[4..6].copy_from_slice(&0x0DBBu16.to_le_bytes()); // build
    out[6..8].copy_from_slice(&0x07CCu16.to_le_bytes()); // year (1996)
    out
}

fn bof_biff5(dt: u16) -> [u8; 8] {
    // BOF record payload (BIFF5).
    // [0..2]  BIFF version (0x0500)
    // [2..4]  stream type (dt)
    // [4..6]  build identifier
    // [6..8]  build year
    let mut out = [0u8; 8];
    out[0..2].copy_from_slice(&0x0500u16.to_le_bytes());
    out[2..4].copy_from_slice(&dt.to_le_bytes());
    out[4..6].copy_from_slice(&0x0DBBu16.to_le_bytes()); // build
    out[6..8].copy_from_slice(&0x07CCu16.to_le_bytes()); // year (1996)
    out
}

fn window1() -> [u8; 18] {
    // WINDOW1 record payload (BIFF8, 18 bytes).
    // Keep fields mostly zeroed; Excel tolerates this and so does calamine.
    let mut out = [0u8; 18];
    // cTabSel = 1
    out[14..16].copy_from_slice(&1u16.to_le_bytes());
    // wTabRatio = 600 (arbitrary non-zero)
    out[16..18].copy_from_slice(&600u16.to_le_bytes());
    out
}

fn window2_with_grbit(grbit: u16) -> [u8; 18] {
    let mut out = [0u8; 18];
    out[0..2].copy_from_slice(&grbit.to_le_bytes());
    out
}

fn window1_with_active_tab(active_tab: u16) -> [u8; 18] {
    let mut out = window1();
    // iTabCur at offset 10.
    out[10..12].copy_from_slice(&active_tab.to_le_bytes());
    out
}

fn window1_with_geometry(x: i16, y: i16, dx: u16, dy: u16, grbit: u16) -> [u8; 18] {
    let mut out = window1();
    out[0..2].copy_from_slice(&x.to_le_bytes());
    out[2..4].copy_from_slice(&y.to_le_bytes());
    out[4..6].copy_from_slice(&dx.to_le_bytes());
    out[6..8].copy_from_slice(&dy.to_le_bytes());
    out[8..10].copy_from_slice(&grbit.to_le_bytes());
    out
}

fn scl(num: u16, denom: u16) -> [u8; 4] {
    let mut out = [0u8; 4];
    out[0..2].copy_from_slice(&num.to_le_bytes());
    out[2..4].copy_from_slice(&denom.to_le_bytes());
    out
}

fn pane(x: u16, y: u16, rw_top: u16, col_left: u16, pnn_act: u16) -> [u8; 10] {
    let mut out = [0u8; 10];
    out[0..2].copy_from_slice(&x.to_le_bytes());
    out[2..4].copy_from_slice(&y.to_le_bytes());
    out[4..6].copy_from_slice(&rw_top.to_le_bytes());
    out[6..8].copy_from_slice(&col_left.to_le_bytes());
    out[8..10].copy_from_slice(&pnn_act.to_le_bytes());
    out
}

fn selection_single_cell(pane: u8, row: u16, col: u16) -> Vec<u8> {
    // SELECTION record payload (best-effort BIFF8 layout):
    // [pnn:u8][rwActive:u16][colActive:u16][irefActive:u16][cref:u16][RefU]
    // RefU: [rwFirst:u16][rwLast:u16][colFirst:u8][colLast:u8]
    let mut out = Vec::<u8>::new();
    out.push(pane);
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // irefActive
    out.extend_from_slice(&1u16.to_le_bytes()); // cref
    out.extend_from_slice(&row.to_le_bytes()); // rwFirst
    out.extend_from_slice(&row.to_le_bytes()); // rwLast
    out.push(col as u8); // colFirst
    out.push(col as u8); // colLast
    out
}

fn build_view_state_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1_with_active_tab(1)); // activeTab = 1 (second sheet)
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // Two worksheets.
    let boundsheet1_start = globals.len();
    let mut boundsheet1 = Vec::<u8>::new();
    boundsheet1.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet1.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet1, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet1);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    let boundsheet2_start = globals.len();
    let mut boundsheet2 = Vec::<u8>::new();
    boundsheet2.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet2.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet2, "Sheet2");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet2);
    let boundsheet2_offset_pos = boundsheet2_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // Sheet1: minimal.
    let sheet1_offset = globals.len();
    let sheet1 = {
        let mut sheet = Vec::<u8>::new();
        push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));
        let mut dims = Vec::<u8>::new();
        dims.extend_from_slice(&0u32.to_le_bytes()); // first row
        dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
        dims.extend_from_slice(&0u16.to_le_bytes()); // first col
        dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
        dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
        push_record(&mut sheet, RECORD_DIMENSIONS, &dims);
        push_record(&mut sheet, RECORD_WINDOW2, &window2());
        push_record(&mut sheet, RECORD_EOF, &[]);
        sheet
    };
    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet1);

    // Sheet2: view state records (zoom/freeze/selection).
    let sheet2_offset = globals.len();
    let sheet2 = {
        let mut sheet = Vec::<u8>::new();
        push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

        // DIMENSIONS: rows [0, 3) cols [0, 3) so C3 exists.
        let mut dims = Vec::<u8>::new();
        dims.extend_from_slice(&0u32.to_le_bytes()); // first row
        dims.extend_from_slice(&3u32.to_le_bytes()); // last row + 1
        dims.extend_from_slice(&0u16.to_le_bytes()); // first col
        dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
        dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
        push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

        // WINDOW2: frozen panes flag set; hide gridlines/headings/zeros to exercise flags parsing.
        let grbit: u16 = 0x0008; // fFrozen
        push_record(&mut sheet, RECORD_WINDOW2, &window2_with_grbit(grbit));

        // 200% zoom.
        push_record(&mut sheet, RECORD_SCL, &scl(200, 100));

        // Freeze first row and first column (top-left cell for bottom-right pane is B2).
        push_record(&mut sheet, RECORD_PANE, &pane(1, 1, 1, 1, 0));

        // Active cell C3 (row=2, col=2) in the bottom-right pane.
        push_record(
            &mut sheet,
            RECORD_SELECTION,
            &selection_single_cell(0, 2, 2),
        );

        push_record(&mut sheet, RECORD_EOF, &[]);
        sheet
    };
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet2);

    globals
}

fn build_workbook_window_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());

    // WINDOW1 window geometry/state.
    let x: i16 = 120;
    let y: i16 = 240;
    let width: u16 = 800;
    let height: u16 = 600;
    let grbit: u16 = 0x0002; // fIconic (minimized)
    push_record(
        &mut globals,
        RECORD_WINDOW1,
        &window1_with_geometry(x, y, width, height, grbit),
    );

    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF (required by some readers).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // Worksheet substream: minimal with DIMENSIONS + WINDOW2.
    let sheet_offset = globals.len();
    let sheet = build_empty_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_window_geometry_workbook_stream() -> Vec<u8> {
    // WINDOW1 grbit flags.
    const WINDOW1_GRBIT_MAXIMIZED: u16 = 0x0040;

    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());

    // x=100, y=200, width=300, height=400, maximized.
    let window1 = window1_with_geometry(100, 200, 300, 400, WINDOW1_GRBIT_MAXIMIZED);
    push_record(&mut globals, RECORD_WINDOW1, &window1);

    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF (required by some readers).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // Worksheet substream: minimal with DIMENSIONS + WINDOW2.
    let sheet_offset = globals.len();
    let sheet = build_empty_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_window_hidden_workbook_stream() -> Vec<u8> {
    // WINDOW1 grbit flags.
    const WINDOW1_GRBIT_HIDDEN: u16 = 0x0001;

    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());

    // x=111, y=222, width=333, height=444, hidden.
    let window1 = window1_with_geometry(111, 222, 333, 444, WINDOW1_GRBIT_HIDDEN);
    push_record(&mut globals, RECORD_WINDOW1, &window1);

    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF (required by some readers).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // Worksheet substream: minimal with DIMENSIONS + WINDOW2.
    let sheet_offset = globals.len();
    let sheet = build_empty_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Filter");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`).
    let filter_db_rgce = ptg_area(0, 4, 0, 2); // $A$1:$C$5
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 5) cols [0, 3) so A1:C5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 3 (A..C).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &3u16.to_le_bytes());
    // FILTERMODE: present (no payload).
    push_record(&mut sheet, RECORD_FILTERMODE, &[]);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_sort_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "FilterSort");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`).
    let filter_db_rgce = ptg_area(0, 4, 0, 2); // $A$1:$C$5
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 5) cols [0, 3) so A1:C5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 3 (A..C).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &3u16.to_le_bytes());

    // SORT record: sort the filtered range A1:C5 by column B descending, with a header row.
    push_record(
        &mut sheet,
        RECORD_SORT,
        &sort_record_payload(
            0,
            4, // rows (rwFirst..rwLast) => A1:C5
            0,
            2,               // cols (colFirst..colLast)
            true,            // has header row
            &[(1u16, true)], // key: column B descending
        ),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter12_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Filter12");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`).
    let filter_db_rgce = ptg_area(0, 4, 0, 2); // $A$1:$C$5
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 5) cols [0, 3) so A1:C5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 3 (A..C).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &3u16.to_le_bytes());

    // AutoFilter12 record with a simple multi-value filter on the first column (colId=0).
    //
    // Layout (best-effort):
    //   FrtHeader (8 bytes): rt, grbitFrt, reserved
    //   colId (u16)
    //   cVals (u16)
    //   cVals * XLUnicodeString
    let mut af12 = Vec::<u8>::new();
    af12.extend_from_slice(&RECORD_AUTOFILTER12.to_le_bytes()); // FrtHeader.rt
    af12.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    af12.extend_from_slice(&0u32.to_le_bytes()); // reserved
    af12.extend_from_slice(&0u16.to_le_bytes()); // colId
    af12.extend_from_slice(&2u16.to_le_bytes()); // cVals
    write_unicode_string(&mut af12, "Alice");
    write_unicode_string(&mut af12, "Bob");
    push_record(&mut sheet, RECORD_AUTOFILTER12, &af12);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_calamine_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF.
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Filter");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // Minimal external reference tables so calamine can decode 3D references in the NAME formula.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(1));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // Workbook-scoped `_xlnm._FilterDatabase` referencing Filter!$A$1:$C$5.
    let rgce = ptg_area3d(0, 0, 4, 0, 2);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record_calamine_compat(XLNM_FILTER_DATABASE, &rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 5) cols [0, 3) so A1:C5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_defined_names_builtins_workbook_stream() -> Vec<u8> {
    // Build workbook globals containing two sheets plus a handful of `NAME` records scoped to
    // individual sheets.
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF (required by some readers).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Two worksheets.
    let boundsheet1_start = globals.len();
    let mut boundsheet1 = Vec::<u8>::new();
    boundsheet1.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet1.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet1, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet1);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    let boundsheet2_start = globals.len();
    let mut boundsheet2 = Vec::<u8>::new();
    boundsheet2.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet2.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet2, "Sheet2");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet2);
    let boundsheet2_offset_pos = boundsheet2_start + 4;

    // Minimal EXTERNSHEET table with two internal sheet entries so we can encode 3D references.
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0), (1, 1)]),
    );

    // Built-in defined names (`NAME` records).
    //
    // Print_Area on Sheet1: Sheet1!$A$1:$A$2,Sheet1!$C$1:$C$2 (hidden).
    let print_area_rgce = [
        ptg_area3d(0, 0, 1, 0, 0),
        ptg_area3d(0, 0, 1, 2, 2),
        vec![0x10], // PtgUnion
    ]
    .concat();
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x06, &print_area_rgce),
    );

    // Print_Titles on Sheet2: Sheet2!$1:$1,Sheet2!$A:$A (not hidden).
    let print_titles_rgce = [
        // Whole-row area: row=1, cols=all (0..255).
        ptg_area3d(1, 0, 0, 0, 0x00FF),
        // Whole-column area: col=A, rows=all (0..65535).
        ptg_area3d(1, 0, 0xFFFF, 0, 0),
        vec![0x10], // PtgUnion
    ]
    .concat();
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(false, 2, 0x07, &print_titles_rgce),
    );

    // _FilterDatabase on Sheet1: Sheet1!$A$1:$C$10 (hidden).
    let filter_db_rgce = ptg_area3d(0, 0, 9, 0, 2);
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet substreams -------------------------------------------------------
    let sheet1_offset = globals.len();
    let sheet1 = build_empty_sheet_stream(xf_general);
    let sheet2_offset = sheet1_offset + sheet1.len();
    let sheet2 = build_empty_sheet_stream(xf_general);

    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet1);
    globals.extend_from_slice(&sheet2);

    globals
}

fn build_defined_names_builtins_chkey_mismatch_workbook_stream() -> Vec<u8> {
    // Same as `build_defined_names_builtins_workbook_stream`, but with a mismatch between `chKey`
    // and the stored built-in name id byte in `rgchName`.
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (style XFs only).
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF (required by some readers).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Two worksheets.
    let boundsheet1_start = globals.len();
    let mut boundsheet1 = Vec::<u8>::new();
    boundsheet1.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet1.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet1, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet1);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    let boundsheet2_start = globals.len();
    let mut boundsheet2 = Vec::<u8>::new();
    boundsheet2.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet2.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet2, "Sheet2");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet2);
    let boundsheet2_offset_pos = boundsheet2_start + 4;

    // Minimal EXTERNSHEET table with two internal sheet entries so we can encode 3D references.
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0), (1, 1)]),
    );

    // Built-in defined names (`NAME` records).
    //
    // Print_Area on Sheet1: Sheet1!$A$1:$A$2,Sheet1!$C$1:$C$2 (hidden).
    let print_area_rgce = [
        ptg_area3d(0, 0, 1, 0, 0),
        ptg_area3d(0, 0, 1, 2, 2),
        vec![0x10], // PtgUnion
    ]
    .concat();
    push_record(
        &mut globals,
        RECORD_NAME,
        // Store the correct built-in id in `rgchName`, but populate `chKey` with an arbitrary,
        // non-zero byte (as some writers do). Excel appears to prefer `rgchName` for built-in
        // names, so the importer should interpret this as Print_Area (0x06), not Print_Titles.
        &builtin_name_record_with_chkey(true, 1, b'X', 0x06, &print_area_rgce),
    );

    // Print_Titles on Sheet2: Sheet2!$1:$1,Sheet2!$A:$A (not hidden).
    let print_titles_rgce = [
        // Whole-row area: row=1, cols=all (0..255).
        ptg_area3d(1, 0, 0, 0, 0x00FF),
        // Whole-column area: col=A, rows=all (0..65535).
        ptg_area3d(1, 0, 0xFFFF, 0, 0),
        vec![0x10], // PtgUnion
    ]
    .concat();
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(false, 2, 0x07, &print_titles_rgce),
    );

    // _FilterDatabase on Sheet1: Sheet1!$A$1:$C$10 (hidden).
    let filter_db_rgce = ptg_area3d(0, 0, 9, 0, 2);
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet substreams -------------------------------------------------------
    let sheet1_offset = globals.len();
    let sheet1 = build_empty_sheet_stream(xf_general);
    let sheet2_offset = sheet1_offset + sheet1.len();
    let sheet2 = build_empty_sheet_stream(xf_general);

    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet1);
    globals.extend_from_slice(&sheet2);

    globals
}

fn builtin_name_record(hidden: bool, itab: u16, builtin_id: u8, rgce: &[u8]) -> Vec<u8> {
    builtin_name_record_with_chkey(hidden, itab, 0, builtin_id, rgce)
}

fn builtin_name_record_with_chkey(
    hidden: bool,
    itab: u16,
    ch_key: u8,
    builtin_id: u8,
    rgce: &[u8],
) -> Vec<u8> {
    // BIFF8 NAME record [MS-XLS] 2.4.150.
    const NAME_FLAG_HIDDEN: u16 = 0x0001;
    const NAME_FLAG_BUILTIN: u16 = 0x0020;

    let mut out = Vec::<u8>::new();

    let mut grbit: u16 = NAME_FLAG_BUILTIN;
    if hidden {
        grbit |= NAME_FLAG_HIDDEN;
    }

    out.extend_from_slice(&grbit.to_le_bytes()); // grbit
    out.push(ch_key); // chKey
    out.push(1); // cch (built-in name id length)
    out.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
    out.extend_from_slice(&0u16.to_le_bytes()); // ixals
    out.extend_from_slice(&itab.to_le_bytes()); // itab (1-based sheet index, 0=workbook)
    out.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

    out.push(builtin_id); // rgchName (built-in id)
    out.extend_from_slice(rgce); // rgce
    out
}

fn ptg_area(rw_first: u16, rw_last: u16, col_first: u16, col_last: u16) -> Vec<u8> {
    // PtgArea token (ref class): [ptg=0x25][rwFirst][rwLast][colFirst][colLast]
    let mut out = Vec::<u8>::new();
    out.push(0x25);
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());
    out
}

fn sort_record_payload(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    has_header: bool,
    keys: &[(u16, bool)], // (col, descending)
) -> Vec<u8> {
    // Minimal BIFF8 SORT record payload (classic SORT).
    //
    // This matches the best-effort parser in `formula-xls`:
    // - Ref8U (rwFirst, rwLast, colFirst, colLast)
    // - grbit (header flag)
    // - cKeys
    // - 3x key column indices (u16, 0xFFFF for unused)
    // - 3x sort orders (u16, 0=ascending, 1=descending)
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());

    let mut grbit: u16 = 0;
    if has_header {
        grbit |= 0x0001;
    }
    out.extend_from_slice(&grbit.to_le_bytes());

    let key_count: u16 = keys.len().min(3) as u16;
    out.extend_from_slice(&key_count.to_le_bytes());

    // Key columns.
    for i in 0..3usize {
        let col = keys.get(i).map(|(c, _)| *c).unwrap_or(0xFFFF);
        out.extend_from_slice(&col.to_le_bytes());
    }

    // Key orders.
    for i in 0..3usize {
        let descending = keys.get(i).map(|(_, d)| *d).unwrap_or(false);
        let order: u16 = if descending { 1 } else { 0 };
        out.extend_from_slice(&order.to_le_bytes());
    }

    out
}
fn window2() -> [u8; 18] {
    // WINDOW2 record payload (BIFF8). Most fields can be zero for our fixtures.
    let mut out = [0u8; 18];
    let grbit: u16 = 0x02B6;
    out[0..2].copy_from_slice(&grbit.to_le_bytes());
    out
}

fn font(name: &str) -> Vec<u8> {
    font_with_options(FontOptions {
        name,
        height_twips: 200, // 10pt
        weight: 400,
        italic: false,
        underline: false,
        strike: false,
        color_idx: COLOR_AUTOMATIC,
    })
}

fn font_biff5(name: &str) -> Vec<u8> {
    // Minimal BIFF5 FONT record payload.
    //
    // The BIFF5 structure is largely compatible with BIFF8 for the fixed-size fields, but the
    // font name is stored as an ANSI short string (no BIFF8 Unicode flags byte).
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&200u16.to_le_bytes()); // height (10pt)
    out.extend_from_slice(&0u16.to_le_bytes()); // option flags
    out.extend_from_slice(&COLOR_AUTOMATIC.to_le_bytes()); // color
    out.extend_from_slice(&400u16.to_le_bytes()); // weight
    out.extend_from_slice(&0u16.to_le_bytes()); // escapement
    out.push(0); // underline
    out.push(0); // family
    out.push(0); // charset
    out.push(0); // reserved
    write_short_ansi_string(&mut out, name);
    out
}

struct FontOptions<'a> {
    name: &'a str,
    height_twips: u16,
    weight: u16,
    italic: bool,
    underline: bool,
    strike: bool,
    color_idx: u16,
}

fn font_with_options(opts: FontOptions<'_>) -> Vec<u8> {
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&opts.height_twips.to_le_bytes()); // height

    let mut flags: u16 = 0;
    if opts.italic {
        flags |= 0x0002;
    }
    if opts.strike {
        flags |= 0x0008;
    }
    out.extend_from_slice(&flags.to_le_bytes()); // option flags

    out.extend_from_slice(&opts.color_idx.to_le_bytes()); // color
    out.extend_from_slice(&opts.weight.to_le_bytes()); // weight
    out.extend_from_slice(&0u16.to_le_bytes()); // escapement
    out.push(if opts.underline { 1 } else { 0 }); // underline
    out.push(0); // family
    out.push(0); // charset
    out.push(0); // reserved
    write_short_unicode_string(&mut out, opts.name);
    out
}

fn palette(colors: &[(u8, u8, u8)]) -> Vec<u8> {
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&(colors.len() as u16).to_le_bytes());
    for &(r, g, b) in colors {
        out.push(r);
        out.push(g);
        out.push(b);
        out.push(0); // reserved
    }
    out
}

fn format_record(id: u16, code: &str) -> Vec<u8> {
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&id.to_le_bytes());
    write_unicode_string(&mut out, code);
    out
}

fn xf_record(font_idx: u16, fmt_idx: u16, is_style_xf: bool) -> [u8; 20] {
    let mut out = [0u8; 20];
    out[0..2].copy_from_slice(&font_idx.to_le_bytes());
    out[2..4].copy_from_slice(&fmt_idx.to_le_bytes());

    // Protection / type / parent:
    // bit0: locked (1)
    // bit2: xfType (1 = style XF, 0 = cell XF)
    // bits4-15: parent style XF index (0)
    let flags: u16 = XF_FLAG_LOCKED | if is_style_xf { XF_FLAG_STYLE } else { 0 };
    out[4..6].copy_from_slice(&flags.to_le_bytes());

    // Default BIFF8 alignment: General + Bottom.
    out[6] = 0x20;

    // Attribute flags: apply all so fixture cell XFs don't rely on inheritance.
    out[9] = 0x3F;
    out
}

fn xf_record_biff5(font_idx: u16, fmt_idx: u16, is_style_xf: bool) -> [u8; 16] {
    // Minimal BIFF5 XF record payload (16 bytes).
    //
    // We only need enough structure for calamine to build a value grid for our fixtures; styling
    // is not the focus of BIFF5 comment tests.
    let mut out = [0u8; 16];
    out[0..2].copy_from_slice(&font_idx.to_le_bytes());
    out[2..4].copy_from_slice(&fmt_idx.to_le_bytes());

    let flags: u16 = XF_FLAG_LOCKED | if is_style_xf { XF_FLAG_STYLE } else { 0 };
    out[4..6].copy_from_slice(&flags.to_le_bytes());

    // Default alignment: General + Bottom.
    out[6] = 0x20;
    // Attribute flags: apply all.
    out[9] = 0x3F;
    out
}

fn xf_record_rich() -> [u8; 20] {
    // This encoding matches the best-effort BIFF8 XF parser in `formula-xls`.
    //
    // Font index = 1 (2nd FONT record)
    // Number format = 10 (built-in percent "0.00%")
    // Protection: unlocked + hidden
    // Alignment: Center + Wrap + Top
    // Rotation: 45 degrees
    // Indent: 2
    // Borders: Thin, palette index 9 (green)
    // Fill: Solid, fg=8 (red), bg=9 (green)
    let mut out = [0u8; 20];
    out[0..2].copy_from_slice(&1u16.to_le_bytes()); // ifnt
    out[2..4].copy_from_slice(&10u16.to_le_bytes()); // ifmt

    // flags (protection/type/parent)
    out[4..6].copy_from_slice(&0x0002u16.to_le_bytes()); // hidden=1, locked=0, cell XF, parent=0

    // alignment / rotation / text props / attribute flags
    out[6] = 0x0A; // horiz=center (2), wrap, vert=top (0)
    out[7] = 45; // rotation
    out[8] = 0x02; // indent=2
    out[9] = 0x3F; // apply all

    // border + fill
    let border1: u32 = 0x8489_1111;
    let border2: u32 = 0x0222_4489;
    let pattern: u16 = 8 | (9 << 7);

    out[10..14].copy_from_slice(&border1.to_le_bytes());
    out[14..18].copy_from_slice(&border2.to_le_bytes());
    out[18..20].copy_from_slice(&pattern.to_le_bytes());
    out
}

fn xf_record_rich_with_fill_pattern(fill_pattern: u8) -> [u8; 20] {
    let mut out = xf_record_rich();
    let mut border2 = u32::from_le_bytes([out[14], out[15], out[16], out[17]]);
    border2 &= !(0x3F_u32 << 25);
    border2 |= ((fill_pattern as u32) & 0x3F) << 25;
    out[14..18].copy_from_slice(&border2.to_le_bytes());
    out
}
fn number_cell(row: u16, col: u16, xf: u16, v: f64) -> [u8; 14] {
    let mut out = [0u8; 14];
    out[0..2].copy_from_slice(&row.to_le_bytes());
    out[2..4].copy_from_slice(&col.to_le_bytes());
    out[4..6].copy_from_slice(&xf.to_le_bytes());
    out[6..14].copy_from_slice(&v.to_le_bytes());
    out
}

fn blank_cell(row: u16, col: u16, xf: u16) -> [u8; 6] {
    let mut out = [0u8; 6];
    out[0..2].copy_from_slice(&row.to_le_bytes());
    out[2..4].copy_from_slice(&col.to_le_bytes());
    out[4..6].copy_from_slice(&xf.to_le_bytes());
    out
}

fn row_record(row: u16, hidden: bool, outline_level: u8, collapsed: bool) -> [u8; 16] {
    // ROW record payload (BIFF8, 16 bytes).
    let mut out = [0u8; 16];
    out[0..2].copy_from_slice(&row.to_le_bytes());
    // colMic=0, colMac=4 (A..D)
    out[2..4].copy_from_slice(&0u16.to_le_bytes());
    out[4..6].copy_from_slice(&4u16.to_le_bytes());
    // miyRw: default height flag set (0x8000).
    out[6..8].copy_from_slice(&0x8000u16.to_le_bytes());

    let mut options: u16 = 0;
    if hidden {
        options |= ROW_OPTION_HIDDEN;
    }
    options |= ((outline_level as u16) & 0x0007) << 8;
    if collapsed {
        options |= ROW_OPTION_COLLAPSED;
    }
    out[12..14].copy_from_slice(&options.to_le_bytes());
    out
}

fn colinfo_record(
    first_col: u16,
    last_col: u16,
    hidden: bool,
    outline_level: u8,
    collapsed: bool,
) -> [u8; 12] {
    // COLINFO record payload (BIFF8, 12 bytes).
    let mut out = [0u8; 12];
    out[0..2].copy_from_slice(&first_col.to_le_bytes());
    out[2..4].copy_from_slice(&last_col.to_le_bytes());
    // cx: arbitrary non-zero width (8.0 characters * 256).
    out[4..6].copy_from_slice(&2048u16.to_le_bytes());

    let mut options: u16 = 0;
    if hidden {
        options |= COLINFO_OPTION_HIDDEN;
    }
    options |= ((outline_level as u16) & 0x0007) << 8;
    if collapsed {
        options |= COLINFO_OPTION_COLLAPSED;
    }
    out[8..10].copy_from_slice(&options.to_le_bytes());
    out
}

fn write_short_unicode_string(out: &mut Vec<u8>, s: &str) {
    // BIFF8 ShortXLUnicodeString: [cch: u8][flags: u8][chars]
    // `cch` counts UTF-16 code units (not UTF-8 bytes).
    let utf16: Vec<u16> = s.encode_utf16().collect();
    let len: u8 = utf16
        .len()
        .try_into()
        .expect("string too long for u8 length");
    out.push(len);
    if utf16.iter().all(|&ch| ch <= 0x00FF) {
        out.push(0); // compressed (8-bit)
        out.extend(utf16.into_iter().map(|ch| ch as u8));
    } else {
        out.push(1); // uncompressed (16-bit)
        for ch in utf16 {
            out.extend_from_slice(&ch.to_le_bytes());
        }
    }
}

fn write_short_ansi_string(out: &mut Vec<u8>, s: &str) {
    write_short_ansi_bytes(out, s.as_bytes());
}

fn write_short_ansi_bytes(out: &mut Vec<u8>, bytes: &[u8]) {
    // BIFF5 ANSI short string: [cch: u8][rgb: u8 * cch]
    let len: u8 = bytes
        .len()
        .try_into()
        .expect("string too long for u8 length");
    out.push(len);
    out.extend_from_slice(bytes);
}

fn write_unicode_string(out: &mut Vec<u8>, s: &str) {
    // BIFF8 XLUnicodeString: [cch: u16][flags: u8][chars]
    let utf16: Vec<u16> = s.encode_utf16().collect();
    let len: u16 = utf16
        .len()
        .try_into()
        .expect("string too long for u16 length");
    out.extend_from_slice(&len.to_le_bytes());
    if utf16.iter().all(|&ch| ch <= 0x00FF) {
        out.push(0); // compressed (8-bit)
        out.extend(utf16.into_iter().map(|ch| ch as u8));
    } else {
        out.push(1); // uncompressed (16-bit)
        for ch in utf16 {
            out.extend_from_slice(&ch.to_le_bytes());
        }
    }
}

fn supbook_internal(sheet_count: u16) -> Vec<u8> {
    // SUPBOOK record payload [MS-XLS 2.4.271] for "internal" workbook references.
    //
    // `virtPath` is an XLUnicodeString containing a single 0x01 marker character.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&sheet_count.to_le_bytes()); // ctab
    write_unicode_string(&mut out, "\u{0001}");
    out
}

fn supbook_external(workbook_name: &str, sheet_names: &[&str]) -> Vec<u8> {
    // SUPBOOK record payload for an external workbook.
    //
    // Layout:
    //   [ctab: u16]
    //   [virtPath: XLUnicodeString]
    //   ctab * [sheetName: XLUnicodeString]
    let mut out = Vec::<u8>::new();
    let ctab: u16 = sheet_names
        .len()
        .try_into()
        .expect("external sheet name count too large for u16");
    out.extend_from_slice(&ctab.to_le_bytes());
    write_unicode_string(&mut out, workbook_name);
    for &sheet in sheet_names {
        write_unicode_string(&mut out, sheet);
    }
    out
}

fn externname_record(name: &str) -> Vec<u8> {
    // EXTERNNAME record payload (best-effort fixture encoding).
    //
    // The BIFF8 EXTERNNAME structure is complex; the importer’s SUPBOOK parser currently expects
    // the common layout:
    //   [grbit: u16][reserved: u32][cch: u8][rgchName: XLUnicodeStringNoCch]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&0u16.to_le_bytes()); // grbit
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved
    let cch: u8 = name
        .len()
        .try_into()
        .expect("EXTERNNAME too long for u8 cch");
    out.push(cch);
    write_unicode_string_no_cch(&mut out, name);
    out
}

fn formula_cell(row: u16, col: u16, xf: u16, cached_result: f64, rgce: &[u8]) -> Vec<u8> {
    // FORMULA record payload (BIFF8) [MS-XLS 2.4.127].
    //
    // This is a minimal encoding sufficient for calamine to surface the formula text.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&xf.to_le_bytes());
    out.extend_from_slice(&cached_result.to_le_bytes()); // cached formula result (IEEE f64)
    out.extend_from_slice(&0u16.to_le_bytes()); // grbit
    out.extend_from_slice(&0u32.to_le_bytes()); // chn
    out.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
    out.extend_from_slice(rgce);
    out
}

fn shrfmla_record(
    rw_first: u16,
    rw_last: u16,
    col_first: u8,
    col_last: u8,
    rgce: &[u8],
) -> Vec<u8> {
    // SHRFMLA record payload (BIFF8) [MS-XLS 2.4.277].
    //
    // Minimal layout (RefU + cce + rgce):
    //   [rwFirst: u16][rwLast: u16][colFirst: u8][colLast: u8][cce: u16][rgce bytes]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.push(col_first);
    out.push(col_last);
    out.extend_from_slice(&(rgce.len() as u16).to_le_bytes());
    out.extend_from_slice(rgce);
    out
}

fn formula_cell_with_raw_value(
    row: u16,
    col: u16,
    xf: u16,
    value: [u8; 8],
    rgce: &[u8],
) -> Vec<u8> {
    // FORMULA record payload (BIFF8) [MS-XLS 2.4.127].
    //
    // This is a variant of `formula_cell` that allows injecting non-numeric cached result
    // encodings (e.g. a string result marker).
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&xf.to_le_bytes());
    out.extend_from_slice(&value);
    out.extend_from_slice(&0u16.to_le_bytes()); // grbit
    out.extend_from_slice(&0u32.to_le_bytes()); // chn
    out.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
    out.extend_from_slice(rgce);
    out
}

fn sheetext_record_rgb(r: u8, g: u8, b: u8) -> Vec<u8> {
    // BIFF8 `SHEETEXT` is an FRT record that begins with an `FrtHeader` (8 bytes).
    //
    // For the purposes of our importer tests, we only need to include an `XColor` payload at the
    // end of the record:
    // - xclrType (u16) = 2 (RGB)
    // - index (u16) = 0 (unused for RGB)
    // - longRGB (4 bytes) = {r,g,b,0}
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&RECORD_SHEETEXT.to_le_bytes()); // rt
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved

    // SheetExt flags (unused for this fixture).
    out.extend_from_slice(&0u32.to_le_bytes());

    // XColor payload.
    out.extend_from_slice(&2u16.to_le_bytes()); // xclrType = RGB
    out.extend_from_slice(&0u16.to_le_bytes()); // index
    out.extend_from_slice(&[r, g, b, 0]);
    out
}

fn sheetext_record_indexed(idx: u16) -> Vec<u8> {
    // Minimal BIFF8 `SHEETEXT` record storing an indexed `XColor`.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&RECORD_SHEETEXT.to_le_bytes()); // rt
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved

    // SheetExt flags (unused for this fixture).
    out.extend_from_slice(&0u32.to_le_bytes());

    // XColor payload (indexed).
    out.extend_from_slice(&1u16.to_le_bytes()); // xclrType = indexed
    out.extend_from_slice(&idx.to_le_bytes()); // icv
    out.extend_from_slice(&[0u8; 4]); // rgb (unused for indexed)
    out
}

fn feat_hdr_record_sheet_protection_allow_mask(allow_mask: u16) -> Vec<u8> {
    // Minimal BIFF8 `FEATHEADR` (shared feature header) payload encoding used by Excel to persist
    // enhanced worksheet protection options.
    //
    // Layout (best-effort):
    // - FrtHeader (8 bytes): rt/grbitFrt/reserved
    // - isf (u16): shared feature type (0x0002 = enhanced sheet protection)
    // - reserved (u16)
    // - cbHdrData (u32)
    // - rgbHdrData (cbHdrData bytes): allow-flag bitmask
    const ISF_SHEET_PROTECTION: u16 = 0x0002;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&RECORD_FEATHEADR.to_le_bytes()); // rt
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved

    out.extend_from_slice(&ISF_SHEET_PROTECTION.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out.extend_from_slice(&2u32.to_le_bytes()); // cbHdrData
    out.extend_from_slice(&allow_mask.to_le_bytes());
    out
}

fn feat_record_sheet_protection_allow_mask(allow_mask: u16) -> Vec<u8> {
    // Minimal BIFF8 `FEAT` payload for enhanced sheet protection.
    //
    // Layout (best-effort):
    // - FrtHeader (8 bytes)
    // - isf (u16)
    // - reserved (u16)
    // - cbFeatData (u32)
    // - rgbFeatData (cbFeatData bytes): allow-flag bitmask
    const ISF_SHEET_PROTECTION: u16 = 0x0002;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&RECORD_FEAT.to_le_bytes()); // rt
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved

    out.extend_from_slice(&ISF_SHEET_PROTECTION.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out.extend_from_slice(&2u32.to_le_bytes()); // cbFeatData
    out.extend_from_slice(&allow_mask.to_le_bytes());
    out
}

fn feat_record_sheet_protection_allow_mask_malformed(allow_mask: u16) -> Vec<u8> {
    // Like `feat_record_sheet_protection_allow_mask`, but the declared `cbFeatData` length is
    // larger than the payload to exercise best-effort warning behavior.
    const ISF_SHEET_PROTECTION: u16 = 0x0002;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&RECORD_FEAT.to_le_bytes()); // rt
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved

    out.extend_from_slice(&ISF_SHEET_PROTECTION.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out.extend_from_slice(&10u32.to_le_bytes()); // cbFeatData (incorrect)
                                                 // Provide only a 2-byte allow mask even though cbFeatData claims 10 bytes.
    out.extend_from_slice(&allow_mask.to_le_bytes());
    out
}

fn palette_record_with_override(idx: u16, r: u8, g: u8, b: u8) -> Vec<u8> {
    // BIFF8 PALETTE record: ccv + array of `LongRGB` entries.
    //
    // Excel defines 56 custom palette entries that correspond to indexed colors 8..=63.
    // We base the record on the default palette and override one entry for testing.
    let mut entries: Vec<[u8; 4]> = Vec::with_capacity(56);
    for index in 8u16..=63u16 {
        let argb = indexed_color_argb(index).unwrap_or(0xFF000000);
        let rr = ((argb >> 16) & 0xFF) as u8;
        let gg = ((argb >> 8) & 0xFF) as u8;
        let bb = (argb & 0xFF) as u8;
        entries.push([rr, gg, bb, 0]);
    }

    if idx >= 8 && idx <= 63 {
        entries[(idx - 8) as usize] = [r, g, b, 0];
    }

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&(entries.len() as u16).to_le_bytes());
    for entry in entries {
        out.extend_from_slice(&entry);
    }
    out
}
