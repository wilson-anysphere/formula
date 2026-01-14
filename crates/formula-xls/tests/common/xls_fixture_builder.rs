#![allow(dead_code)]

use std::io::{Cursor, Write};

use formula_model::{
    indexed_color_argb, EXCEL_MAX_COLS, XLNM_FILTER_DATABASE, XLNM_PRINT_AREA, XLNM_PRINT_TITLES,
};
use sha1::{Digest as _, Sha1};

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
const RECORD_SST: u16 = 0x00FC;
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
/// TABLE [MS-XLS 2.4.313]
///
/// Used for What-If Analysis data tables (`TABLE()` formulas) referenced by BIFF8 `PtgTbl` tokens.
const RECORD_TABLE: u16 = 0x0236;
/// SHRFMLA [MS-XLS 2.4.277] stores a shared formula (rgce) for a range.
const RECORD_SHRFMLA: u16 = 0x04BC;
const RECORD_LABELSST: u16 = 0x00FD;
/// ARRAY [MS-XLS 2.4.19] stores an array (CSE) formula (rgce) for a range.
const RECORD_ARRAY: u16 = 0x0221;
const RECORD_HLINK: u16 = 0x01B8;
const RECORD_AUTOFILTERINFO: u16 = 0x009D;
const RECORD_AUTOFILTER: u16 = 0x009E;
const RECORD_SORT: u16 = 0x0090;
// Excel 2007+ may store newer sort semantics in BIFF8 via future records.
const RECORD_SORT12: u16 = 0x0890;
const RECORD_SORTDATA12: u16 = 0x0895;
const RECORD_FILTERMODE: u16 = 0x009B;
// Excel 2007+ may store newer filter semantics in BIFF8 via future records.
const RECORD_AUTOFILTER12: u16 = 0x087E;
// Some FRT records (including AutoFilter12/Sort12/SortData12) can be continued via `ContinueFrt12`.
const RECORD_CONTINUEFRT12: u16 = 0x087F;
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

// SETUP record `grbit` flags (BIFF8).
// See [MS-XLS] 2.4.296 (SETUP).
const SETUP_GRBIT_F_PORTRAIT: u16 = 0x0002; // 0=landscape, 1=portrait
const SETUP_GRBIT_F_NOPLS: u16 = 0x0004;
const SETUP_GRBIT_F_NOORIENT: u16 = 0x0040;

// WSBOOL options.
const WSBOOL_OPTION_FIT_TO_PAGE: u16 = 0x0100;
// ---------------------------------------------------------------------------
// BIFF8 AutoFilter helpers (AUTOFILTER record payload)
// ---------------------------------------------------------------------------
//
// These helpers cover the minimal subset of [MS-XLS] AUTOFILTER/DOPER needed by
// our importer tests:
// - a single custom-filter criterion against a string or number
// - no advanced filters (Top10, dynamic, etc)
//
// The importer’s real decoder (`biff/autofilter_criteria.rs`) expects the BIFF8
// layout:
//   [iEntry:u16][grbit:u16][DOPER1:8][DOPER2:8][XLUnicodeString...]
// where string operand text is stored *after* the fixed-size DOPER payloads.
//
// Operator codes follow `AutoFilterOp::from_biff_code` in `autofilter_criteria.rs`.
const AUTOFILTER_OP_NONE: u8 = 0;
const AUTOFILTER_OP_BETWEEN: u8 = 1;
const AUTOFILTER_OP_EQUAL: u8 = 3;
const AUTOFILTER_OP_GREATER_THAN: u8 = 5;
const AUTOFILTER_OP_LESS_THAN: u8 = 6;

// DOPER "vt" values are based on [MS-XLS] / VARIANT type tags. The importer is
// tolerant, but we aim to emit the canonical values we see in the wild.
const AUTOFILTER_VT_EMPTY: u8 = 0;
const AUTOFILTER_VT_NUMBER: u8 = 5; // VT_R8 (stored as RK by some producers)
                                    // Many BIFF8 AUTOFILTER records use vt=4 for string operands (string stored as trailing
                                    // XLUnicodeString). The importer also supports vt=8, but we prefer vt=4 here to exercise that path.
const AUTOFILTER_VT_STRING: u8 = 4;
// Common boolean DOPER encodings observed in BIFF8 AutoFilter records.
const AUTOFILTER_VT_BOOL: u8 = 6;

// AUTOFILTER.grbit flag: combine DOPER1/DOPER2 with AND when set, else OR.
const AUTOFILTER_GRBIT_AND: u16 = 0x0001;

const BOF_VERSION_BIFF8: u16 = 0x0600;
const BOF_DT_WORKBOOK_GLOBALS: u16 = 0x0005;
const BOF_DT_WORKSHEET: u16 = 0x0010;

const XF_FLAG_LOCKED: u16 = 0x0001;
const XF_FLAG_STYLE: u16 = 0x0004;

const COLOR_AUTOMATIC: u16 = 0x7FFF;

// FORMULA.grbit flag indicating a shared formula (`SHRFMLA` follows the base FORMULA record).
const FORMULA_FLAG_SHARED: u16 = 0x0008;

// ExtRst "rt" type for phonetic blocks inside BIFF8 `XLUnicodeRichExtendedString.ExtRst`.
const EXT_RST_TYPE_PHONETIC: u16 = 0x0001;

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

/// Build a minimal BIFF8 `.xls` fixture that stores phonetic (furigana) metadata in the workbook
/// shared string table (SST `ExtRst`) and references it from a worksheet `LABELSST` cell.
pub fn build_sst_phonetic_fixture_xls(phonetic_text: &str) -> Vec<u8> {
    let workbook_stream = build_sst_phonetic_workbook_stream(phonetic_text);

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

/// Build a BIFF8 `.xls` fixture containing many worksheets with malformed `SELECTION` records.
///
/// Each malformed record yields an import warning via the BIFF worksheet view-state parser, which
/// is internally capped. By distributing warnings across multiple sheets we can exceed the `.xls`
/// importer's *global* warning cap.
pub fn build_many_malformed_selection_records_fixture_xls(
    sheet_count: usize,
    records_per_sheet: usize,
) -> Vec<u8> {
    let workbook_stream =
        build_many_malformed_selection_records_workbook_stream(sheet_count, records_per_sheet);

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

fn build_many_malformed_selection_records_workbook_stream(
    sheet_count: usize,
    records_per_sheet: usize,
) -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table. Many readers expect at least 16 style XFs before any cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for sheet_idx in 0..sheet_count {
        let sheet_name = format!("Sheet{}", sheet_idx + 1);
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, &sheet_name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    for boundsheet_offset_pos in boundsheet_offset_positions {
        let sheet_offset = globals.len();
        let sheet = build_many_malformed_selection_records_sheet_stream(xf_cell, records_per_sheet);
        globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
            .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
        globals.extend_from_slice(&sheet);
    }

    globals
}

fn build_many_malformed_selection_records_sheet_stream(
    xf_cell: u16,
    record_count: usize,
) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    // A valid WINDOW2 record so the view-state parser runs and can associate warnings with the
    // worksheet.
    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Malformed SELECTION records (empty payload) to generate many warnings.
    for _ in 0..record_count {
        push_record(&mut sheet, RECORD_SELECTION, &[]);
    }

    // Provide a single cell so calamine reports a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
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

/// Build a BIFF8 `.xls` fixture containing a shared formula (`SHRFMLA`) whose base `rgce`
/// uses `PtgRef` with the row/col-relative flags set.
///
/// This exercises BIFF8 shared-formula materialization: follower cells must shift the stored
/// row/col coordinates by the delta between the base cell and the follower cell.
pub fn build_shared_formula_ptgref_relative_flags_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgref_relative_flags_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula (`SHRFMLA`) whose base `rgce`
/// uses `PtgArea` with row/col-relative flags set.
///
/// Expected decoded formulas:
/// - B1: `SUM(A1:A2)`
/// - B2: `SUM(A2:A3)`
pub fn build_shared_formula_ptgarea_relative_flags_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgarea_relative_flags_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture where a shared formula is defined only by a `SHRFMLA` record
/// (no `FORMULA` records in the shared range), and the shared `rgce` uses `PtgRef` with the
/// row/col-relative flags set.
///
/// Expected decoded formulas:
/// - B1: `A1+1`
/// - B2: `A2+1`
pub fn build_shared_formula_shrfmla_only_ptgref_relative_flags_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_shrfmla_only_ptgref_relative_flags_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture where a shared formula is defined only by a `SHRFMLA` record
/// (no `FORMULA` records in the shared range), and the shared `rgce` uses `PtgArea` with
/// row/col-relative flags set.
///
/// Expected decoded formulas:
/// - B1: `SUM(A1:A2)`
/// - B2: `SUM(A2:A3)`
pub fn build_shared_formula_shrfmla_only_ptgarea_relative_flags_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_shared_formula_shrfmla_only_ptgarea_relative_flags_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a malformed shared-formula pattern like
/// [`build_shared_formula_ptgexp_missing_shrfmla_fixture_xls`], but where the follower-cell `PtgExp`
/// stores a non-standard payload width (row u32 + col u16).
///
/// The `.xls` importer should still recover the follower formula by materializing from the base
/// cell's `FORMULA.rgce`.
pub fn build_shared_formula_ptgexp_wide_payload_missing_shrfmla_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgexp_wide_payload_missing_shrfmla_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing shared-formula groups where the base cell's `FORMULA`
/// record is missing or degenerate (`PtgExp`).
///
/// Some `.xls` writers rely entirely on the `SHRFMLA` record (shared formula definition) to store
/// the shared rgce token stream and only emit `PtgExp` in cells within the shared formula range.
///
/// This fixture contains two sheets:
/// - `MissingBase`: the shared formula range is `B1:B2`, but `B1` has **no** `FORMULA` record. `B2`
///   contains a `FORMULA` record with `PtgExp(B1)`.
/// - `DegenerateBase`: `B1` has a `FORMULA` record but its rgce is `PtgExp(B1)` (self-reference);
///   `B2` contains `PtgExp(B1)`.
///
/// In both cases, the `SHRFMLA` record defines the shared rgce for `A(row)+1`:
/// - `B1` → `A1+1`
/// - `B2` → `A2+1`
pub fn build_shared_formula_shrfmla_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_shrfmla_only_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a shared formula where the `SHRFMLA` record is split
/// across a `CONTINUE` boundary inside a `PtgStr` token.
///
/// This exercises fragment-aware `rgce` parsing for SHRFMLA-only shared formulas. The sheet is
/// named `ShrfmlaContinue` and defines a shared formula over `B1:B2` with no base-cell `FORMULA`
/// record:
/// - `B1`: `A1&\"ABCDE\"` (recovered from SHRFMLA)
/// - `B2`: `A2&\"ABCDE\"` (PtgExp to B1)
pub fn build_shared_formula_shrfmla_only_continued_ptgstr_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_shrfmla_only_continued_ptgstr_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a SHRFMLA-only shared formula whose token stream
/// includes a `PtgArray` constant (backed by trailing `rgcb` bytes).
///
/// Some writers omit the base cell's `FORMULA` record entirely but still use `PtgArray`. Our
/// SHRFMLA-only shared-formula recovery must preserve the `rgcb` bytes so array constants decode to
/// `{...}` literals rather than `#UNKNOWN!`.
///
/// Sheet name: `ShrfmlaArray`
/// Shared formula range: `B1:B2`
/// - `B1`: recovered from SHRFMLA (no FORMULA record)
/// - `B2`: `PtgExp(B1)`
pub fn build_shared_formula_shrfmla_only_ptgarray_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_shrfmla_only_ptgarray_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a `PtgExp` shared formula whose backing `SHRFMLA` record
/// includes a `PtgArray` constant with trailing `rgcb` bytes.
///
/// This fixture is constructed so that the BIFF8 formula-override pass (which is intentionally
/// conservative and only applies overrides when formula decoding emits **no warnings**) will skip
/// overriding the `PtgExp` cell. This forces the importer to rely on the earlier
/// `recover_ptgexp_formulas_from_shrfmla_and_array` path, which must decode `PtgArray` using the
/// `SHRFMLA` record’s `rgcb` bytes.
///
/// To achieve this, the shared formula body intentionally encodes `SUM` using a `PtgFunc` token
/// rather than the canonical `PtgFuncVar`, which triggers a decode warning while still yielding a
/// parseable formula text.
///
/// Sheet name: `ShrfmlaArrayWarn`
/// Shared formula range: `B1:B2`
/// - `B1`: no FORMULA record (recovered from SHRFMLA)
/// - `B2`: `PtgExp(B1)` (must be recovered via `PtgExp` resolver using SHRFMLA `rgcb`)
pub fn build_shared_formula_ptgexp_ptgarray_warning_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgexp_ptgarray_warning_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture that exercises ambiguity in SHRFMLA range-header parsing when a
/// `RefU` header is followed by a small non-zero `cUse` value.
///
/// Some decoders (including older versions of this crate) would mistakenly treat the first 8 bytes
/// of `RefU + cUse` as a `Ref8` range header when the shared range is `A..A`, causing the decoded
/// range to spuriously include additional columns.
///
/// This fixture constructs two SHRFMLA records:
/// - `A1:A2` (RefU + `cUse=2`), shared rgce is `"LEFT"`
/// - `B1:B10`, shared rgce is `C(row)+1`
///
/// Cell `B2` stores `PtgExp(B2)` (self-reference), forcing shared-formula resolution to rely on
/// range containment rather than an exact master-cell key match. Correct behaviour is to resolve
/// `B2` against the `B1:B10` SHRFMLA (producing `C2+1`), not the `A1:A2` record.
pub fn build_shared_formula_shrfmla_range_cuse_ambiguity_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_shrfmla_range_cuse_ambiguity_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a shared formula range (`B1:B2`) where the shared
/// `SHRFMLA.rgce` contains a `PtgStr` token that is split across a `CONTINUE` boundary.
///
/// When a `PtgStr` payload is continued, Excel inserts a 1-byte continued-segment option flags
/// prefix at the start of the continued fragment. Consumers must skip that byte when reconstructing
/// the canonical `rgce` stream.
pub fn build_shared_formula_shrfmla_continued_ptgstr_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_shrfmla_continued_ptgstr_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a shared formula whose body references a workbook-global
/// defined name via `PtgName`.
///
/// This exercises the importer's BIFF8 shared-formula recovery path (`SHRFMLA` + `PtgExp`) while
/// ensuring formula decoding uses the workbook NAME table for `PtgName` indices.
pub fn build_shared_formula_ptgname_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgname_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a shared formula whose `SHRFMLA.rgce` includes a
/// `PtgStr` token split across a `CONTINUE` record boundary.
///
/// Shared formula range: `B1:B2`
/// - `B1`: `"ABCDE"` via `PtgExp` referencing itself
/// - `B2`: `"ABCDE"` via `PtgExp` referencing `B1`
///
/// The SHRFMLA record is split inside the `PtgStr` character payload, and the `CONTINUE` payload
/// begins with the required 1-byte continued-segment option flags prefix (`fHighByte`).
pub fn build_shared_formula_continued_ptgstr_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_continued_ptgstr_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a shared-formula definition (`SHRFMLA`) whose range
/// header uses the `Ref8` encoding (u16 column fields) instead of `RefU` (u8 columns).
///
/// The shared range starts at column A; if the importer incorrectly treats the header as `RefU`,
/// it will truncate the range to a single column and drop formulas in subsequent columns.
///
/// Shared formula range: `A1:B2`
/// - `A1:A2` and `B1:B2` formulas are recovered from the sheet-level `SHRFMLA` record.
pub fn build_shared_formula_shrfmla_ref8_header_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_shrfmla_ref8_header_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture where the `SHRFMLA` range header uses the `Ref8` encoding but omits
/// the `cUse` field.
///
/// Shared formula range: `A1:B2`
/// - `A1:A2` and `B1:B2` formulas are recovered from the sheet-level `SHRFMLA` record.
pub fn build_shared_formula_shrfmla_ref8_no_cuse_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_shrfmla_ref8_no_cuse_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture with a shared formula whose shared `SHRFMLA.rgce` includes a
/// `PtgMemAreaN` token.
///
/// Shared formula range: `B1:B2`
/// - `B1`: `A1+1`
/// - `B2`: `A2+1` (via `PtgExp`)
///
/// The shared formula `rgce` intentionally includes:
/// `PtgRefN` + `PtgMemAreaN(cce=0)` + `PtgMemAreaN(cce=3, rgce=PtgInt(0))` + `PtgInt(1)` + `PtgAdd`
///
/// `PtgMem*` tokens are no-ops for printing but carry a variable-length payload; if the shared
/// formula decoder mishandles them, subsequent tokens will be mis-parsed.
pub fn build_shared_formula_ptgmemarean_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgmemarean_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared-formula range (`SharedArea3D!B1:B2`) whose
/// shared `rgce` uses a 3D area reference (`PtgArea3d`) with relative flags.
///
/// This exercises shared-formula materialization for area endpoints with independent relative
/// bits. The intended decoded formulas are:
/// - `SharedArea3D!B1`: `Sheet1!A1:A2`
/// - `SharedArea3D!B2`: `Sheet1!A2:A3`
pub fn build_shared_formula_area3d_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_area3d_workbook_stream();
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

/// Build a minimal BIFF8 `.xls` fixture containing an array (CSE) formula.
///
/// The fixture contains a single sheet named `Array` with an array formula over `B1:B2` whose
/// formula is `A1:A2`. Both cells in the array range have `FORMULA` records whose `rgce` is
/// `PtgExp` referencing the base cell (`B1`), and the `ARRAY` record stores the shared `rgce`
/// token stream.
pub fn build_array_formula_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_array_formula_workbook_stream();

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

/// Build a minimal BIFF8 `.xls` fixture containing an array (CSE) formula whose array `rgce`
/// references a workbook-defined name via `PtgName`.
///
/// The fixture contains a single sheet named `ArrayName` with:
/// - A workbook-defined name `MyName` (NAME index 1)
/// - An array formula over `B1:B2` whose formula is `MyName+1`
///
/// Both cells in the array range have `FORMULA` records whose `rgce` is `PtgExp` referencing the
/// base cell (`B1`), and the `ARRAY` record stores the shared `rgce` token stream containing
/// `PtgName`.
pub fn build_array_formula_ptgname_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_array_formula_ptgname_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a What-If Analysis data table formula.
///
/// BIFF8 encodes legacy `TABLE(...)` formulas as:
/// - a `FORMULA.rgce` stream containing only `PtgTbl`, plus
/// - a worksheet-level `TABLE` record containing the input-cell references.
///
/// Calamine may omit these formulas because the `TABLE` record context lives outside the `rgce`
/// stream. This fixture is used to ensure our importer recovers a parseable `TABLE(A1,B2)` formula.
pub fn build_table_formula_fixture_xls() -> Vec<u8> {
    let xf_general = 16u16;
    let sheet_stream = build_table_formula_sheet_stream(xf_general);
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

/// Build a BIFF8 `.xls` fixture like [`build_table_formula_fixture_xls`], but where the `PtgTbl`
/// token in the result cell uses a non-canonical payload width (row u32 + col u16, 6 bytes).
///
/// Some `.xls` producers embed BIFF12-style coordinate widths in `PtgTbl`/`PtgExp` tokens. The `.xls`
/// importer should still recover a stable, parseable `TABLE(A1,B2)` formula string.
pub fn build_table_formula_ptgtbl_wide_payload_fixture_xls() -> Vec<u8> {
    let xf_general = 16u16;
    let sheet_stream = build_table_formula_ptgtbl_wide_payload_sheet_stream(xf_general);
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

/// Build a BIFF8 `.xls` fixture containing array formulas (`ARRAY` + `PtgExp`) whose array `rgce`
/// references an external workbook (`SUPBOOK`/`EXTERNSHEET`) and an external name (`PtgNameX`).
///
/// The fixture contains a single sheet named `ArrayExt` with two independent array-formula ranges:
/// - `B1:B2`: `'[Book1.xlsx]ExtSheet'!$A$1+1` via `PtgRef3d`
/// - `C1:C2`: `'[Book1.xlsx]ExtSheet'!ExtDefined+1` via `PtgNameX`
pub fn build_array_formula_external_refs_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_array_formula_external_refs_workbook_stream();
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

/// Build a minimal BIFF8 `.xls` fixture containing an array (CSE) formula whose array `rgce`
/// includes a `PtgArray` constant that must be decoded from trailing `rgcb` bytes in the `ARRAY`
/// record.
///
/// The fixture contains a single sheet named `ArrayConst` with an array formula over `B1:B2` whose
/// decoded formula includes `{1,2;3,4}`.
pub fn build_array_formula_ptgarray_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_array_formula_ptgarray_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture that exercises ambiguity in ARRAY range-header parsing when a
/// `RefU` header is followed by a small non-zero flags/reserved value.
///
/// Some decoders (including older versions of this crate) would mistakenly treat the first 8 bytes
/// of `RefU + flags` as a `Ref8` range header when the array range is `A..A`, causing the decoded
/// range to spuriously include additional columns.
///
/// This fixture constructs two ARRAY records:
/// - `A1:A2` (RefU + `flags=2`), array rgce is `"LEFT"`
/// - `B1:B10`, array rgce is `C1+1` (via `PtgRefN` decoded relative to base `B1`)
///
/// Cell `B2` stores `PtgExp(B2)` (self-reference), forcing array-formula resolution to rely on
/// range containment rather than an exact base-cell key match. Correct behaviour is to resolve
/// `B2` against the `B1:B10` ARRAY record (producing `C1+1`), not the `A1:A2` record.
pub fn build_array_formula_range_flags_ambiguity_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_array_formula_range_flags_ambiguity_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a malformed array-formula pattern:
/// - Base cell contains a full `FORMULA.rgce` token stream (no `PtgExp`).
/// - Follower cell contains only `PtgExp` pointing at the base cell and sets `FORMULA.grbit.fArray`.
/// - The expected `ARRAY` record is intentionally missing.
///
/// The `.xls` importer should still recover the follower formula, and array membership should
/// preserve the base-cell coordinate space (no shared-formula-style reference shifting).
pub fn build_array_formula_ptgexp_missing_array_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_array_formula_ptgexp_missing_array_workbook_stream();
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

fn build_table_formula_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: keep large enough to cover our TABLE base cell and formula cell.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&25u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&10u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    // WINDOW2 is required by some consumers; keep defaults.
    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide an input cell value (A1) so calamine reports a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // TABLE record anchored at F11 (row=10, col=5), with two input cells: A1 (row input) and
    // B2 (col input).
    let base_row: u16 = 10;
    let base_col: u16 = 5;
    let grbit: u16 = 0x0003; // best-effort: both inputs present

    let mut table_payload = Vec::new();
    table_payload.extend_from_slice(&base_row.to_le_bytes());
    table_payload.extend_from_slice(&base_col.to_le_bytes());
    table_payload.extend_from_slice(&grbit.to_le_bytes());
    // row input: A1
    table_payload.extend_from_slice(&0u16.to_le_bytes()); // rwInpRow
    table_payload.extend_from_slice(&0u16.to_le_bytes()); // colInpRow
                                                          // col input: B2
    table_payload.extend_from_slice(&1u16.to_le_bytes()); // rwInpCol
    table_payload.extend_from_slice(&1u16.to_le_bytes()); // colInpCol
    push_record(&mut sheet, RECORD_TABLE, &table_payload);

    // FORMULA record at D21 whose rgce is a single PtgTbl referencing the TABLE base cell.
    let cell_row: u16 = 20;
    let cell_col: u16 = 3;
    let rgce = [
        0x02u8, // PtgTbl
        base_row.to_le_bytes()[0],
        base_row.to_le_bytes()[1],
        base_col.to_le_bytes()[0],
        base_col.to_le_bytes()[1],
    ];

    let mut formula_payload = Vec::new();
    formula_payload.extend_from_slice(&cell_row.to_le_bytes());
    formula_payload.extend_from_slice(&cell_col.to_le_bytes());
    formula_payload.extend_from_slice(&xf_cell.to_le_bytes()); // xf
    formula_payload.extend_from_slice(&0f64.to_le_bytes()); // cached result
    formula_payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
    formula_payload.extend_from_slice(&0u32.to_le_bytes()); // chn
    formula_payload.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
    formula_payload.extend_from_slice(&rgce);
    push_record(&mut sheet, RECORD_FORMULA, &formula_payload);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_table_formula_ptgtbl_wide_payload_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: keep large enough to cover our TABLE base cell and formula cell.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&25u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&10u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    // WINDOW2 is required by some consumers; keep defaults.
    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide an input cell value (A1) so calamine reports a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // TABLE record anchored at F11 (row=10, col=5), with two input cells: A1 (row input) and
    // B2 (col input).
    let base_row: u16 = 10;
    let base_col: u16 = 5;
    let grbit: u16 = 0x0003; // best-effort: both inputs present

    let mut table_payload = Vec::new();
    table_payload.extend_from_slice(&base_row.to_le_bytes());
    table_payload.extend_from_slice(&base_col.to_le_bytes());
    table_payload.extend_from_slice(&grbit.to_le_bytes());
    // row input: A1
    table_payload.extend_from_slice(&0u16.to_le_bytes()); // rwInpRow
    table_payload.extend_from_slice(&0u16.to_le_bytes()); // colInpRow
                                                          // col input: B2
    table_payload.extend_from_slice(&1u16.to_le_bytes()); // rwInpCol
    table_payload.extend_from_slice(&1u16.to_le_bytes()); // colInpCol
    push_record(&mut sheet, RECORD_TABLE, &table_payload);

    // FORMULA record at D21 whose rgce is a single PtgTbl with a *wide* payload:
    //   [ptg:0x02][rw:u32][col:u16]
    let cell_row: u16 = 20;
    let cell_col: u16 = 3;
    let mut rgce: Vec<u8> = Vec::new();
    rgce.push(0x02); // PtgTbl
    rgce.extend_from_slice(&(base_row as u32).to_le_bytes());
    rgce.extend_from_slice(&base_col.to_le_bytes());

    let mut formula_payload = Vec::new();
    formula_payload.extend_from_slice(&cell_row.to_le_bytes());
    formula_payload.extend_from_slice(&cell_col.to_le_bytes());
    formula_payload.extend_from_slice(&xf_cell.to_le_bytes()); // xf
    formula_payload.extend_from_slice(&0f64.to_le_bytes()); // cached result
    formula_payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
    formula_payload.extend_from_slice(&0u32.to_le_bytes()); // chn
    formula_payload.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
    formula_payload.extend_from_slice(&rgce);
    push_record(&mut sheet, RECORD_FORMULA, &formula_payload);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

/// Build a BIFF8 `.xls` fixture like [`build_shared_formula_area3d_fixture_xls`], but using a
/// `PtgArea3d` token whose endpoints have *different* relative flags.
///
/// Intended decoded formulas:
/// - `SharedArea3D!B1`: `Sheet1!A1:$A2`
/// - `SharedArea3D!B2`: `Sheet1!A2:$A3`
///
/// This catches bugs where materialization/decoding incorrectly reuses the first endpoint’s
/// relative flags for the second endpoint (or drops the flags entirely).
pub fn build_shared_formula_area3d_mixed_flags_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_area3d_mixed_flags_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing shared formulas (`SHRFMLA` + `PtgExp`) whose shared
/// `rgce` references an external workbook (`SUPBOOK`/`EXTERNSHEET`) and an external name
/// (`PtgNameX`).
///
/// This is used to validate that our BIFF shared-formula decode path preserves external reference
/// fidelity, matching non-shared formulas.
pub fn build_shared_formula_external_refs_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_external_refs_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a **2D** shared-formula range (`B1:C2`).
///
/// The shared formula base (`SHRFMLA.rgce`) uses `PtgRefN(col_off=-1) + 1 + +`, so each materialized
/// cell formula depends on decoding the shared `rgce` with the *correct base cell* (row and column).
pub fn build_shared_formula_2d_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_2d_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula over `B1:B2` where the shared formula
/// definition (`SHRFMLA`) includes a `PtgArray` constant stored in trailing `rgcb` bytes.
///
/// This is used to validate that BIFF8 shared-formula expansion (via `PtgExp`) preserves `rgcb`
/// blocks so array constants decode to `{...}` literals rather than `#UNKNOWN!`.
pub fn build_shared_formula_ptgarray_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgarray_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula over `B1:B2` where:
/// - the shared formula definition (`SHRFMLA`) includes a `PtgArray` constant stored in trailing
///   `rgcb` bytes, and
/// - the follower cell (`B2`) uses a **wide** non-standard `PtgExp` payload layout (row u32 + col
///   u16, 6 bytes; `cce=7`).
///
/// This exercises the `.xls` importer's wide-payload `PtgExp` recovery path and ensures it
/// preserves and uses SHRFMLA trailing `rgcb` so array constants decode to `{...}` literals rather
/// than `#UNKNOWN!`.
pub fn build_shared_formula_ptgarray_wide_ptgexp_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgarray_wide_ptgexp_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula where a follower cell uses a
/// non-standard `PtgExp` payload layout (row u32 + col u16, 6 bytes).
///
/// Some producers emit BIFF12-style coordinate widths even in BIFF8 `.xls` files. The `.xls`
/// importer should still resolve the follower formula against the sheet’s SHRFMLA definition.
pub fn build_shared_formula_ptgexp_u32_row_u16_col_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgexp_u32_row_u16_col_workbook_stream();

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

/// Build a minimal BIFF8 `.xls` fixture containing a shared formula (`SHRFMLA` + `PtgExp`).
///
/// The fixture contains a single sheet (`Sheet1`) with:
/// - `Sheet1!B1` formula: `A1+1`
/// - `Sheet1!B2` FORMULA record containing only `PtgExp` referencing `B1`
/// - a `SHRFMLA` record covering `B1:B2` with a shared token stream using `PtgRefN` so the
///   per-cell formula decodes as `A1+1` / `A2+1`.
pub fn build_shared_formula_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a worksheet formula that uses an array constant
/// (`PtgArray` + trailing `rgcb`).
///
/// This exercises the importer's ability to decode BIFF8 array constants into parseable Excel
/// syntax like `{1,2;3,4}` instead of degrading to `#UNKNOWN!`.
pub fn build_formula_array_constant_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_formula_array_constant_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula range at the BIFF8 row limit.
///
/// Shared formulas are stored once (in `SHRFMLA`) and referenced via `PtgExp` tokens in follower
/// cells. When the shared token stream contains relative references (`PtgRefN` / `PtgAreaN`) that
/// become out-of-bounds after shifting, Excel renders them as `#REF!` in the affected cells.
///
/// This fixture encodes a shared formula range `B65535:B65536` where the second cell's shifted
/// reference points beyond BIFF8's max row (65536 in 1-based A1 terms), so the follower should
/// materialize to `#REF!+1` rather than being dropped/unresolved.
pub fn build_shared_formula_out_of_bounds_relative_refs_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_out_of_bounds_relative_refs_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula where the follower cell uses a wide
/// `PtgExp` payload (row u32 + col u16) **and** the shared `SHRFMLA.rgce` uses `PtgRef` with
/// row/col-relative flags (rather than `PtgRefN` offsets).
///
/// Expected decoded formula in B3: `A3+1`.
pub fn build_shared_formula_ptgexp_wide_payload_ptgref_relative_flags_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_shared_formula_ptgexp_wide_payload_ptgref_relative_flags_workbook_stream();
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

/// Build a BIFF8 `.xls` fixture containing a worksheet formula with an array constant (`PtgArray`)
/// whose `rgcb` payload is split across a `CONTINUE` record boundary inside the UTF-16 string bytes.
///
/// BIFF8 inserts a 1-byte continued-segment option flags prefix (`fHighByte`) at the start of the
/// continued fragment when a string payload crosses a `CONTINUE` boundary. The importer should skip
/// that byte so the decoded array literal remains parseable (e.g. `{\"ABCDE\"}`).
pub fn build_formula_array_constant_continued_rgcb_string_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_formula_array_constant_continued_rgcb_string_workbook_stream();

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

/// Like [`build_formula_array_constant_continued_rgcb_string_fixture_xls`], but the continued string
/// segment is stored in the *compressed* (single-byte) form (`fHighByte=0`).
pub fn build_formula_array_constant_continued_rgcb_string_compressed_fixture_xls() -> Vec<u8> {
    let workbook_stream =
        build_formula_array_constant_continued_rgcb_string_compressed_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture that references out-of-range XF indices from ROW/COLINFO `ixfe`
/// fields.
///
/// The importer should ignore the invalid indices but emit a warning (without panicking).
pub fn build_row_col_style_out_of_range_xf_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_row_col_style_out_of_range_xf_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture that combines:
/// - BIFF8 RC4 CryptoAPI `FILEPASS` encryption, and
/// - a defined name (`NAME` record) split across `CONTINUE` records.
///
/// This is used to ensure that after decryption, the workbook stream is still routed through the
/// continued-NAME sanitizer before opening via `calamine` (preventing panics while building the
/// `NAME` table needed for `PtgName` tokens).
pub fn build_encrypted_continued_name_record_fixture_xls_rc4_cryptoapi(password: &str) -> Vec<u8> {
    let workbook_stream = build_continued_name_record_workbook_stream();
    let encrypted_workbook_stream =
        encrypt_biff8_workbook_stream_rc4_cryptoapi(&workbook_stream, password);

    let cursor = Cursor::new(Vec::new());
    let mut ole =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&encrypted_workbook_stream)
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

/// Build a BIFF8 `.xls` fixture that contains a shared formula range that uses a 3D reference
/// token (`PtgRef3d`) with relative row/column flags.
///
/// The fixture contains two sheets:
/// - `Sheet1`: contains a value in A1 (not required for formula text import, but keeps the sheet
///   non-empty for calamine).
/// - `Shared3D`: contains a shared formula over `B1:C2`.
///
/// The shared formula definition is stored in a `SHRFMLA` record with base `rgce`:
///   `Sheet1!A1 + 1`
/// where the `PtgRef3d` token sets both row and column *relative* flags so the reference is
/// materialized across the shared range.
///
/// All cells in the range have `FORMULA` records; non-base cells use `PtgExp` to reference the
/// base cell.
pub fn build_shared_formula_ref3d_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ref3d_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula near the BIFF8 row limit where
/// materialization of a 3D reference (`PtgRef3d`) shifts out of bounds and must become `#REF!`.
///
/// Sheets:
/// - `Sheet1`: contains a NUMBER cell at `A65536` (row index 65535).
/// - `Shared3D_OOB`: shared formula over `B65535:B65536`.
///   - Shared rgce (SHRFMLA): `Sheet1!A65536+1`, where the `PtgRef3d` token sets both row/col
///     relative flags so the follower shifts by +1 row.
///   - Both cells store `PtgExp` pointing at the base cell so calamine does not need to decode the
///     out-of-range shifted reference itself; the importer recovers the materialized formulas from
///     BIFF.
///
/// Expected decoded formulas:
/// - `Shared3D_OOB!B65535` = `Sheet1!A65536+1`
/// - `Shared3D_OOB!B65536` = `#REF!+1`
pub fn build_shared_formula_3d_oob_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_3d_oob_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_shared_formula_3d_oob_fixture_xls`], but omitting any
/// cell-level `FORMULA` records in the shared range.
///
/// The only source of the shared formula is the worksheet-level `SHRFMLA` record, so the importer
/// must populate per-cell formulas by expanding/materializing the shared rgce across the range.
pub fn build_shared_formula_3d_oob_shrfmla_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_3d_oob_shrfmla_only_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula at the BIFF8 max column where a 3D
/// reference (`PtgRef3d`) shifts out of bounds **horizontally** and must become `#REF!`.
///
/// This variant stores the shared formula definition only in a `SHRFMLA` record (no `FORMULA`
/// records in the shared range), so the importer must expand/materialize formulas itself.
///
/// Sheets:
/// - `Sheet1`: a minimal value sheet (target of the 3D reference).
/// - `Shared3D_ColOOB_ShrFmlaOnly`: shared range `XFC1:XFD1` (0x3FFE..=0x3FFF).
///
/// Expected decoded formulas:
/// - `Shared3D_ColOOB_ShrFmlaOnly!XFC1` = `Sheet1!XFD1+1`
/// - `Shared3D_ColOOB_ShrFmlaOnly!XFD1` = `#REF!+1`
pub fn build_shared_formula_3d_col_oob_shrfmla_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_3d_col_oob_shrfmla_only_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula near the BIFF8 row limit where
/// materialization of a 3D **area** reference (`PtgArea3d`) shifts out of bounds and must become
/// `#REF!`.
///
/// Sheets:
/// - `Sheet1`: contains NUMBER cells at `A65535` and `A65536` (row indices 65534 and 65535).
/// - `SharedArea3D_OOB`: shared formula over `B65535:B65536`.
///   - Shared rgce (SHRFMLA): `Sheet1!A65535:A65536+1`, where the `PtgArea3d` token sets row/col
///     relative flags on both endpoints so the follower shifts by +1 row.
///   - Both cells store `PtgExp` pointing at the base cell so calamine does not need to decode the
///     out-of-range shifted area itself; the importer recovers the materialized formulas from BIFF.
///
/// Expected decoded formulas:
/// - `SharedArea3D_OOB!B65535` = `Sheet1!A65535:A65536+1`
/// - `SharedArea3D_OOB!B65536` = `#REF!+1`
pub fn build_shared_formula_area3d_oob_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_area3d_oob_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_shared_formula_area3d_oob_fixture_xls`], but omitting
/// any cell-level `FORMULA` records in the shared range (SHRFMLA-only).
///
/// The importer must expand/materialize the shared rgce across the range.
pub fn build_shared_formula_area3d_oob_shrfmla_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_area3d_oob_shrfmla_only_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a shared formula at the BIFF8 max column where a 3D
/// **area** reference (`PtgArea3d`) shifts out of bounds horizontally and must become `#REF!`.
///
/// This variant stores the shared formula definition only in a `SHRFMLA` record (no `FORMULA`
/// records in the shared range), so the importer must expand/materialize formulas itself.
///
/// Sheets:
/// - `Sheet1`: a minimal value sheet (target of the 3D reference).
/// - `SharedArea3D_ColOOB_ShrFmlaOnly`: shared range `XFC1:XFD1`.
///
/// Expected decoded formulas:
/// - `SharedArea3D_ColOOB_ShrFmlaOnly!XFC1` = `Sheet1!XFC1:XFD1+1`
/// - `SharedArea3D_ColOOB_ShrFmlaOnly!XFD1` = `#REF!+1`
pub fn build_shared_formula_area3d_col_oob_shrfmla_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_area3d_col_oob_shrfmla_only_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing an invalid FitTo page setup:
/// - `WSBOOL.fFitToPage=1`
/// - `SETUP.iFitWidth=40000`, `SETUP.iFitHeight=40000` (out of spec; should be clamped by importer)
pub fn build_fit_to_clamp_fixture_xls() -> Vec<u8> {
    let xf_cell: u16 = 16; // First cell XF after the 16 style XFs in `build_single_sheet_workbook_stream`.
    let sheet_stream = build_fit_to_clamp_sheet_stream(xf_cell);
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

/// Build a BIFF8 `.xls` fixture containing a single worksheet with malformed/truncated page setup
/// records.
///
/// This exercises the importer's best-effort BIFF page setup parsing:
/// - truncated `SETUP` record payload (<34 bytes)
/// - truncated `LEFTMARGIN` record payload (<8 bytes)
/// - truncated `HORIZONTALPAGEBREAKS` record where `cbrk` claims more entries than present
pub fn build_page_setup_malformed_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_page_setup_malformed_workbook_stream();

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

/// Build a minimal BIFF8 `.xls` fixture containing a single sheet named `Sheet1` with worksheet
/// page setup records (WSBOOL + SETUP + margins) and manual page breaks (horizontal + vertical).
///
/// The sheet uses percent-based scaling (`WSBOOL.fFitToPage = 0`) and a non-default `SETUP.iScale`.
pub fn build_page_setup_and_breaks_fixture_xls_percent_scaling() -> Vec<u8> {
    build_page_setup_and_breaks_fixture_xls(PageSetupScalingMode::Percent)
}

/// Build a minimal BIFF8 `.xls` fixture containing a single sheet named `Sheet1` with worksheet
/// page setup records (WSBOOL + SETUP + margins) and manual page breaks (horizontal + vertical).
///
/// The sheet uses fit-to-page scaling (`WSBOOL.fFitToPage = 1`) and non-default `SETUP.iFitWidth`
/// + `SETUP.iFitHeight` values.
pub fn build_page_setup_and_breaks_fixture_xls_fit_to_scaling() -> Vec<u8> {
    build_page_setup_and_breaks_fixture_xls(PageSetupScalingMode::FitTo)
}

fn build_page_setup_and_breaks_fixture_xls(mode: PageSetupScalingMode) -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_page_setup_fixture_sheet_stream(16, mode);
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

/// Build a minimal BIFF8 `.xls` fixture containing a single worksheet with page margin records but
/// **no** `SETUP` record.
///
/// This validates that legacy margin records alone (`LEFTMARGIN`/`RIGHTMARGIN`/`TOPMARGIN`/
/// `BOTTOMMARGIN`) are sufficient to populate non-default [`formula_model::PageMargins`].
pub fn build_margins_without_setup_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_margins_without_setup_sheet_stream(16, 1.25, 1.5, 0.5, 2.25);
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

/// Build a BIFF8 `.xls` fixture containing a worksheet `SETUP` record that is intentionally
/// truncated (payload < 34 bytes), but still contains enough bytes to recover the paper size,
/// orientation, scaling, and header margin (`numHdr`).
pub fn build_truncated_setup_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_truncated_setup_sheet_stream(16);
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

/// Build a BIFF8 `.xls` fixture containing a worksheet `WSBOOL` record that is intentionally
/// truncated (payload < 2 bytes).
pub fn build_truncated_wsbool_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_truncated_wsbool_sheet_stream(16);
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

/// Build a BIFF8 `.xls` fixture containing a truncated worksheet margin record followed by a valid
/// record of the same type (to validate "last valid wins" semantics).
pub fn build_truncated_margin_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_truncated_margin_sheet_stream(16);
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
/// using percent scaling (`WSBOOL.fFitToPage=0`, `SETUP.iScale=85`).
pub fn build_page_setup_percent_scaling_fixture_xls() -> Vec<u8> {
    build_page_setup_and_breaks_fixture_xls_percent_scaling()
}

/// Build a BIFF8 `.xls` fixture containing worksheet page setup + margins + manual page breaks,
/// using fit-to scaling (`WSBOOL.fFitToPage=1`, `SETUP.iFitWidth=2`, `SETUP.iFitHeight=3`).
pub fn build_page_setup_fit_to_scaling_fixture_xls() -> Vec<u8> {
    build_page_setup_and_breaks_fixture_xls_fit_to_scaling()
}

/// Build a BIFF8 `.xls` fixture containing a worksheet `SETUP` record with a custom paper size
/// (`iPaperSize=0`, `fNoPls=0`).
pub fn build_custom_paper_size_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_custom_paper_size_sheet_stream(16, 0);
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
/// (`iPaperSize>=256`, `fNoPls=0`).
pub fn build_custom_paper_size_ge_256_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_custom_paper_size_sheet_stream(16, 256);
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

/// Build a BIFF8 `.xls` fixture containing a single sheet with a `SETUP` record where
/// `SETUP.fNoPls=1`.
///
/// This is used to validate that the importer ignores paper size/orientation/scaling fields while
/// still importing header/footer margins (`numHdr`/`numFtr`).
pub fn build_page_setup_flags_nopls_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_page_setup_flags_sheet_stream(
        16,
        setup_record_with_grbit(
            9,                   // iPaperSize (ignored due to fNoPls)
            80,                  // iScale (ignored due to fNoPls)
            0,                   // iFitWidth
            0,                   // iFitHeight
            SETUP_GRBIT_F_NOPLS, // fNoPls=1, fPortrait=0 (landscape)
            0.5,                 // numHdr (non-default)
            0.6,                 // numFtr (non-default)
        ),
    );
    let workbook_stream = build_single_sheet_workbook_stream("PageSetupNoPls", &sheet_stream, 1252);

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

/// Build a BIFF8 `.xls` fixture containing a single sheet with a `SETUP` record where
/// `SETUP.fNoPls=1` and a `WSBOOL` record where `WSBOOL.fFitToPage=1`.
///
/// This exercises the nuance that `fNoPls` only makes *printer-related* SETUP fields undefined
/// (paper size/orientation/percent scaling), while `iFitWidth`/`iFitHeight` can still be honored
/// when fit-to-page mode is enabled.
pub fn build_page_setup_flags_nopls_fit_to_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_page_setup_flags_sheet_stream_with_wsbool(
        16,
        setup_record_with_grbit(
            9,                   // iPaperSize (ignored due to fNoPls)
            80,                  // iScale (ignored due to fNoPls)
            2,                   // iFitWidth (preserved)
            3,                   // iFitHeight (preserved)
            SETUP_GRBIT_F_NOPLS, // fNoPls=1, fPortrait=0 (landscape)
            0.5,                 // numHdr (non-default)
            0.6,                 // numFtr (non-default)
        ),
        WSBOOL_OPTION_FIT_TO_PAGE,
    );
    let workbook_stream =
        build_single_sheet_workbook_stream("PageSetupNoPlsFitTo", &sheet_stream, 1252);

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

/// Build a BIFF8 `.xls` fixture containing a single sheet with a `SETUP` record where
/// `SETUP.fNoOrient=1` (and `SETUP.fNoPls=0`).
///
/// This is used to validate that the importer treats orientation as the default portrait and
/// ignores the `fPortrait` bit when `fNoOrient` is set.
pub fn build_page_setup_flags_noorient_fixture_xls() -> Vec<u8> {
    // `build_single_sheet_workbook_stream` always emits a single cell XF at index 16 (after 16
    // style XFs).
    let sheet_stream = build_page_setup_flags_sheet_stream(
        16,
        setup_record_with_grbit(
            9,                      // A4
            80,                     // iScale=80%
            0,                      // iFitWidth
            0,                      // iFitHeight
            SETUP_GRBIT_F_NOORIENT, // fNoOrient=1, fPortrait=0 (landscape)
            0.5,                    // numHdr (non-default)
            0.6,                    // numFtr (non-default)
        ),
    );
    let workbook_stream =
        build_single_sheet_workbook_stream("PageSetupNoOrient", &sheet_stream, 1252);

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
            // Note: keep non-zero iFitWidth/iFitHeight even in percent mode so the importer must
            // consult WSBOOL.fFitToPage to decide which scaling fields apply.
            &setup_record(9, 85, 2, 3, true, 0.9, 1.0),
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
    //
    // BIFF8 page breaks store the 0-based index of the first row/col *after* the break, while the
    // model stores the 0-based row/col index after which the break occurs. The record helpers below
    // take the model form (`breaks_after`) and encode BIFF values by adding 1.
    push_record(&mut sheet, RECORD_HPAGEBREAKS, &hpagebreaks_record(&[4]));
    push_record(&mut sheet, RECORD_VPAGEBREAKS, &vpagebreaks_record(&[2]));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 0.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_truncated_setup_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // WSBOOL with fFitToPage cleared so scaling comes from SETUP.iScale.
    let wsbool: u16 = 0x0C01;
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool.to_le_bytes());

    // SETUP record truncated to 24 bytes (through numHdr).
    let mut setup = setup_record(9, 80, 2, 3, true, 0.55, 0.66);
    setup.truncate(24);
    push_record(&mut sheet, RECORD_SETUP, &setup);

    // A1: a single cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_truncated_wsbool_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // WSBOOL record with a truncated (1-byte) payload.
    push_record(&mut sheet, RECORD_WSBOOL, &[0x00]);

    // Full-length SETUP record with non-default iScale and zero fit dimensions so scaling falls
    // back to percent mode when WSBOOL is ignored.
    push_record(
        &mut sheet,
        RECORD_SETUP,
        &setup_record(9, 80, 0, 0, true, 0.3, 0.3),
    );

    // A1: a single cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_truncated_margin_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Truncated TOPMARGIN payload (expected 8 bytes / f64).
    push_record(&mut sheet, RECORD_TOPMARGIN, &[0x00, 0x01, 0x02, 0x03]);
    // Followed by a valid TOPMARGIN record; the importer should use this value.
    push_record(&mut sheet, RECORD_TOPMARGIN, &2.25f64.to_le_bytes());

    // A1: a single cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

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

/// Build a BIFF8 `.xls` fixture containing two sheets where sanitizing one invalid sheet name
/// collides with another sheet's original name.
///
/// Sheet 0: `Bad:Name` (invalid, sanitizes to `Bad_Name`). Includes distinctive page setup and
/// manual breaks (same payload as [`build_page_setup_sanitized_sheet_name_fixture_xls`]).
///
/// Sheet 1: `Bad_Name` (valid, collides with sheet 0's sanitized name and is deduped to
/// `Bad_Name (2)`). Includes *different* page setup and manual breaks.
///
/// This is used to ensure BIFF-derived print settings are applied by sheet identity (or final
/// sanitized name) rather than the original BIFF BoundSheet name. A buggy implementation that
/// applies print settings by the BIFF name would mis-attach sheet 1's settings to sheet 0.
pub fn build_page_setup_sanitized_sheet_name_collision_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_page_setup_sanitized_sheet_name_collision_workbook_stream();

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
    //
    // Note: BIFF8 page breaks store the 0-based index of the first row/col *after* the break.
    // Our fixture helpers accept the model’s “after which break occurs” form (0-based) and encode
    // them into BIFF8’s representation by adding 1.
    //
    // Break after row 1 and row 4, and after column 2 (i.e. between C and D).
    push_record(&mut sheet, RECORD_HPAGEBREAKS, &hpagebreaks_record(&[1, 4]));
    push_record(&mut sheet, RECORD_VPAGEBREAKS, &vpagebreaks_record(&[2]));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_page_setup_sanitized_sheet_name_collision_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One default cell XF (General).
    let xf_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Sheet 0: invalid BIFF name (contains ':') that sanitizes to `Bad_Name`.
    let boundsheet0_start = globals.len();
    let mut boundsheet0 = Vec::<u8>::new();
    boundsheet0.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet0.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet0, "Bad:Name");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet0);
    let boundsheet0_offset_pos = boundsheet0_start + 4;

    // Sheet 1: already has the sanitized base name, will be deduped to `Bad_Name (2)`.
    let boundsheet1_start = globals.len();
    let mut boundsheet1 = Vec::<u8>::new();
    boundsheet1.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet1.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet1, "Bad_Name");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet1);
    let boundsheet1_offset_pos = boundsheet1_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet substreams ------------------------------------------------------
    let sheet0_offset = globals.len();
    let sheet0 = build_page_setup_sanitized_sheet_name_sheet_stream(xf_general);
    globals[boundsheet0_offset_pos..boundsheet0_offset_pos + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet0);

    let sheet1_offset = globals.len();
    let sheet1 = build_page_setup_sheet_stream(
        xf_general,
        PageSetupFixtureSheet {
            paper_size: 1,
            landscape: false,
            scale_percent: 77,
            header_margin: 0.11,
            footer_margin: 0.22,
            left_margin: 5.55,
            right_margin: 6.66,
            top_margin: 7.77,
            bottom_margin: 8.88,
            row_break_after: 2,
            col_break_after: 1,
            cell_value: 2.0,
        },
    );
    globals[boundsheet1_offset_pos..boundsheet1_offset_pos + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet1);

    globals
}

/// Build a BIFF8 `.xls` fixture with worksheet page setup records designed to exercise print
/// settings record-order merging:
///
/// - LEFT/RIGHT/TOP/BOTTOMMARGIN records appear **before** SETUP.
/// - The margin records set left/right/top/bottom to non-default values.
/// - SETUP sets header/footer margins plus paper/orientation/scaling fields.
pub fn build_page_setup_margins_before_setup_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_page_setup_margins_before_setup_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture that sets WSBOOL.fFitToPage=1 and includes a SETUP record with both
/// iScale and iFitWidth/iFitHeight populated (non-default).
pub fn build_page_setup_scaling_fit_to_page_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_page_setup_scaling_fit_to_page_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture that sets WSBOOL.fFitToPage=0 and includes a SETUP record with both
/// iScale and iFitWidth/iFitHeight populated (non-default).
pub fn build_page_setup_scaling_percent_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_page_setup_scaling_percent_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture that stores `SETUP` **before** `WSBOOL` and sets `WSBOOL.fFitToPage=1`.
///
/// This exists to validate that print settings import is resilient to record-order variations:
/// `WSBOOL` can appear after `SETUP` and should still control whether iFitWidth/iFitHeight apply.
pub fn build_page_setup_scaling_fit_to_page_wsbool_after_setup_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_page_setup_scaling_fit_to_page_wsbool_after_setup_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture that includes two `WSBOOL` records:
/// - the first enables fit-to-page
/// - the last disables it
///
/// This validates "last record wins" semantics for `WSBOOL.fFitToPage`.
pub fn build_page_setup_scaling_wsbool_last_wins_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_page_setup_scaling_wsbool_last_wins_workbook_stream();

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

    // WSBOOL controls scaling mode (`fFitToPage`). Enable fit-to-page so
    // `SETUP.iFitWidth/iFitHeight` are used instead of percent scaling.
    let wsbool: u16 = 0x0C01 | WSBOOL_OPTION_FIT_TO_PAGE;
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool.to_le_bytes());

    // Margins.
    push_record(&mut sheet, RECORD_LEFTMARGIN, &1.1f64.to_le_bytes());
    push_record(&mut sheet, RECORD_RIGHTMARGIN, &1.2f64.to_le_bytes());
    push_record(&mut sheet, RECORD_TOPMARGIN, &1.3f64.to_le_bytes());
    push_record(&mut sheet, RECORD_BOTTOMMARGIN, &1.4f64.to_le_bytes());

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
    // WSBOOL.fFitToPage (bit 8) controls whether SETUP.iScale or SETUP.iFit* fields apply.
    push_record(&mut sheet, RECORD_WSBOOL, &0x0100u16.to_le_bytes()); // fFitToPage=1

    // WSBOOL: Excel stores `Fit to page` as a worksheet boolean flag (fFitToPage = bit 0x0100).
    // Use the same base value (0x0C01) as other fixtures so outline defaults remain consistent.
    push_record(&mut sheet, RECORD_WSBOOL, &0x0D01u16.to_le_bytes());

    // Manual page breaks.
    //
    // BIFF8 page breaks store the 0-based index of the first row/col *after* the break, while the
    // model stores the 0-based row/col index after which the break occurs. The record helpers below
    // take the model form (`breaks_after`) and encode BIFF values by adding 1.
    push_record(&mut sheet, RECORD_HPAGEBREAKS, &hpagebreaks_record(&[2, 4]));
    push_record(&mut sheet, RECORD_VPAGEBREAKS, &vpagebreaks_record(&[1]));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_custom_paper_size_sheet_stream(xf_cell: u16, paper_size: u16) -> Vec<u8> {
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
            paper_size, // iPaperSize (custom/unmappable)
            100,        // iScale
            0,          // iFitWidth
            0,          // iFitHeight
            false,      // portrait (keep model default)
            0.3,        // header margin (default)
            0.3,        // footer margin (default)
        ),
    );

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_page_setup_margins_before_setup_workbook_stream() -> Vec<u8> {
    let xf_general = 16u16;
    let sheet_stream = build_page_setup_margins_before_setup_sheet_stream(xf_general);
    build_single_sheet_workbook_stream("PageSetup", &sheet_stream, 1252)
}

fn build_page_setup_scaling_fit_to_page_workbook_stream() -> Vec<u8> {
    let xf_general = 16u16;
    let sheet_stream = build_page_setup_scaling_sheet_stream(xf_general, true);
    build_single_sheet_workbook_stream("ScaleFitTo", &sheet_stream, 1252)
}

fn build_page_setup_scaling_percent_workbook_stream() -> Vec<u8> {
    let xf_general = 16u16;
    let sheet_stream = build_page_setup_scaling_sheet_stream(xf_general, false);
    build_single_sheet_workbook_stream("ScalePercent", &sheet_stream, 1252)
}

fn build_page_setup_scaling_fit_to_page_wsbool_after_setup_workbook_stream() -> Vec<u8> {
    let xf_general = 16u16;
    let sheet_stream = build_page_setup_scaling_sheet_stream_wsbool_after_setup(xf_general, true);
    build_single_sheet_workbook_stream("ScaleFitToAfterSetup", &sheet_stream, 1252)
}

fn build_page_setup_scaling_wsbool_last_wins_workbook_stream() -> Vec<u8> {
    let xf_general = 16u16;
    let sheet_stream = build_page_setup_scaling_sheet_stream_wsbool_last_wins(xf_general);
    build_single_sheet_workbook_stream("ScaleWsboolLastWins", &sheet_stream, 1252)
}

fn build_page_setup_margins_before_setup_sheet_stream(xf_general: u16) -> Vec<u8> {
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

    // Margin records first (record-order precedence test).
    push_record(&mut sheet, RECORD_LEFTMARGIN, &1.25f64.to_le_bytes());
    push_record(&mut sheet, RECORD_RIGHTMARGIN, &1.5f64.to_le_bytes());
    push_record(&mut sheet, RECORD_TOPMARGIN, &2.25f64.to_le_bytes());
    push_record(&mut sheet, RECORD_BOTTOMMARGIN, &2.5f64.to_le_bytes());

    // SETUP record (header/footer margins + paper/orientation/scaling fields).
    push_record(
        &mut sheet,
        RECORD_SETUP,
        &setup_record(
            9,    // A4
            88,   // non-default percent
            2,    // fit width
            3,    // fit height
            true, // landscape
            0.25, // header margin
            0.5,  // footer margin
        ),
    );

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 0.0),
    );
    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_page_setup_scaling_sheet_stream(xf_general: u16, fit_to_page: bool) -> Vec<u8> {
    const WSBOOL_DEFAULT: u16 = 0x0C01;

    let wsbool = if fit_to_page {
        WSBOOL_DEFAULT | WSBOOL_OPTION_FIT_TO_PAGE
    } else {
        WSBOOL_DEFAULT
    };

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

    // WSBOOL controls whether `SETUP.iFitWidth/iFitHeight` are honored.
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool.to_le_bytes());

    push_record(
        &mut sheet,
        RECORD_SETUP,
        &setup_record(
            1,     // Letter
            77,    // non-default percent
            2,     // fit width
            3,     // fit height
            false, // portrait
            0.3,   // header margin
            0.3,   // footer margin
        ),
    );

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 0.0),
    );
    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_page_setup_scaling_sheet_stream_wsbool_after_setup(
    xf_general: u16,
    fit_to_page: bool,
) -> Vec<u8> {
    const WSBOOL_DEFAULT: u16 = 0x0C01;

    let wsbool = if fit_to_page {
        WSBOOL_DEFAULT | WSBOOL_OPTION_FIT_TO_PAGE
    } else {
        WSBOOL_DEFAULT
    };

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

    // SETUP first.
    push_record(
        &mut sheet,
        RECORD_SETUP,
        &setup_record(
            1,     // Letter
            77,    // non-default percent
            2,     // fit width
            3,     // fit height
            false, // portrait
            0.3,   // header margin
            0.3,   // footer margin
        ),
    );

    // WSBOOL controls whether `SETUP.iFitWidth/iFitHeight` are honored.
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

fn build_page_setup_scaling_sheet_stream_wsbool_last_wins(xf_general: u16) -> Vec<u8> {
    const WSBOOL_DEFAULT: u16 = 0x0C01;

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

    // First: enable fit-to-page.
    let wsbool_fit = WSBOOL_DEFAULT | WSBOOL_OPTION_FIT_TO_PAGE;
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool_fit.to_le_bytes());

    push_record(
        &mut sheet,
        RECORD_SETUP,
        &setup_record(
            1,     // Letter
            77,    // non-default percent
            2,     // fit width
            3,     // fit height
            false, // portrait
            0.3,   // header margin
            0.3,   // footer margin
        ),
    );

    // Last: disable fit-to-page (last record wins).
    let wsbool_percent = WSBOOL_DEFAULT;
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool_percent.to_le_bytes());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 0.0),
    );
    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn xnum(v: f64) -> [u8; 8] {
    v.to_le_bytes()
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
    let mut grbit = 0u16;
    if !landscape {
        grbit |= SETUP_GRBIT_F_PORTRAIT;
    }
    setup_record_with_grbit(
        paper_size,
        scale,
        fit_width,
        fit_height,
        grbit,
        header_margin,
        footer_margin,
    )
}

fn setup_record_with_grbit(
    paper_size: u16,
    scale: u16,
    fit_width: u16,
    fit_height: u16,
    grbit: u16,
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
    out.extend_from_slice(&grbit.to_le_bytes()); // grbit
    out.extend_from_slice(&600u16.to_le_bytes()); // iRes
    out.extend_from_slice(&600u16.to_le_bytes()); // iVRes
    out.extend_from_slice(&xnum(header_margin)); // numHdr
    out.extend_from_slice(&xnum(footer_margin)); // numFtr
    out.extend_from_slice(&1u16.to_le_bytes()); // iCopies
    out
}

fn build_page_setup_flags_sheet_stream(xf_cell: u16, setup: Vec<u8>) -> Vec<u8> {
    // WSBOOL: explicit fFitToPage=0 so SETUP.iScale semantics are stable for the fixture.
    build_page_setup_flags_sheet_stream_with_wsbool(xf_cell, setup, 0)
}

fn build_page_setup_flags_sheet_stream_with_wsbool(
    xf_cell: u16,
    setup: Vec<u8>,
    wsbool: u16,
) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // SETUP record under test.
    push_record(&mut sheet, RECORD_SETUP, &setup);

    push_record(&mut sheet, RECORD_WSBOOL, &wsbool.to_le_bytes());

    // A1: ensure calamine produces a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 0.0));

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn horizontal_page_breaks(break_rows_below: &[(u16, u16, u16)]) -> Vec<u8> {
    // HORIZONTALPAGEBREAKS record payload (BIFF8) [MS-XLS 2.4.122].
    // [cbrk:u16][(rw:u16,colStart:u16,colEnd:u16)*]
    let mut out = Vec::with_capacity(2 + 6 * break_rows_below.len());
    out.extend_from_slice(&(break_rows_below.len() as u16).to_le_bytes());
    for &(rw, col_start, col_end) in break_rows_below {
        out.extend_from_slice(&rw.to_le_bytes());
        out.extend_from_slice(&col_start.to_le_bytes());
        out.extend_from_slice(&col_end.to_le_bytes());
    }
    out
}

fn vertical_page_breaks(break_cols_right: &[(u16, u16, u16)]) -> Vec<u8> {
    // VERTICALPAGEBREAKS record payload (BIFF8) [MS-XLS 2.4.350].
    // [cbrk:u16][(col:u16,rwStart:u16,rwEnd:u16)*]
    let mut out = Vec::with_capacity(2 + 6 * break_cols_right.len());
    out.extend_from_slice(&(break_cols_right.len() as u16).to_le_bytes());
    for &(col, rw_start, rw_end) in break_cols_right {
        out.extend_from_slice(&col.to_le_bytes());
        out.extend_from_slice(&rw_start.to_le_bytes());
        out.extend_from_slice(&rw_end.to_le_bytes());
    }
    out
}

fn hpagebreaks_record(breaks: &[u16]) -> Vec<u8> {
    // HORIZONTALPAGEBREAKS payload:
    // [cbrk:u16][(rw:u16, colStart:u16, colEnd:u16) * cbrk]
    //
    // BIFF8 stores `rw` as the 0-based index of the *first* row below the break.
    // Our fixtures express breaks as the 0-based row index *after which* the break occurs, so we
    // encode them as `rw = after + 1`.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&(breaks.len() as u16).to_le_bytes());
    for &after in breaks {
        let rw = after.saturating_add(1);
        out.extend_from_slice(&rw.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // colStart
        out.extend_from_slice(&255u16.to_le_bytes()); // colEnd (BIFF8 max col for Excel 97-2003)
    }
    out
}

fn vpagebreaks_record(breaks: &[u16]) -> Vec<u8> {
    // VERTICALPAGEBREAKS payload:
    // [cbrk:u16][(col:u16, rwStart:u16, rwEnd:u16) * cbrk]
    //
    // BIFF8 stores `col` as the 0-based index of the *first* column to the right of the break.
    // Our fixtures express breaks as the 0-based column index *after which* the break occurs, so
    // we encode them as `col = after + 1`.
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&(breaks.len() as u16).to_le_bytes());
    for &after in breaks {
        let col = after.saturating_add(1);
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

/// Build a BIFF8 `.xls` fixture containing a worksheet with a corrupt page margin value.
///
/// This is used to validate that the importer ignores invalid BIFF margin records and surfaces a
/// warning, rather than importing NaN/Inf/negative/out-of-range values into `PageSetup`.
pub fn build_invalid_margins_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_invalid_margins_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a worksheet with a NaN page margin value.
///
/// This is used to validate that the importer ignores non-finite BIFF margin records and surfaces a
/// warning, rather than importing NaN/Inf into `PageSetup`.
pub fn build_invalid_margins_nan_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_invalid_margins_nan_workbook_stream();

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
    let workbook_stream = build_hyperlink_workbook_stream(
        "UrlNul",
        hlink_external_url(0, 0, 0, 0, url, "Example", "Tooltip"),
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

/// Build a BIFF8 `.xls` fixture containing a URL hyperlink whose URL moniker length field stores a
/// *character count* (UTF-16 code units) rather than a byte length.
///
/// Some producers are inconsistent about whether moniker string lengths are in bytes or chars. The
/// importer should accept both.
pub fn build_url_hyperlink_char_count_len_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "UrlLenChars",
        hlink_external_url_len_as_char_count(
            0,
            0,
            0,
            0,
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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_filtermode_hidden_rows_fixture_xls`], but
/// where the hidden ROW record corresponds to the AutoFilter header row (row 1 / 1-based).
///
/// FILTERMODE should only apply to data rows below the header row; header-row visibility should
/// remain user-hidden (not reclassified as filter-hidden).
pub fn build_autofilter_filtermode_hidden_header_row_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_filtermode_hidden_header_row_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_filtermode_hidden_rows_fixture_xls`], but
/// where the hidden ROW record corresponds to an outline-hidden row (collapsed group).
///
/// The importer should preserve outline-hidden rows and not reclassify them as filter-hidden rows.
pub fn build_autofilter_filtermode_outline_hidden_rows_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_filtermode_outline_hidden_rows_workbook_stream();

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

fn build_autofilter_filtermode_hidden_header_row_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilteredHiddenHeaderRow");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`).
    let filter_db_rgce = ptg_area(0, 2, 0, 1); // $A$1:$B$3
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_autofilter_filtermode_hidden_header_row_sheet_stream(xf_cell);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_filtermode_outline_hidden_rows_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilteredOutlineHiddenRows");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`).
    //
    // Use a filter range that includes the outline group summary row so the outline-hidden detail
    // rows lie within the filter data range.
    let filter_db_rgce = ptg_area(0, 3, 0, 1); // $A$1:$B$4
    push_record(
        &mut globals,
        RECORD_NAME,
        &builtin_name_record(true, 1, 0x0D, &filter_db_rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_autofilter_filtermode_outline_hidden_rows_sheet_stream(xf_cell);

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

fn build_autofilter_filterdatabase_arean_filtermode_hidden_row_sheet_stream(
    xf_cell: u16,
) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 10) cols [0, 4) => A1:D10.
    // This intentionally exceeds the workbook-scoped `_FilterDatabase` range (A1:B3) so the
    // importer initially infers a larger AutoFilter range from the worksheet stream.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&10u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&4u16.to_le_bytes()); // last col + 1 (A..D)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Mark row 4 (1-based) as hidden via the ROW record. This row falls inside the DIMENSIONS-based
    // inferred AutoFilter range (A1:D10), but outside the true `_FilterDatabase` range (A1:B3).
    push_record(&mut sheet, RECORD_ROW, &row_record(3, true, 0, false));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: 4 columns (A..D). This will cause DIMENSIONS-based inference to yield A1:D10.
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &4u16.to_le_bytes());
    // FILTERMODE indicates an active filter state (filtered rows).
    push_record(&mut sheet, RECORD_FILTERMODE, &[]);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_autofilter_filtermode_hidden_header_row_sheet_stream(xf_cell: u16) -> Vec<u8> {
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

    // Mark row 1 (1-based) as hidden via the ROW record.
    // When FILTERMODE is present, the importer should *not* reclassify the header row as a
    // filter-hidden row.
    push_record(&mut sheet, RECORD_ROW, &row_record(0, true, 0, false));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: 2 columns (A..B).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &2u16.to_le_bytes());
    // AUTOFILTER: filter criteria for column 0 (`colId=0` in A1:B3).
    //
    // Criterion: column value equals "X".
    let doper1 = autofilter_doper_string(AUTOFILTER_OP_EQUAL, "X");
    let doper2 = autofilter_doper_none();
    let autofilter = autofilter_record(0, false, &doper1, &doper2);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);
    // FILTERMODE indicates an active filter state (filtered rows).
    push_record(&mut sheet, RECORD_FILTERMODE, &[]);

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
    // When FILTERMODE is present, the importer should treat this as a filter-hidden row (not a
    // user-hidden row).
    push_record(&mut sheet, RECORD_ROW, &row_record(1, true, 0, false));

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: 2 columns (A..B).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &2u16.to_le_bytes());
    // AUTOFILTER: filter criteria for column 0 (`colId=0` in A1:B3).
    //
    // Criterion: column value equals "X".
    let doper1 = autofilter_doper_string(AUTOFILTER_OP_EQUAL, "X");
    let doper2 = autofilter_doper_none();
    let autofilter = autofilter_record(0, false, &doper1, &doper2);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);
    // FILTERMODE indicates an active filter state (filtered rows).
    push_record(&mut sheet, RECORD_FILTERMODE, &[]);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_autofilter_filtermode_outline_hidden_rows_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 4) cols [0, 2) so A1:B4 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&4u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // WSBOOL: keep Excel's default worksheet boolean options so outline summary flags decode.
    push_record(&mut sheet, RECORD_WSBOOL, &0x0C01u16.to_le_bytes());

    // Outline rows:
    // - Rows 2-3 (1-based) are detail rows: outline level 1 and hidden (collapsed).
    // - Row 4 (1-based) is the collapsed summary row (level 0, collapsed).
    push_record(&mut sheet, RECORD_ROW, &row_record(1, true, 1, false));
    push_record(&mut sheet, RECORD_ROW, &row_record(2, true, 1, false));
    push_record(&mut sheet, RECORD_ROW, &row_record(3, false, 0, true));

    // A1: NUMBER cell to seed calamine's range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: 2 columns (A..B).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &2u16.to_le_bytes());
    // FILTERMODE: present (no payload).
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

/// Build a BIFF8 `.xls` fixture with worksheet protection enabled where the enhanced allow-mask is
/// stored *after* some prefix bytes inside the `FEATHEADR` payload.
///
/// This exercises the importer's best-effort scanning for the allow-mask within FEAT/FEATHEADR
/// payloads (some producers embed the mask inside a larger structure).
pub fn build_sheet_protection_allow_flags_mask_offset_fixture_xls() -> Vec<u8> {
    let sheet_stream = build_sheet_protection_allow_flags_mask_offset_sheet_stream();
    let workbook_stream =
        build_single_sheet_workbook_stream("ProtectedAllowFlagsMaskOffset", &sheet_stream, 1252);

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

/// Build a BIFF8 `.xls` fixture that includes horizontal/vertical manual page breaks with edge
/// cases:
/// - horizontal break row=0 (invalid; represents a break before first row)
/// - vertical break col=0 (invalid; represents a break before first col)
pub fn build_page_break_edge_cases_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_page_break_edge_cases_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture that exercises `SETUP.grbit.fNoPls` handling for page setup import.
///
/// The worksheet contains:
/// - `WSBOOL.fFitToPage=1`
/// - `SETUP.fNoPls=1` with:
///   - `iFitWidth=2`, `iFitHeight=3`
///   - `numHdr=0.9`, `numFtr=1.1`
///   - but also non-default `iScale=80`, `iPaperSize=9` (A4), and a landscape orientation bit
///
/// Expected import behavior:
/// - Fit-to scaling + header/footer margins are imported
/// - Paper size / orientation / percent scaling are ignored due to `fNoPls`
pub fn build_setup_fnopls_fit_to_page_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_setup_fnopls_fit_to_page_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a single sheet named `FilterCriteria` with:
/// - a sheet-scoped `_xlnm._FilterDatabase` defined name referencing `$A$1:$C$5`, and
/// - BIFF8 `AUTOFILTER` records storing simple per-column filter criteria (e.g. column A equals
///   "Alice").
///
/// This exercises end-to-end import of `SheetAutoFilter.filter_columns` from legacy BIFF8
/// `AUTOFILTER (0x009E)` records.
pub fn build_autofilter_criteria_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_fixture_xls`], but with a
/// multi-criterion `AUTOFILTER` record whose two conditions are joined with AND.
///
/// The sheet is named `FilterCriteriaJoinAll` and the `_xlnm._FilterDatabase` defined name points
/// at `$A$1:$A$5`.
pub fn build_autofilter_criteria_join_all_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_join_all_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_fixture_xls`], but with
/// `AUTOFILTER` records that use the legacy BIFF operator codes for BETWEEN / NOT BETWEEN.
///
/// The sheet is named `FilterCriteriaBetween` and the `_xlnm._FilterDatabase` defined name points at
/// `$A$1:$B$5`.
pub fn build_autofilter_criteria_between_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_between_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_fixture_xls`], but with
/// `AUTOFILTER` records that represent blank / non-blank criteria.
///
/// The sheet is named `FilterCriteriaBlanks` and the `_xlnm._FilterDatabase` defined name points at
/// `$A$1:$D$5`.
pub fn build_autofilter_criteria_blanks_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_blanks_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_fixture_xls`], but with
/// `AUTOFILTER` records that use BIFF text operator codes (`contains`, `beginsWith`, `endsWith`).
///
/// These are preserved as `FilterCriterion::OpaqueCustom` so they can round-trip to XLSX.
///
/// The sheet is named `FilterCriteriaTextOps` and the `_xlnm._FilterDatabase` defined name points at
/// `$A$1:$C$5`.
pub fn build_autofilter_criteria_text_ops_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_text_ops_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_text_ops_fixture_xls`], but using
/// BIFF operator codes for the *negative* text operators (`doesNotContain`, `doesNotBeginWith`,
/// `doesNotEndWith`).
///
/// These are preserved as `FilterCriterion::OpaqueCustom` so they can round-trip to XLSX.
///
/// The sheet is named `FilterCriteriaTextOpsNeg` and the `_xlnm._FilterDatabase` defined name points
/// at `$A$1:$C$5`.
pub fn build_autofilter_criteria_text_ops_negative_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_text_ops_negative_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_fixture_xls`], but with
/// `AUTOFILTER` records that store boolean values (vt=BOOL).
///
/// The sheet is named `FilterCriteriaBool` and the `_xlnm._FilterDatabase` defined name points at
/// `$A$1:$B$5`.
pub fn build_autofilter_criteria_bool_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_bool_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_fixture_xls`], but with an
/// `AUTOFILTER` record that encodes a legacy Top10 filter (via `AUTOFILTER.grbit` flags).
///
/// The sheet is named `FilterCriteriaTop10` and the `_xlnm._FilterDatabase` defined name points at
/// `$A$1:$A$5`.
pub fn build_autofilter_criteria_top10_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_top10_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_fixture_xls`], but where the
/// DOPER operator code is stored in the second byte ("grbit") rather than `wOper`.
///
/// Some producers have been observed to write `AUTOFILTER.DOPER.grbit=<op>` while leaving `wOper=0`;
/// our parser handles this best-effort by falling back to the byte1 value.
///
/// The sheet is named `FilterCriteriaOpByte1` and the `_xlnm._FilterDatabase` defined name points at
/// `$A$1:$A$5`.
pub fn build_autofilter_criteria_operator_byte1_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_operator_byte1_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_fixture_xls`], but where the
/// `AUTOFILTER.iEntry` field stores an **absolute worksheet column index** (observed in some
/// producers) instead of an index relative to the AutoFilter range start.
///
/// The sheet is named `FilterCriteriaAbsEntry` and the `_xlnm._FilterDatabase` defined name points
/// at `$D$1:$F$5` (range start is column D, index 3).
pub fn build_autofilter_criteria_absolute_entry_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_absolute_entry_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_criteria_fixture_xls`], but where the
/// trailing `XLUnicodeString` payload for the text criterion is split across an `AUTOFILTER` record
/// followed by a `CONTINUE` record.
///
/// This exercises end-to-end continued-record handling in the legacy BIFF8 AutoFilter criteria
/// importer.
pub fn build_autofilter_criteria_continued_string_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_criteria_continued_string_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a single sheet named `FilterSort12` with a
/// sheet-scoped `_xlnm._FilterDatabase` defined name referencing `$A$1:$C$5`, plus BIFF8 future
/// records (`Sort12`/`SortData12`) describing an AutoFilter sort state.
///
/// This exercises best-effort import of `SheetAutoFilter.sort_state` from Excel 2007+ sort metadata
/// stored in BIFF8 via Future Record Type records.
pub fn build_autofilter_sort12_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_sort12_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_sort12_fixture_xls`], but with the `Sort12`
/// payload split across `Sort12` + `ContinueFrt12`.
///
/// This exercises best-effort continuation handling for future sort records: import should not
/// panic and should recover the sort state when possible.
pub fn build_autofilter_sort12_continuefrt12_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter_sort12_continuefrt12_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a single sheet with an inferred AutoFilter range
/// (DIMENSIONS + AUTOFILTERINFO) and a large number of malformed BIFF8 `SORT` records.
///
/// Each `SORT` record yields a best-effort warning during the importer's sort-state recovery pass.
/// This is intended to stress-test the global warning cap.
pub fn build_many_sort_state_warnings_fixture_xls(record_count: usize) -> Vec<u8> {
    let workbook_stream = build_many_sort_state_warnings_workbook_stream(record_count);

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

fn build_many_sort_state_warnings_workbook_stream(record_count: usize) -> Vec<u8> {
    // `build_single_sheet_workbook_stream` writes 16 style XFs followed by one cell XF, so the
    // first cell XF index is 16.
    let xf_cell = 16u16;
    let sheet_stream = build_many_sort_state_warnings_sheet_stream(xf_cell, record_count);
    build_single_sheet_workbook_stream("SortWarnings", &sheet_stream, 1252)
}

fn build_many_sort_state_warnings_sheet_stream(xf_cell: u16, record_count: usize) -> Vec<u8> {
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

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // AUTOFILTERINFO: cEntries = 3 (A..C).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &3u16.to_le_bytes());

    // Malformed/unsupported SORT records: valid shape but no usable key columns.
    // The importer should emit a warning per record while still continuing the import.
    let sort_payload = sort_record_payload(0, 4, 0, 2, true, &[]);

    for _ in 0..record_count {
        push_record(&mut sheet, RECORD_SORT, &sort_payload);
    }

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
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

/// Build a minimal BIFF8 `.xls` fixture like [`build_autofilter12_fixture_xls`], but with the
/// `AutoFilter12` payload split across an `AutoFilter12` record + a `ContinueFrt12` record.
///
/// This exercises best-effort continuation handling for future records: import should not panic and
/// should recover the full multi-value filter list when possible.
pub fn build_autofilter12_continuefrt12_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_autofilter12_continuefrt12_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture like [`build_autofilter_workbook_scope_unqualified_multisheet_fixture_xls`],
/// but with `FILTERMODE` and a user-hidden row in the worksheet stream that falls **outside** the
/// true `_FilterDatabase` range and therefore must not be reclassified as filter-hidden.
pub fn build_autofilter_workbook_scope_unqualified_multisheet_filtermode_hidden_row_fixture_xls(
) -> Vec<u8> {
    let workbook_stream =
        build_autofilter_workbook_scope_unqualified_multisheet_filtermode_hidden_row_workbook_stream(
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

fn build_autofilter_workbook_scope_unqualified_multisheet_filtermode_hidden_row_workbook_stream(
) -> Vec<u8> {
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
    let sheet1 = build_autofilter_filterdatabase_arean_filtermode_hidden_row_sheet_stream(xf_cell);
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

fn build_page_break_edge_cases_workbook_stream() -> Vec<u8> {
    // build_single_sheet_workbook_stream constructs a minimal workbook with 16 style XFs followed by
    // a single cell XF (General). That cell XF has index 16.
    const XF_CELL: u16 = 16;
    let sheet_stream = build_page_break_edge_cases_sheet_stream(XF_CELL);
    build_single_sheet_workbook_stream("PageBreaks", &sheet_stream, 1252)
}

fn build_page_break_edge_cases_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 10) cols [0, 5) => A1:E10.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&10u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&5u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // HORIZONTALPAGEBREAKS: two breaks (row=0 is invalid, row=5 should import as break-after=4).
    let mut hpb = Vec::<u8>::new();
    hpb.extend_from_slice(&2u16.to_le_bytes()); // cbrk
    hpb.extend_from_slice(&0u16.to_le_bytes()); // row
    hpb.extend_from_slice(&0u16.to_le_bytes()); // colStart
    hpb.extend_from_slice(&0u16.to_le_bytes()); // colEnd
    hpb.extend_from_slice(&5u16.to_le_bytes()); // row
    hpb.extend_from_slice(&0u16.to_le_bytes()); // colStart
    hpb.extend_from_slice(&0u16.to_le_bytes()); // colEnd
    push_record(&mut sheet, RECORD_HORIZONTALPAGEBREAKS, &hpb);

    // VERTICALPAGEBREAKS: two breaks (col=0 is invalid, col=3 should import as break-after=2).
    let mut vpb = Vec::<u8>::new();
    vpb.extend_from_slice(&2u16.to_le_bytes()); // cbrk
    vpb.extend_from_slice(&0u16.to_le_bytes()); // col
    vpb.extend_from_slice(&0u16.to_le_bytes()); // rwStart
    vpb.extend_from_slice(&0u16.to_le_bytes()); // rwEnd
    vpb.extend_from_slice(&3u16.to_le_bytes()); // col
    vpb.extend_from_slice(&0u16.to_le_bytes()); // rwStart
    vpb.extend_from_slice(&0u16.to_le_bytes()); // rwEnd
    push_record(&mut sheet, RECORD_VERTICALPAGEBREAKS, &vpb);

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
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

fn build_fit_to_clamp_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // WSBOOL.fFitToPage is stored in bit 0x0100. Use the typical Excel defaults (0x0C01) plus that
    // bit.
    let wsbool = (0x0C01u16 | 0x0100u16).to_le_bytes();
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool);

    // SETUP [MS-XLS 2.4.257]
    // Most fields are boilerplate; the important part is iFitWidth/iFitHeight are intentionally
    // out of range (>32767).
    let mut setup = Vec::<u8>::new();
    setup.extend_from_slice(&9u16.to_le_bytes()); // iPaperSize (A4)
    setup.extend_from_slice(&100u16.to_le_bytes()); // iScale (percent)
    setup.extend_from_slice(&1u16.to_le_bytes()); // iPageStart
    setup.extend_from_slice(&40000u16.to_le_bytes()); // iFitWidth (invalid)
    setup.extend_from_slice(&40000u16.to_le_bytes()); // iFitHeight (invalid)
    setup.extend_from_slice(&0x0083u16.to_le_bytes()); // grbit
    setup.extend_from_slice(&300u16.to_le_bytes()); // iRes
    setup.extend_from_slice(&300u16.to_le_bytes()); // iVRes
    setup.extend_from_slice(&0.1f64.to_le_bytes()); // numHdr
    setup.extend_from_slice(&0.1f64.to_le_bytes()); // numFtr
    setup.extend_from_slice(&1u16.to_le_bytes()); // iCopies
    push_record(&mut sheet, RECORD_SETUP, &setup);

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_invalid_margins_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream("Sheet1", &build_invalid_margins_sheet_stream(), 1252)
}

fn build_invalid_margins_sheet_stream() -> Vec<u8> {
    build_invalid_margins_sheet_stream_with_left_margin(50.0f64.to_le_bytes())
}

fn build_invalid_margins_nan_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "Sheet1",
        &build_invalid_margins_sheet_stream_with_left_margin(f64::NAN.to_le_bytes()),
        1252,
    )
}

fn build_invalid_margins_sheet_stream_with_left_margin(invalid_left: [u8; 8]) -> Vec<u8> {
    // The workbook globals above create 16 style XFs + 1 cell XF, so the first usable
    // cell XF index is 16.
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // LEFTMARGIN: emit a valid margin followed by an invalid one to validate that invalid values do
    // not clobber earlier valid values ("last valid wins").
    push_record(&mut sheet, RECORD_LEFTMARGIN, &1.0f64.to_le_bytes());
    push_record(&mut sheet, RECORD_LEFTMARGIN, &invalid_left);

    // Other margins: invalid-only (these should remain at model defaults).
    push_record(&mut sheet, RECORD_RIGHTMARGIN, &(-1.0f64).to_le_bytes());
    push_record(&mut sheet, RECORD_TOPMARGIN, &f64::NAN.to_le_bytes());
    push_record(
        &mut sheet,
        RECORD_BOTTOMMARGIN,
        &f64::INFINITY.to_le_bytes(),
    );

    // SETUP record with invalid header/footer margins.
    push_record(
        &mut sheet,
        RECORD_SETUP,
        &setup_record(
            1,     // Letter (default)
            100,   // iScale (default)
            0,     // iFitWidth
            0,     // iFitHeight
            false, // portrait (default)
            -1.0,  // header margin (invalid)
            50.0,  // footer margin (invalid)
        ),
    );

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, XF_GENERAL_CELL, 1.0),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
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

fn build_shared_formula_ptgexp_wide_payload_missing_shrfmla_workbook_stream() -> Vec<u8> {
    // Like `build_shared_formula_ptgexp_missing_shrfmla_workbook_stream`, but the follower `PtgExp`
    // uses a non-standard row u32 + col u16 payload width.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgexp_wide_payload_missing_shrfmla_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedWide", &sheet_stream, 1252)
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

fn build_shared_formula_ptgexp_wide_payload_missing_shrfmla_sheet_stream(xf_cell: u16) -> Vec<u8> {
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

    // Follower formula in B2: `PtgExp` pointing at base cell B1 (row=0, col=1), encoded as
    // [ptg:0x01][row:u32][col:u16].
    let rgce_ptgexp = ptg_exp_row_u32_col_u16(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &rgce_ptgexp),
    );

    // Intentionally omit SHRFMLA/ARRAY definition records.

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SharedFormulaBaseKind {
    MissingFormulaRecord,
    PtgExpSelf,
}

fn build_shared_formula_shrfmla_only_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One cell XF (General).
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Two worksheets: MissingBase and DegenerateBase.
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["MissingBase", "DegenerateBase"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet 0: MissingBase -----------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_formula_shrfmla_only_sheet_stream(
        xf_cell,
        SharedFormulaBaseKind::MissingFormulaRecord,
    ));

    // -- Sheet 1: DegenerateBase --------------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_formula_shrfmla_only_sheet_stream(
        xf_cell,
        SharedFormulaBaseKind::PtgExpSelf,
    ));

    globals
}

fn build_shared_formula_shrfmla_only_sheet_stream(
    xf_cell: u16,
    base_kind: SharedFormulaBaseKind,
) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0,2) cols [0,2) => A1:B2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Provide value cells in column A so the shared formula has something to reference.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0)); // A1

    // Shared formula rgce: `A(row)+1` for B1:B2.
    // In BIFF8 this uses `PtgRefN` offsets relative to the cell containing the formula.
    // - row_off = 0
    // - col_off = -1 (left)
    let shared_rgce: [u8; 9] = [
        0x2C, // PtgRefN
        0x00, 0x00, // row_off = 0 (i16 stored in u16 field)
        0xFF, 0xFF, // col_off = -1 (14-bit two's complement + relative bits)
        0x1E, // PtgInt
        0x01, 0x00, // 1
        0x03, // PtgAdd
    ];

    // Depending on the variant, emit a degenerate base FORMULA record or omit it entirely.
    if base_kind == SharedFormulaBaseKind::PtgExpSelf {
        let ptgexp_self: [u8; 5] = [
            0x01, // PtgExp
            0x00, 0x00, // rw=0 (B1)
            0x01, 0x00, // col=1 (B)
        ];
        push_record(
            &mut sheet,
            RECORD_FORMULA,
            &formula_cell(0, 1, xf_cell, 0.0, &ptgexp_self),
        );
    }

    // SHRFMLA defines the shared formula range and token stream.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &shared_rgce),
    );

    // A2 value cell (row 1).
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0)); // A2

    // B2: FORMULA record containing only `PtgExp(B1)`.
    let ptgexp_b1: [u8; 5] = [
        0x01, // PtgExp
        0x00, 0x00, // rw=0 (B1)
        0x01, 0x00, // col=1 (B)
    ];
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &ptgexp_b1),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_shrfmla_continued_ptgstr_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_shrfmla_continued_ptgstr_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedStr", &sheet_stream, 1252)
}

fn build_shared_formula_shrfmla_continued_ptgstr_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0,2) cols [0,2) => A1:B2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0)); // A1

    // Shared formula rgce: a string literal `"ABCDE"` represented as a single `PtgStr`.
    //
    // We split the physical SHRFMLA record across `CONTINUE` boundaries inside the string payload.
    // The continued fragment starts with the required 1-byte continued-segment option flags prefix.
    let literal = b"ABCDE";
    let rgce: Vec<u8> = [vec![0x17, literal.len() as u8, 0u8], literal.to_vec()].concat();
    let cce = rgce.len() as u16;

    // Split after `"AB"` inside the string payload.
    let first_rgce = &rgce[..(3 + 2)]; // ptg+len+flags + 2 chars
    let remaining = &literal[2..]; // "CDE"

    // SHRFMLA header: RefU (6) + cUse (2) + cce (2) + rgce...
    //
    // Range B1:B2 => rows 0..1, col 1.
    let mut shrfmla_part1 = Vec::new();
    shrfmla_part1.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
    shrfmla_part1.extend_from_slice(&1u16.to_le_bytes()); // rwLast
    shrfmla_part1.push(1u8); // colFirst (B)
    shrfmla_part1.push(1u8); // colLast (B)
    shrfmla_part1.extend_from_slice(&0u16.to_le_bytes()); // cUse
    shrfmla_part1.extend_from_slice(&cce.to_le_bytes());
    shrfmla_part1.extend_from_slice(first_rgce);
    push_record(&mut sheet, RECORD_SHRFMLA, &shrfmla_part1);

    // CONTINUE fragment: [continued_segment_flags][remaining string bytes]
    let mut cont = Vec::new();
    cont.push(0u8); // continued segment option flags (compressed 8-bit)
    cont.extend_from_slice(remaining);
    push_record(&mut sheet, RECORD_CONTINUE, &cont);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_ptgexp_u32_row_u16_col_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgexp_u32_row_u16_col_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("Shared", &sheet_stream, 1252)
}

fn build_shared_formula_ptgexp_u32_row_u16_col_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Worksheet containing:
    // - Numeric inputs in A2/A3
    // - Shared formula range B2:B3 whose rgce is `A?*2`
    // - Follower cell B3 stores `PtgExp` coordinates as row u32 + col u16 (6 bytes)
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [1, 3) cols [0, 2) => A2:B3.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&1u32.to_le_bytes()); // first row
    dims.extend_from_slice(&3u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Input values.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 10.0)); // A2
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(2, 0, xf_cell, 20.0)); // A3

    // Shared formula rgce: `PtgRefN(row+0,col-1) * 2`.
    let shared_rgce: Vec<u8> = vec![
        0x2C, // PtgRefN
        0x00, 0x00, // row offset (0)
        0xFF, 0xFF, // col offset (-1)
        0x1E, // PtgInt
        0x02, 0x00, // 2
        0x05, // PtgMul
    ];

    // Anchor cell B2 (row=1,col=1): store canonical PtgExp(B2) followed by SHRFMLA containing the
    // shared rgce.
    let anchor_ptgexp = ptg_exp(1, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &anchor_ptgexp),
    );
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(1, 2, 1, 1, &shared_rgce),
    );

    // Follower cell B3 (row=2,col=1): PtgExp with a 6-byte payload (row u32 + col u16).
    let follower_ptgexp = ptg_exp_row_u32_col_u16(1, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(2, 1, xf_cell, 0.0, &follower_ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("Sheet1", &sheet_stream, 1252)
}

fn build_shared_formula_sheet_stream(xf_cell: u16) -> Vec<u8> {
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

    // A1/A2: number cells (values not important for formula decoding).
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));

    // B1: base formula cell, `A1+1` (PtgRef + PtgInt + PtgAdd).
    let rgce_b1: Vec<u8> = [
        vec![0x24],                       // PtgRef
        0u16.to_le_bytes().to_vec(),      // rw = 0 (A1 row)
        0xC000u16.to_le_bytes().to_vec(), // col = 0 with row+col relative flags (A)
        vec![0x1E],                       // PtgInt
        1u16.to_le_bytes().to_vec(),      // 1
        vec![0x03],                       // PtgAdd
    ]
    .concat();
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &rgce_b1),
    );

    // Shared formula token stream for B1:B2: `A?+1` using `PtgRefN` so the reference shifts with
    // the base cell. From B1 the ref is A1 (col_off=-1); from B2 it becomes A2.
    let rgce_shared: Vec<u8> = [
        vec![0x2C],                       // PtgRefN
        0u16.to_le_bytes().to_vec(),      // row_off = 0
        0xFFFFu16.to_le_bytes().to_vec(), // col_off = -1 (14-bit) + row/col relative flags
        vec![0x1E],                       // PtgInt
        1u16.to_le_bytes().to_vec(),      // 1
        vec![0x03],                       // PtgAdd
    ]
    .concat();
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &rgce_shared),
    );

    // B2: FORMULA record contains only PtgExp referencing the base cell (B1).
    let rgce_b2: Vec<u8> = [
        vec![0x01],
        0u16.to_le_bytes().to_vec(),
        1u16.to_le_bytes().to_vec(),
    ]
    .concat();
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &rgce_b2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_shrfmla_only_continued_ptgstr_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_shrfmla_only_continued_ptgstr_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ShrfmlaContinue", &sheet_stream, 1252)
}

fn build_shared_formula_shrfmla_only_continued_ptgstr_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0,2) cols [0,2) => A1:B2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Provide value cells in column A so the shared formula has something to reference.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0)); // A1
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0)); // A2

    // Shared formula rgce: `A(row)&"ABCDE"` for B1:B2, with the PtgStr payload split across a
    // CONTINUE boundary.
    let shared_rgce: Vec<u8> = vec![
        // PtgRefN(row_off=0,col_off=-1).
        0x2C, 0x00, 0x00, // row_off = 0
        0xFF, 0xFF, // col_off = -1 with relative bits
        // PtgStr("ABCDE") [ptg=0x17][cch=5][flags=0][chars...]
        0x17, 5, 0, b'A', b'B', b'C', b'D', b'E', // PtgConcat
        0x08,
    ];

    // Build the full SHRFMLA payload then split it into SHRFMLA + CONTINUE fragments that cut the
    // PtgStr token after the first two characters ("AB"). Excel inserts a 1-byte continuation flags
    // byte at the start of the continued fragment; include that so fragment-aware rgce parsing can
    // reconstruct the canonical token stream.
    let mut shrfmla_full = Vec::<u8>::new();
    shrfmla_full.extend_from_slice(&0u16.to_le_bytes()); // rwFirst = 0
    shrfmla_full.extend_from_slice(&1u16.to_le_bytes()); // rwLast = 1
    shrfmla_full.push(1); // colFirst = B
    shrfmla_full.push(1); // colLast = B
    shrfmla_full.extend_from_slice(&0u16.to_le_bytes()); // cUse
    shrfmla_full.extend_from_slice(&(shared_rgce.len() as u16).to_le_bytes()); // cce
    shrfmla_full.extend_from_slice(&shared_rgce);

    // Split point: header (10) + PtgRefN (5) + PtgStr header (3) + 2 chars = 20 bytes.
    let split_at = 10usize + 5 + 3 + 2;
    let shrfmla_frag1 = &shrfmla_full[..split_at];
    let shrfmla_remaining = &shrfmla_full[split_at..];

    let mut shrfmla_continue = Vec::<u8>::new();
    shrfmla_continue.push(0); // continued-segment option flags (compressed)
    shrfmla_continue.extend_from_slice(shrfmla_remaining);

    push_record(&mut sheet, RECORD_SHRFMLA, shrfmla_frag1);
    push_record(&mut sheet, RECORD_CONTINUE, &shrfmla_continue);

    // B2: FORMULA record containing only `PtgExp(B1)`.
    let ptgexp_b1: [u8; 5] = [
        0x01, // PtgExp
        0x00, 0x00, // rw=0 (B1)
        0x01, 0x00, // col=1 (B)
    ];
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &ptgexp_b1),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_shrfmla_only_ptgarray_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_shrfmla_only_ptgarray_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ShrfmlaArray", &sheet_stream, 1252)
}

fn build_shared_formula_ptgexp_ptgarray_warning_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgexp_ptgarray_warning_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ShrfmlaArrayWarn", &sheet_stream, 1252)
}

fn build_shared_formula_shrfmla_only_ptgarray_sheet_stream(xf_cell: u16) -> Vec<u8> {
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
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));

    // Shared formula body stored in SHRFMLA (range B1:B2):
    //   PtgRefN (A(row)) + PtgArray + PtgFuncVar(SUM) + PtgAdd
    let rgce_shared = {
        let mut v = Vec::new();
        // PtgRefN: row_off=0, col_off=-1 relative to the formula cell.
        v.push(0x2C);
        v.extend_from_slice(&0u16.to_le_bytes()); // row_off = 0
        v.extend_from_slice(&0xFFFFu16.to_le_bytes()); // col_off = -1 (14-bit), row+col relative

        // PtgArray (array constant; data stored in trailing rgcb).
        v.push(0x20);
        v.extend_from_slice(&[0u8; 7]); // reserved

        // PtgFuncVar: SUM(argc=1).
        v.push(0x22);
        v.push(1);
        v.extend_from_slice(&4u16.to_le_bytes());

        // PtgAdd.
        v.push(0x03);
        v
    };

    // rgcb payload for the array constant `{1,2;3,4}`.
    let rgcb = rgcb_array_constant_numbers_2x2(&[1.0, 2.0, 3.0, 4.0]);

    // SHRFMLA record defining shared rgce + rgcb for the range B1:B2.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record_with_rgcb(0, 1, 1, 1, &rgce_shared, &rgcb),
    );

    // B2 formula record: PtgExp(B1).
    let ptgexp = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_ptgexp_ptgarray_warning_sheet_stream(xf_cell: u16) -> Vec<u8> {
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
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));

    // Shared formula body stored in SHRFMLA (range B1:B2):
    //   PtgRefN (A(row)) + PtgArray + PtgFunc(SUM) + PtgAdd
    //
    // Note: `SUM` is canonically encoded via `PtgFuncVar`, but we intentionally use `PtgFunc`
    // (fixed-arity token) to trigger a decode warning. This ensures the formula-override pass
    // (which only applies overrides when decoding yields no warnings) does not mask failures in the
    // earlier `PtgExp` recovery path.
    let rgce_shared = {
        let mut v = Vec::new();
        // PtgRefN: row_off=0, col_off=-1 relative to the formula cell.
        v.push(0x2C);
        v.extend_from_slice(&0u16.to_le_bytes()); // row_off = 0
        v.extend_from_slice(&0xFFFFu16.to_le_bytes()); // col_off = -1 (14-bit), row+col relative

        // PtgArray (array constant; data stored in trailing rgcb).
        v.push(0x20);
        v.extend_from_slice(&[0u8; 7]); // reserved

        // PtgFunc: SUM (function id 4).
        v.push(0x21);
        v.extend_from_slice(&4u16.to_le_bytes());

        // PtgAdd.
        v.push(0x03);
        v
    };

    // rgcb payload for the array constant `{1,2;3,4}`.
    let rgcb = rgcb_array_constant_numbers_2x2(&[1.0, 2.0, 3.0, 4.0]);

    // SHRFMLA record defining shared rgce + rgcb for the range B1:B2.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record_with_rgcb(0, 1, 1, 1, &rgce_shared, &rgcb),
    );

    // B2 formula record: PtgExp(B1).
    let ptgexp = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_shrfmla_range_cuse_ambiguity_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_shrfmla_range_cuse_ambiguity_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ShrfmlaCuseAmbiguity", &sheet_stream, 1252)
}

fn build_shared_formula_shrfmla_range_cuse_ambiguity_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 10) cols [0, 3) => A1:C10.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&10u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide numeric inputs in C1/C2 so the shared formula `C(row)+1` has in-range references.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 2, xf_cell, 1.0)); // C1
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 2, xf_cell, 2.0)); // C2

    // -- SHRFMLA #1 --------------------------------------------------------------
    // Range A1:A2 encoded as RefU + cUse + cce, with a *non-zero* cUse value to trigger the RefU vs
    // Ref8 ambiguity (when interpreted incorrectly as Ref8, `cUse` can be treated as colLast).
    let rgce_left: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x17); // PtgStr
        v.push(4); // cch
        v.push(0); // flags (compressed)
        v.extend_from_slice(b"LEFT");
        v
    };

    let mut shrfmla_left = Vec::<u8>::new();
    shrfmla_left.extend_from_slice(&0u16.to_le_bytes()); // rwFirst = 0
    shrfmla_left.extend_from_slice(&1u16.to_le_bytes()); // rwLast = 1
    shrfmla_left.push(0); // colFirst = A
    shrfmla_left.push(0); // colLast = A
    shrfmla_left.extend_from_slice(&2u16.to_le_bytes()); // cUse = 2 (non-zero)
    shrfmla_left.extend_from_slice(&(rgce_left.len() as u16).to_le_bytes()); // cce
    shrfmla_left.extend_from_slice(&rgce_left);
    push_record(&mut sheet, RECORD_SHRFMLA, &shrfmla_left);

    // -- SHRFMLA #2 --------------------------------------------------------------
    // Shared formula range B1:B10 whose body is `C(row)+1`.
    // PtgRefN(row_off=0,col_off=+1) + PtgInt(1) + PtgAdd
    let rgce_right: Vec<u8> = vec![
        0x2C, // PtgRefN
        0x00, 0x00, // row_off = 0
        0x01, 0xC0, // col_off = +1 with row+col relative flags
        0x1E, // PtgInt
        0x01, 0x00, // 1
        0x03, // PtgAdd
    ];
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 9, 1, 1, &rgce_right),
    );

    // B2 FORMULA record: PtgExp(B2) (self-reference).
    let ptgexp_b2 = ptg_exp(1, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &ptgexp_b2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_shrfmla_ref8_header_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_shrfmla_ref8_header_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedRef8", &sheet_stream, 1252)
}

fn build_shared_formula_shrfmla_ref8_header_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 2) cols [0, 3) => A1:C2.
    // The shared formula range itself is only `A1:B2`; we include `C1` as a value cell so the
    // sheet is non-empty without overlapping the formula range.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // C1: a numeric cell so calamine sees a non-empty sheet value range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 2, xf_cell, 0.0));

    // Shared formula body stored in SHRFMLA: `"X"` (PtgStr).
    let rgce: [u8; 4] = [0x17, 1, 0, b'X']; // [PtgStr][cch=1][flags=0][X]

    // SHRFMLA definition for range A1:B2 using a Ref8 range header:
    //   [rwFirst:u16][rwLast:u16][colFirst:u16][colLast:u16]
    let mut shrfmla = Vec::<u8>::new();
    shrfmla.extend_from_slice(&0u16.to_le_bytes()); // rwFirst = 0
    shrfmla.extend_from_slice(&1u16.to_le_bytes()); // rwLast = 1
    shrfmla.extend_from_slice(&0u16.to_le_bytes()); // colFirst = 0 (A)
    shrfmla.extend_from_slice(&1u16.to_le_bytes()); // colLast = 1 (B)
    shrfmla.extend_from_slice(&0u16.to_le_bytes()); // cUse
    shrfmla.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
    shrfmla.extend_from_slice(&rgce);
    push_record(&mut sheet, RECORD_SHRFMLA, &shrfmla);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_shrfmla_ref8_no_cuse_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_shrfmla_ref8_no_cuse_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedRef8NoCuse", &sheet_stream, 1252)
}

fn build_shared_formula_shrfmla_ref8_no_cuse_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 2) cols [0, 3) => A1:C2.
    // The shared formula range itself is only `A1:B2`; we include `C1` as a value cell so the
    // sheet is non-empty without overlapping the formula range.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // C1: a numeric cell so calamine sees a non-empty sheet value range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 2, xf_cell, 0.0));

    // Shared formula body stored in SHRFMLA: `"X"` (PtgStr).
    let rgce: [u8; 4] = [0x17, 1, 0, b'X']; // [PtgStr][cch=1][flags=0][X]

    // SHRFMLA definition for range A1:B2 using a Ref8 range header, omitting `cUse`:
    //   [rwFirst:u16][rwLast:u16][colFirst:u16][colLast:u16][cce:u16][rgce]
    let mut shrfmla = Vec::<u8>::new();
    shrfmla.extend_from_slice(&0u16.to_le_bytes()); // rwFirst = 0
    shrfmla.extend_from_slice(&1u16.to_le_bytes()); // rwLast = 1
    shrfmla.extend_from_slice(&0u16.to_le_bytes()); // colFirst = 0 (A)
    shrfmla.extend_from_slice(&1u16.to_le_bytes()); // colLast = 1 (B)
    shrfmla.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
    shrfmla.extend_from_slice(&rgce);
    push_record(&mut sheet, RECORD_SHRFMLA, &shrfmla);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_ptgexp_wide_payload_ptgref_relative_flags_workbook_stream() -> Vec<u8> {
    // Similar to `build_shared_formula_ptgexp_u32_row_u16_col_workbook_stream`, but the shared
    // SHRFMLA rgce uses `PtgRef` with row/col-relative flags. Without BIFF8 shared-formula
    // materialization, follower cells would decode unshifted references.
    let xf_cell = 16u16;
    let sheet_stream =
        build_shared_formula_ptgexp_wide_payload_ptgref_relative_flags_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedWideRefFlags", &sheet_stream, 1252)
}

fn build_shared_formula_ptgexp_wide_payload_ptgref_relative_flags_sheet_stream(
    xf_cell: u16,
) -> Vec<u8> {
    // Worksheet containing:
    // - Numeric inputs in A2/A3
    // - Shared formula range B2:B3 where SHRFMLA.rgce is `A2+1` encoded as
    //   PtgRef(row=1,col=0xC000)
    // - Follower cell B3 stores `PtgExp` coordinates as row u32 + col u16 (6 bytes)
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [1, 3) cols [0, 2) => A2:B3.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&1u32.to_le_bytes()); // first row
    dims.extend_from_slice(&3u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Input values.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 10.0)); // A2
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(2, 0, xf_cell, 20.0)); // A3

    // Anchor cell B2 (row=1,col=1): store canonical PtgExp(B2) followed by SHRFMLA containing the
    // shared rgce.
    let anchor_ptgexp = ptg_exp(1, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &anchor_ptgexp),
    );

    // Shared formula rgce: `A2+1`, encoded using PtgRef with relative flags.
    let shared_rgce: Vec<u8> = vec![
        0x24, // PtgRef
        0x01, 0x00, // rw = 1 (row 2)
        0x00, 0xC0, // col = 0 | row_rel | col_rel
        0x1E, // PtgInt
        0x01, 0x00, // 1
        0x03, // PtgAdd
    ];

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(1, 2, 1, 1, &shared_rgce),
    );

    // Follower cell B3 (row=2,col=1): PtgExp with a 6-byte payload (row u32 + col u16).
    let follower_ptgexp = ptg_exp_row_u32_col_u16(1, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(2, 1, xf_cell, 0.0, &follower_ptgexp),
    );

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

fn build_shared_formula_continued_ptgstr_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_continued_ptgstr_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedStr", &sheet_stream, 1252)
}

fn build_shared_formula_continued_ptgstr_sheet_stream(xf_cell: u16) -> Vec<u8> {
    const FORMULA_GRBIT_F_SHR_FMLA: u16 = 0x0008;

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

    // A1: a numeric cell so the sheet is non-empty (calamine value range).
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 0.0));

    // Shared formula body stored in SHRFMLA: `"ABCDE"` (PtgStr).
    let literal = "ABCDE";
    let mut shared_rgce = Vec::new();
    shared_rgce.push(0x17); // PtgStr
    shared_rgce.push(literal.len() as u8); // cch
    shared_rgce.push(0); // flags (compressed)
    shared_rgce.extend_from_slice(literal.as_bytes());

    // Base cell B1: `PtgExp` referencing itself (row=0, col=1).
    let ptgexp_b1 = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(0, 1, xf_cell, 0.0, FORMULA_GRBIT_F_SHR_FMLA, &ptgexp_b1),
    );

    // SHRFMLA definition for the range B1:B2, with the PtgStr character bytes split across a
    // CONTINUE boundary. The continued fragment begins with the required 1-byte "continued
    // segment" option flags prefix (`fHighByte`).
    let cce = shared_rgce.len() as u16;
    let split_at = 3 + 2; // ptg + cch + flags + "AB"
    let first_rgce = &shared_rgce[..split_at];
    let remaining_rgce = &shared_rgce[split_at..]; // "CDE"

    let mut shrfmla_part1 = Vec::new();
    shrfmla_part1.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
    shrfmla_part1.extend_from_slice(&1u16.to_le_bytes()); // rwLast
    shrfmla_part1.push(1u8); // colFirst (B)
    shrfmla_part1.push(1u8); // colLast (B)
    shrfmla_part1.extend_from_slice(&0u16.to_le_bytes()); // cUse
    shrfmla_part1.extend_from_slice(&cce.to_le_bytes());
    shrfmla_part1.extend_from_slice(first_rgce);
    push_record(&mut sheet, RECORD_SHRFMLA, &shrfmla_part1);

    let mut cont = Vec::new();
    cont.push(0); // continued segment option flags (compressed)
    cont.extend_from_slice(remaining_rgce);
    push_record(&mut sheet, RECORD_CONTINUE, &cont);

    // Follower cell B2: `PtgExp` referencing base cell B1.
    let ptgexp_b2 = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, 1, xf_cell, 0.0, FORMULA_GRBIT_F_SHR_FMLA, &ptgexp_b2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_ptgarray_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgarray_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedArray", &sheet_stream, 1252)
}

fn build_shared_formula_ptgarray_wide_ptgexp_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgarray_wide_ptgexp_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedArrayWide", &sheet_stream, 1252)
}

fn build_shared_formula_ptgarray_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula range B1:B2.
    //
    // Both cells contain PtgExp, and the shared formula body is stored in the trailing SHRFMLA
    // record along with the `rgcb` payload required to decode `PtgArray` constants.
    //
    // Expected decoded formulas:
    // - B1: `A1+SUM({1,2;3,4})`
    // - B2: `A2+SUM({1,2;3,4})`
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
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));

    // Set FORMULA.grbit.fShrFmla (0x0008) so parsers recognize the shared-formula membership.
    let grbit_shared: u16 = 0x0008;

    // B1 formula: PtgExp pointing to itself (rw=0,col=1).
    let ptgexp = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(0, 1, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    // Shared formula body stored in SHRFMLA.
    let rgce_shared = {
        let mut v = Vec::new();
        // PtgRefN: row_off=0, col_off=-1 relative to the formula cell.
        v.push(0x2C);
        v.extend_from_slice(&0u16.to_le_bytes()); // row_off = 0
        v.extend_from_slice(&0xFFFFu16.to_le_bytes()); // col_off = -1 (14-bit), row+col relative

        // PtgArray (array constant; data stored in rgcb).
        v.push(0x20);
        v.extend_from_slice(&[0u8; 7]); // reserved

        // PtgFuncVar: SUM(argc=1).
        v.push(0x22);
        v.push(1);
        v.extend_from_slice(&4u16.to_le_bytes());

        // PtgAdd.
        v.push(0x03);
        v
    };

    // rgcb payload for the array constant `{1,2;3,4}`.
    let rgcb = rgcb_array_constant_numbers_2x2(&[1.0, 2.0, 3.0, 4.0]);

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record_with_rgcb(0, 1, 1, 1, &rgce_shared, &rgcb),
    );

    // B2 formula: PtgExp pointing to base cell B1 (rw=0,col=1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, 1, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_ptgarray_wide_ptgexp_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula range B1:B2.
    //
    // Like `build_shared_formula_ptgarray_sheet_stream`, but the follower cell `B2` stores a
    // non-standard wide `PtgExp` payload width (row u32 + col u16, `cce=7`).
    //
    // Expected decoded formulas:
    // - B1: `A1+SUM({1,2;3,4})`
    // - B2: `A2+SUM({1,2;3,4})`
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
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));

    // Set FORMULA.grbit.fShrFmla (0x0008) so parsers recognize the shared-formula membership.
    let grbit_shared: u16 = 0x0008;

    // B1 formula: canonical PtgExp pointing to itself (rw=0,col=1).
    let base_ptgexp = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(0, 1, xf_cell, 0.0, grbit_shared, &base_ptgexp),
    );

    // Shared formula body stored in SHRFMLA: `A(row)+SUM({1,2;3,4})`.
    let rgce_shared = {
        let mut v = Vec::new();
        // PtgRefN: row_off=0, col_off=-1 relative to the formula cell.
        v.push(0x2C);
        v.extend_from_slice(&0u16.to_le_bytes()); // row_off = 0
        v.extend_from_slice(&0xFFFFu16.to_le_bytes()); // col_off = -1 (14-bit), row+col relative

        // PtgArray (array constant; data stored in rgcb).
        v.push(0x20);
        v.extend_from_slice(&[0u8; 7]); // reserved

        // PtgFuncVar: SUM(argc=1).
        v.push(0x22);
        v.push(1);
        v.extend_from_slice(&4u16.to_le_bytes());

        // PtgAdd.
        v.push(0x03);
        v
    };

    // rgcb payload for the array constant `{1,2;3,4}`.
    let rgcb = rgcb_array_constant_numbers_2x2(&[1.0, 2.0, 3.0, 4.0]);
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record_with_rgcb(0, 1, 1, 1, &rgce_shared, &rgcb),
    );

    // B2 formula: wide-payload PtgExp pointing to base cell B1 (rw=0,col=1).
    let follower_ptgexp = ptg_exp_row_u32_col_u16(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, 1, xf_cell, 0.0, grbit_shared, &follower_ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_formula_array_constant_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_formula_array_constant_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ArrayConst", &sheet_stream, 1252)
}

fn build_formula_array_constant_continued_rgcb_string_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_formula_array_constant_continued_rgcb_string_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ArrayConstStr", &sheet_stream, 1252)
}

fn build_formula_array_constant_continued_rgcb_string_compressed_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream =
        build_formula_array_constant_continued_rgcb_string_compressed_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ArrayConstStrCompressed", &sheet_stream, 1252)
}

fn build_formula_array_constant_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Formula: SUM({1,2;3,4})
    let rgce = [
        0x20u8, // PtgArray
        0, 0, 0, 0, 0, 0, 0,    // 7-byte header
        0x22, // PtgFuncVar
        0x01, // argc=1
        0x04, 0x00, // iftab=4 (SUM)
    ];

    let rgcb = rgcb_array_constant_numbers_2x2(&[1.0, 2.0, 3.0, 4.0]);

    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_rgcb(0, 0, xf_cell, 10.0, &rgce, &rgcb),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_formula_array_constant_continued_rgcb_string_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Formula: SUM({\"ABCDE\"})
    let rgce = [
        0x20u8, // PtgArray
        0, 0, 0, 0, 0, 0, 0, // 7-byte header
        0x22, // PtgFuncVar
        0x01, // argc=1
        0x04, 0x00, // iftab=4 (SUM)
    ];

    let rgcb = rgcb_array_constant_string_1x1("ABCDE");
    let full_payload = formula_cell_with_rgcb(0, 0, xf_cell, 0.0, &rgce, &rgcb);

    // Split the rgcb string's UTF-16 bytes across a CONTINUE boundary after 2 characters, and
    // prefix the continued fragment with the required 1-byte option flags (fHighByte=1).
    let formula_header_len = 22usize;
    let rgce_len = rgce.len();
    let rgcb_string_prefix_len = 4 /* dims */ + 1 /* ty */ + 2 /* cch */;
    let start_utf16 = formula_header_len + rgce_len + rgcb_string_prefix_len;
    let split_at = start_utf16 + 4; // two UTF-16 code units ("AB")
    let (part1, rest) = full_payload.split_at(split_at);

    push_record(&mut sheet, RECORD_FORMULA, part1);

    let mut cont = Vec::new();
    cont.push(1); // continued segment option flags (fHighByte=1)
    cont.extend_from_slice(rest);
    push_record(&mut sheet, RECORD_CONTINUE, &cont);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_formula_array_constant_continued_rgcb_string_compressed_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Formula: SUM({\"ABCDE\"})
    let rgce = [
        0x20u8, // PtgArray
        0, 0, 0, 0, 0, 0, 0, // 7-byte header
        0x22, // PtgFuncVar
        0x01, // argc=1
        0x04, 0x00, // iftab=4 (SUM)
    ];

    let rgcb = rgcb_array_constant_string_1x1("ABCDE");
    let full_payload = formula_cell_with_rgcb(0, 0, xf_cell, 0.0, &rgce, &rgcb);

    // Split the rgcb string's UTF-16 bytes across a CONTINUE boundary after 2 characters. The
    // continued fragment is stored in compressed form (fHighByte=0), so we emit single-byte
    // characters and rely on the importer to expand them back to canonical UTF-16LE.
    let formula_header_len = 22usize;
    let rgce_len = rgce.len();
    let rgcb_string_prefix_len = 4 /* dims */ + 1 /* ty */ + 2 /* cch */;
    let start_utf16 = formula_header_len + rgce_len + rgcb_string_prefix_len;
    let split_at = start_utf16 + 4; // two UTF-16 code units ("AB")
    let (part1, rest_utf16) = full_payload.split_at(split_at);

    push_record(&mut sheet, RECORD_FORMULA, part1);

    let mut rest_compressed = Vec::with_capacity(rest_utf16.len() / 2);
    assert!(
        rest_utf16.len() % 2 == 0,
        "expected UTF-16LE payload to have even length"
    );
    for chunk in rest_utf16.chunks_exact(2) {
        assert_eq!(chunk[1], 0, "expected ASCII string for compressed segment");
        rest_compressed.push(chunk[0]);
    }

    let mut cont = Vec::new();
    cont.push(0); // continued segment option flags (fHighByte=0 => compressed)
    cont.extend_from_slice(&rest_compressed);
    push_record(&mut sheet, RECORD_CONTINUE, &cont);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_area3d_workbook_stream() -> Vec<u8> {
    // Two-sheet workbook:
    // - Sheet1: a simple value sheet used as the 3D reference target.
    // - SharedArea3D: contains a shared-formula range B1:B2 whose shared rgce uses PtgArea3d with
    //   relative flags in the column fields.
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
    for name in ["Sheet1", "SharedArea3D"] {
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
    // EXTERNSHEET entry for Sheet1 (itab=0).
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0 (Sheet1) ---------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 1.0));

    // -- Sheet 1 (SharedArea3D) ---------------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_area3d_shared_formula_sheet_stream(xf_cell));

    globals
}

fn build_array_formula_workbook_stream() -> Vec<u8> {
    // Use the generic single-sheet workbook builder: it creates a minimal BIFF8 globals stream
    // including a default cell XF at index 16.
    let xf_cell = 16u16;
    let sheet_stream = build_array_formula_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("Array", &sheet_stream, 1252)
}

fn build_array_formula_ptgarray_workbook_stream() -> Vec<u8> {
    // Minimal single-sheet workbook containing an array formula (`ARRAY` + `PtgExp`) whose ARRAY
    // record includes trailing `rgcb` data for a `PtgArray` constant.
    let xf_cell = 16u16;
    let sheet_stream = build_array_formula_ptgarray_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ArrayConst", &sheet_stream, 1252)
}

fn build_array_formula_range_flags_ambiguity_workbook_stream() -> Vec<u8> {
    // Minimal single-sheet workbook containing multiple ARRAY records whose range headers can be
    // ambiguous when parsed as Ref8.
    let xf_cell = 16u16;
    let sheet_stream = build_array_formula_range_flags_ambiguity_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ArrayRangeAmbiguity", &sheet_stream, 1252)
}

fn build_array_formula_ptgname_workbook_stream() -> Vec<u8> {
    // Minimal single-sheet workbook containing an array formula whose ARRAY.rgce includes PtgName.
    //
    // We intentionally construct a workbook NAME table so BIFF8 formula decoding can resolve the
    // PtgName index into the defined name string.
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

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "ArrayName");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // External reference tables required by the NAME formula token stream (PtgRef3d).
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(1));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // One workbook-scoped defined name: MyName -> ArrayName!$A$1.
    let rgce = ptg_ref3d(0, 0, 0);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("MyName", 0, false, None, &rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_array_formula_ptgname_sheet_stream(xf_cell);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_array_formula_ptgexp_missing_array_workbook_stream() -> Vec<u8> {
    // Minimal single-sheet workbook containing a malformed array formula group where the base cell
    // stores a full `FORMULA.rgce` but a follower cell uses `PtgExp` + fArray without an `ARRAY`
    // definition record.
    let xf_cell = 16u16;
    let sheet_stream = build_array_formula_ptgexp_missing_array_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("ArrayMissing", &sheet_stream, 1252)
}

fn build_shared_formula_ptgmemarean_workbook_stream() -> Vec<u8> {
    // Minimal single-sheet workbook containing a shared formula where the shared SHRFMLA.rgce
    // includes PtgMemAreaN tokens (one with cce=0 and one with cce=3).
    let xf_cell = 16u16;
    let sheet = build_shared_formula_ptgmemarean_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("Shared", &sheet, 1252)
}

fn build_shared_formula_area3d_mixed_flags_workbook_stream() -> Vec<u8> {
    // Two-sheet workbook:
    // - Sheet1: a simple value sheet used as the 3D reference target.
    // - SharedArea3D: contains a shared-formula range B1:B2 whose shared rgce uses PtgArea3d with
    //   different relative flags per endpoint.
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
    for name in ["Sheet1", "SharedArea3D"] {
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
    // EXTERNSHEET entry for Sheet1 (itab=0).
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0 (Sheet1) ---------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 1.0));

    // -- Sheet 1 (SharedArea3D) ---------------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals
        .extend_from_slice(&build_shared_area3d_shared_formula_mixed_flags_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_ptgref_relative_flags_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgref_relative_flags_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("Sheet1", &sheet_stream, 1252)
}

fn build_shared_formula_ptgref_relative_flags_sheet_stream(xf_cell: u16) -> Vec<u8> {
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

    // Shared-formula reference: `PtgExp` to base cell B1.
    let rgce_ptgexp = vec![0x01, 0x00, 0x00, 0x01, 0x00];

    // B1 (base) and B2 (follower): both store a PtgExp token stream.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &rgce_ptgexp),
    );
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &rgce_ptgexp),
    );

    // Shared formula definition (`SHRFMLA`) for range B1:B2.
    //
    // Base rgce is `A1+1`, encoded using `PtgRef` with row/col-relative flags set.
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
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &rgce_base),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_ptgarea_relative_flags_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgarea_relative_flags_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedArea2D", &sheet_stream, 1252)
}

fn build_shared_formula_ptgarea_relative_flags_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 3) cols [0, 2) => A1:B3.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&3u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide numeric inputs in A1:A3 so the area references are within the sheet's used range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(2, 0, xf_cell, 3.0));

    // Shared-formula reference: `PtgExp` to base cell B1.
    let rgce_ptgexp = vec![0x01, 0x00, 0x00, 0x01, 0x00];

    // B1 (base) and B2 (follower): both store a PtgExp token stream.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &rgce_ptgexp),
    );
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &rgce_ptgexp),
    );

    // Shared formula definition (`SHRFMLA`) for range B1:B2.
    //
    // Base rgce is `SUM(A1:A2)`, encoded using `PtgArea` with row/col-relative flags set.
    let rgce_base = vec![
        0x25, // PtgArea
        0x00, 0x00, // rwFirst = 0
        0x01, 0x00, // rwLast = 1
        0x00, 0xC0, // colFirst = 0 | row_rel | col_rel
        0x00, 0xC0, // colLast = 0 | row_rel | col_rel
        0x22, // PtgFuncVar
        0x01, // argc = 1
        0x04, 0x00, // SUM
    ];

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &rgce_base),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_ptgname_workbook_stream() -> Vec<u8> {
    // Minimal single-sheet workbook containing a shared formula whose shared rgce includes PtgName.
    //
    // We intentionally construct a workbook NAME table so BIFF8 formula decoding can resolve the
    // PtgName index into the defined name string.
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

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "SharedName");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // External reference tables required by the NAME formula token stream (PtgRef3d).
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(1));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    // One workbook-scoped defined name: MyName -> SharedName!$A$1.
    let rgce = ptg_ref3d(0, 0, 0);
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record("MyName", 0, false, None, &rgce),
    );

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_shared_formula_ptgname_sheet_stream(xf_cell);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_shared_formula_ptgname_sheet_stream(xf_cell: u16) -> Vec<u8> {
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

    // Provide numeric cells in A1/A2 so the sheet is not empty.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // Base cell B1 formula: `MyName` (PtgName index 1).
    let base_rgce = ptg_name(1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &base_rgce),
    );

    // Shared formula rgce stored in SHRFMLA: `MyName`.
    let shared_rgce = ptg_name(1);
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

fn build_shared_formula_shrfmla_only_ptgref_relative_flags_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet_stream =
        build_shared_formula_shrfmla_only_ptgref_relative_flags_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedOnlyRefFlags", &sheet_stream, 1252)
}

fn build_shared_formula_shrfmla_only_ptgarea_relative_flags_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet_stream =
        build_shared_formula_shrfmla_only_ptgarea_relative_flags_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedOnlyAreaFlags", &sheet_stream, 1252)
}

fn build_shared_formula_shrfmla_only_ptgref_relative_flags_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula range B1:B2, but without any cell `FORMULA` records. The importer should still
    // recover formulas from the SHRFMLA record definition.
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

    // Provide value cells in column A so the references are within the sheet's used range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0)); // A1
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0)); // A2

    // Shared formula definition (`SHRFMLA`) for range B1:B2.
    //
    // Base rgce is `A1+1`, encoded using `PtgRef` with row/col-relative flags set (rather than
    // `PtgRefN` offsets).
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
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &rgce_base),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_out_of_bounds_relative_refs_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_out_of_bounds_relative_refs_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedOutOfBounds", &sheet_stream, 1252)
}

fn build_shared_formula_shrfmla_only_ptgarea_relative_flags_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula range B1:B2, but without any cell `FORMULA` records. The importer should still
    // recover formulas from the SHRFMLA record definition.
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 3) cols [0, 2) => A1:B3.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&3u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide numeric inputs in A1:A3 so the area references are within the sheet's used range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(2, 0, xf_cell, 3.0));

    // Shared formula definition (`SHRFMLA`) for range B1:B2.
    //
    // Base rgce is `SUM(A1:A2)`, encoded using `PtgArea` with row/col-relative flags set.
    let rgce_base = vec![
        0x25, // PtgArea
        0x00, 0x00, // rwFirst = 0
        0x01, 0x00, // rwLast = 1
        0x00, 0xC0, // colFirst = 0 | row_rel | col_rel
        0x00, 0xC0, // colLast = 0 | row_rel | col_rel
        0x22, // PtgFuncVar
        0x01, // argc = 1
        0x04, 0x00, // SUM
    ];

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &rgce_base),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
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

fn build_row_col_style_out_of_range_xf_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes()); // CODEPAGE: Windows-1252
    push_record(&mut globals, RECORD_WINDOW1, &window1()); // WINDOW1
    push_record(&mut globals, RECORD_FONT, &font("Arial")); // FONT

    // XF table. Many readers expect at least 16 style XFs before cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }

    // One General cell XF for the lone NUMBER record in the sheet stream.
    let xf_cell_general = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Row/column default formatting references out-of-range indices.
    let xf_row_oob = 5000u16;
    let xf_col_oob = 6000u16;

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "RowColStylesOutOfRange");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_row_col_style_sheet_stream(xf_cell_general, xf_row_oob, xf_col_oob);

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

    // Add a malformed NAME record whose `cce` claims far more bytes than the physical payload.
    //
    // This should be repaired by `sanitize_biff8_continued_name_records_for_calamine()` (it patches
    // `cce` to 0) and will trigger calamine panics in older/newer versions that still assume `cce`
    // always fits in the record payload.
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record_malformed_claimed_cce("BadCce", 0xFFFF),
    );

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

fn build_shared_formula_external_refs_workbook_stream() -> Vec<u8> {
    // This workbook contains:
    // - One internal sheet (`Shared`)
    // - SUPBOOK[0]: internal workbook marker
    // - SUPBOOK[1]: external workbook `Book1.xlsx` with sheet `ExtSheet`
    // - EXTERNNAME entries for `PtgNameX` references
    // - EXTERNSHEET[0] pointing at SUPBOOK[1] / ExtSheet
    // - Worksheet `Shared` containing two shared formula ranges:
    //   - B1:B2: `'[Book1.xlsx]ExtSheet'!$A$1+1` via `PtgRef3d`
    //   - C1:C2: `'[Book1.xlsx]ExtSheet'!ExtDefined+1` via `PtgNameX`
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

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Shared");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // External reference tables.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(1)); // internal workbook marker
    push_record(
        &mut globals,
        RECORD_SUPBOOK,
        &supbook_external("Book1.xlsx", &["ExtSheet"]),
    );
    // External name #1 for PtgNameX.
    push_record(
        &mut globals,
        RECORD_EXTERNNAME,
        &externname_record("ExtDefined"),
    );
    // EXTERNSHEET entry 0 => [Book1.xlsx]ExtSheet
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record_with_supbook(&[(1, 0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_shared_formula_external_refs_sheet_stream(xf_cell);

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);

    globals
}

fn build_array_formula_external_refs_workbook_stream() -> Vec<u8> {
    // This workbook contains:
    // - One internal sheet (`ArrayExt`)
    // - SUPBOOK[0]: internal workbook marker
    // - SUPBOOK[1]: external workbook `Book1.xlsx` with sheet `ExtSheet`
    // - EXTERNNAME entries for `PtgNameX` references
    // - EXTERNSHEET[0] pointing at SUPBOOK[1] / ExtSheet
    // - Worksheet `ArrayExt` containing two array-formula ranges:
    //   - B1:B2: `'[Book1.xlsx]ExtSheet'!$A$1+1` via `PtgRef3d`
    //   - C1:C2: `'[Book1.xlsx]ExtSheet'!ExtDefined+1` via `PtgNameX`
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

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "ArrayExt");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // External reference tables.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(1)); // internal workbook marker
    push_record(
        &mut globals,
        RECORD_SUPBOOK,
        &supbook_external("Book1.xlsx", &["ExtSheet"]),
    );
    push_record(
        &mut globals,
        RECORD_EXTERNNAME,
        &externname_record("ExtDefined"),
    );
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record_with_supbook(&[(1, 0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_array_formula_external_refs_sheet_stream(xf_cell);

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

fn build_shared_formula_ref3d_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records (workbook sheet list).
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["Sheet1", "Shared3D"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    // Use a single internal SUPBOOK so ixti=0 refers to Sheet1.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0: Sheet1 ---------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 1.0));

    // -- Sheet 1: Shared3D -------------------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_ref3d_shared_formula_sheet_stream(xf_cell));

    globals
}

fn build_shared_ref3d_shared_formula_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [0, 3) => A1:C2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Shared formula base cell is B1 (row=0,col=1).
    let base_row: u16 = 0;
    let base_col: u16 = 1;
    // FORMULA.grbit bit indicating the formula is part of a shared-formula group.
    // [MS-XLS] 2.4.127 (FORMULA), fShrFmla.
    let grbit_shared: u16 = 0x0008;

    // Base cell FORMULA record stores the full rgce.
    //
    // The shared formula definition is also stored in the following SHRFMLA record so that
    // follower cells (PtgExp) can be materialized.
    let mut base_formula = Vec::<u8>::new();
    base_formula.extend_from_slice(&ptg_ref3d(0, 0, 0xC000)); // ixti=0 => Sheet1, A1, relative row/col
    base_formula.push(0x1E); // PtgInt
    base_formula.extend_from_slice(&1u16.to_le_bytes());
    base_formula.push(0x03); // PtgAdd
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(
            base_row,
            base_col,
            xf_cell,
            0.0,
            grbit_shared,
            &base_formula,
        ),
    );

    // SHRFMLA: shared formula over B1:C2.
    // Base rgce: Sheet1!A1 + 1, where the PtgRef3d token sets both row+col relative flags.
    let shrfmla = shrfmla_record(0, 1, 1, 2, &base_formula);
    push_record(&mut sheet, RECORD_SHRFMLA, &shrfmla);

    // Non-base cells use PtgExp pointing at the base cell.
    let follower = ptg_exp(base_row, base_col);
    // C1
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(0, 2, xf_cell, 0.0, grbit_shared, &follower),
    );
    // B2
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, 1, xf_cell, 0.0, grbit_shared, &follower),
    );
    // C2
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, 2, xf_cell, 0.0, grbit_shared, &follower),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_3d_oob_workbook_stream() -> Vec<u8> {
    // Workbook with:
    // - Sheet1: A65536 contains a number (bottom BIFF8 row).
    // - Shared3D_OOB: shared formula at the BIFF8 row limit that materializes an out-of-bounds 3D
    //   reference to `#REF!` in the follower cell.
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records (workbook sheet list).
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["Sheet1", "Shared3D_OOB"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    // Use a single internal SUPBOOK so ixti=0 refers to Sheet1.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0: Sheet1 ---------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_sheet1_bottom_number_sheet_stream(xf_cell));

    // -- Sheet 1: Shared3D_OOB ----------------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_ref3d_oob_shared_formula_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_3d_oob_shrfmla_only_workbook_stream() -> Vec<u8> {
    // Workbook with:
    // - Sheet1: A65536 contains a number (bottom BIFF8 row).
    // - Shared3D_OOB_ShrFmlaOnly: shared formula definition stored only in SHRFMLA (no FORMULA
    //   records). The importer must materialize the shared 3D reference per cell and convert the
    //   out-of-bounds shifted reference to `#REF!` (PtgRefErr3d).
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records (workbook sheet list).
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["Sheet1", "Shared3D_OOB_ShrFmlaOnly"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    // Use a single internal SUPBOOK so ixti=0 refers to Sheet1.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0: Sheet1 ---------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_sheet1_bottom_number_sheet_stream(xf_cell));

    // -- Sheet 1: Shared3D_OOB_ShrFmlaOnly ---------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_ref3d_oob_shrfmla_only_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_area3d_oob_workbook_stream() -> Vec<u8> {
    // Workbook with:
    // - Sheet1: A65535:A65536 contains numbers (near BIFF8 row limit).
    // - SharedArea3D_OOB: shared formula at the BIFF8 row limit that materializes an out-of-bounds
    //   3D area reference to `#REF!` in the follower cell.
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records (workbook sheet list).
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["Sheet1", "SharedArea3D_OOB"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    // Use a single internal SUPBOOK so ixti=0 refers to Sheet1.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0: Sheet1 ---------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_sheet1_bottom_two_number_sheet_stream(xf_cell));

    // -- Sheet 1: SharedArea3D_OOB -----------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_area3d_oob_shared_formula_sheet_stream(
        xf_cell,
    ));

    globals
}

fn build_shared_formula_area3d_oob_shrfmla_only_workbook_stream() -> Vec<u8> {
    // Workbook with:
    // - Sheet1: A65535:A65536 contains numbers (near BIFF8 row limit).
    // - SharedArea3D_OOB_ShrFmlaOnly: shared formula definition stored only in SHRFMLA (no FORMULA
    //   records). The importer must materialize the 3D area reference per cell and convert the
    //   out-of-bounds shifted reference to `#REF!` (PtgAreaErr3d).
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records (workbook sheet list).
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["Sheet1", "SharedArea3D_OOB_ShrFmlaOnly"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    // Use a single internal SUPBOOK so ixti=0 refers to Sheet1.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0: Sheet1 ---------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_sheet1_bottom_two_number_sheet_stream(xf_cell));

    // -- Sheet 1: SharedArea3D_OOB_ShrFmlaOnly -----------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_area3d_oob_shrfmla_only_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_3d_col_oob_shrfmla_only_workbook_stream() -> Vec<u8> {
    // Workbook with:
    // - Sheet1: minimal value sheet (target of 3D reference).
    // - Shared3D_ColOOB_ShrFmlaOnly: shared formula near the max BIFF8 column that materializes an
    //   out-of-bounds 3D reference in the follower cell.
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records (workbook sheet list).
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["Sheet1", "Shared3D_ColOOB_ShrFmlaOnly"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    // Use a single internal SUPBOOK so ixti=0 refers to Sheet1.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0: Sheet1 ---------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 1.0));

    // -- Sheet 1: Shared3D_ColOOB_ShrFmlaOnly ------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_ref3d_col_oob_shrfmla_only_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_area3d_col_oob_shrfmla_only_workbook_stream() -> Vec<u8> {
    // Workbook with:
    // - Sheet1: minimal value sheet (target of 3D reference).
    // - SharedArea3D_ColOOB_ShrFmlaOnly: shared 3D area reference near the max BIFF8 column that
    //   shifts out of bounds in the follower cell.
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table: 16 style XFs + one cell XF.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // BoundSheet records (workbook sheet list).
    let mut boundsheet_offset_positions: Vec<usize> = Vec::new();
    for name in ["Sheet1", "SharedArea3D_ColOOB_ShrFmlaOnly"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // External reference tables used by 3D formula tokens.
    // Use a single internal SUPBOOK so ixti=0 refers to Sheet1.
    push_record(&mut globals, RECORD_SUPBOOK, &supbook_internal(2));
    push_record(
        &mut globals,
        RECORD_EXTERNSHEET,
        &externsheet_record(&[(0, 0)]),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0: Sheet1 ---------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 1.0));

    // -- Sheet 1: SharedArea3D_ColOOB_ShrFmlaOnly --------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_area3d_col_oob_shrfmla_only_sheet_stream(xf_cell));

    globals
}

fn build_shared_ref3d_col_oob_shrfmla_only_sheet_stream(_xf_cell: u16) -> Vec<u8> {
    // Shared formula definition stored only in SHRFMLA for range XFC1:XFD1.
    //
    // Materialized formulas:
    // - XFC1: Sheet1!XFD1+1
    // - XFD1: #REF!+1 (because Sheet1!XFE1 is out of bounds)
    const ROW: u16 = 0;
    const COL_FIRST: u16 = 0x3FFE; // XFC
    const COL_LAST: u16 = 0x3FFF; // XFD

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [XFC, XFD+1).
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&COL_FIRST.to_le_bytes()); // first col
    dims.extend_from_slice(&(COL_LAST + 1).to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Shared formula rgce: Sheet1!XFD1 + 1, where PtgRef3d carries row+col relative flags so
    // filling right shifts the column.
    let col_with_flags = pack_biff8_col_flags(COL_LAST, true, true); // XFD + rowRel + colRel
    let mut shared_rgce = Vec::<u8>::new();
    shared_rgce.extend_from_slice(&ptg_ref3d(0, ROW, col_with_flags));
    shared_rgce.push(0x1E); // PtgInt
    shared_rgce.extend_from_slice(&1u16.to_le_bytes());
    shared_rgce.push(0x03); // PtgAdd

    // SHRFMLA record defining shared rgce for range XFC1:XFD1 using a Ref8 header (u16 cols).
    let mut shrfmla = Vec::<u8>::new();
    shrfmla.extend_from_slice(&ROW.to_le_bytes()); // rwFirst
    shrfmla.extend_from_slice(&ROW.to_le_bytes()); // rwLast
    shrfmla.extend_from_slice(&COL_FIRST.to_le_bytes()); // colFirst (u16)
    shrfmla.extend_from_slice(&COL_LAST.to_le_bytes()); // colLast (u16)
    shrfmla.extend_from_slice(&(shared_rgce.len() as u16).to_le_bytes()); // cce
    shrfmla.extend_from_slice(&shared_rgce);
    push_record(&mut sheet, RECORD_SHRFMLA, &shrfmla);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_area3d_col_oob_shrfmla_only_sheet_stream(_xf_cell: u16) -> Vec<u8> {
    // Shared formula definition stored only in SHRFMLA for range XFC1:XFD1.
    //
    // Materialized formulas:
    // - XFC1: Sheet1!XFC1:XFD1+1
    // - XFD1: #REF!+1 (because Sheet1!XFD1:XFE1 is out of bounds)
    const ROW: u16 = 0;
    const COL_FIRST: u16 = 0x3FFE; // XFC
    const COL_LAST: u16 = 0x3FFF; // XFD

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [XFC, XFD+1).
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&COL_FIRST.to_le_bytes()); // first col
    dims.extend_from_slice(&(COL_LAST + 1).to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Shared formula rgce: Sheet1!XFC1:XFD1 + 1, where PtgArea3d carries row+col relative flags on
    // both endpoints so filling right shifts both columns.
    let col_first = pack_biff8_col_flags(COL_FIRST, true, true);
    let col_last = pack_biff8_col_flags(COL_LAST, true, true);
    let mut shared_rgce = Vec::<u8>::new();
    shared_rgce.extend_from_slice(&ptg_area3d(0, ROW, ROW, col_first, col_last));
    shared_rgce.push(0x1E); // PtgInt
    shared_rgce.extend_from_slice(&1u16.to_le_bytes());
    shared_rgce.push(0x03); // PtgAdd

    // SHRFMLA record defining shared rgce for range XFC1:XFD1 using a Ref8 header (u16 cols).
    let mut shrfmla = Vec::<u8>::new();
    shrfmla.extend_from_slice(&ROW.to_le_bytes()); // rwFirst
    shrfmla.extend_from_slice(&ROW.to_le_bytes()); // rwLast
    shrfmla.extend_from_slice(&COL_FIRST.to_le_bytes()); // colFirst (u16)
    shrfmla.extend_from_slice(&COL_LAST.to_le_bytes()); // colLast (u16)
    shrfmla.extend_from_slice(&(shared_rgce.len() as u16).to_le_bytes()); // cce
    shrfmla.extend_from_slice(&shared_rgce);
    push_record(&mut sheet, RECORD_SHRFMLA, &shrfmla);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_sheet1_bottom_number_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Sheet1 contains a single cell at the bottom BIFF8 row: A65536 (row index 65535).
    const ROW: u16 = u16::MAX;
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [65535, 65536) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A65536: NUMBER record.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(ROW, 0, xf_cell, 1.0),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_sheet1_bottom_two_number_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Sheet1 contains two cells near the bottom BIFF8 rows: A65535 and A65536.
    const ROW1: u16 = u16::MAX - 1; // 65534 => row 65535 (1-based)
    const ROW2: u16 = u16::MAX; // 65535 => row 65536 (1-based)
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [65534, 65536) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(ROW1 as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(ROW2 as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(ROW1, 0, xf_cell, 1.0),
    ); // A65535
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(ROW2, 0, xf_cell, 2.0),
    ); // A65536

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_ref3d_oob_shared_formula_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula in B65535:B65536:
    // - B65535: Sheet1!A65536+1
    // - B65536: #REF!+1 (because Sheet1!A65537 is out of BIFF8 bounds)
    const BASE_ROW: u16 = u16::MAX - 1; // 65534 => row 65535 (1-based)
    const FOLLOW_ROW: u16 = u16::MAX; // 65535 => row 65536 (1-based)
    const COL_B: u16 = 1;
    let grbit_shared: u16 = 0x0008;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [65534, 65536) cols [1, 2) => B65535:B65536.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(BASE_ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(FOLLOW_ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&COL_B.to_le_bytes()); // first col (B)
    dims.extend_from_slice(&(COL_B + 1).to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Cells store PtgExp pointing at the base cell; the shared rgce is stored in SHRFMLA.
    let ptgexp = ptg_exp(BASE_ROW, COL_B);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(BASE_ROW, COL_B, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    // Shared formula rgce: Sheet1!A65536 + 1, where PtgRef3d carries relative flags so filling down
    // shifts the row.
    let mut shared_rgce = Vec::<u8>::new();
    shared_rgce.extend_from_slice(&ptg_ref3d(0, u16::MAX, 0xC000)); // ixti=0 => Sheet1, A65536, row+col relative
    shared_rgce.push(0x1E); // PtgInt
    shared_rgce.extend_from_slice(&1u16.to_le_bytes());
    shared_rgce.push(0x03); // PtgAdd

    // SHRFMLA record defining shared rgce for range B65535:B65536.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(BASE_ROW, FOLLOW_ROW, COL_B as u8, COL_B as u8, &shared_rgce),
    );

    // Follower B65536: PtgExp referencing base cell B65535.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(FOLLOW_ROW, COL_B, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_ref3d_oob_shrfmla_only_sheet_stream(_xf_cell: u16) -> Vec<u8> {
    // Shared formula definition in SHRFMLA for range B65535:B65536, with no cell-level FORMULA
    // records. The importer must expand/materialize the shared rgce itself.
    //
    // Materialized formulas should match `build_shared_ref3d_oob_shared_formula_sheet_stream`:
    // - B65535: Sheet1!A65536+1
    // - B65536: #REF!+1
    const BASE_ROW: u16 = u16::MAX - 1; // 65534 => row 65535 (1-based)
    const FOLLOW_ROW: u16 = u16::MAX; // 65535 => row 65536 (1-based)
    const COL_B: u16 = 1;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [65534, 65536) cols [1, 2) => B65535:B65536.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(BASE_ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(FOLLOW_ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&COL_B.to_le_bytes()); // first col (B)
    dims.extend_from_slice(&(COL_B + 1).to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Shared formula rgce: Sheet1!A65536 + 1, where PtgRef3d carries relative flags so filling down
    // shifts the row.
    let mut shared_rgce = Vec::<u8>::new();
    shared_rgce.extend_from_slice(&ptg_ref3d(0, u16::MAX, 0xC000)); // ixti=0 => Sheet1, A65536, row+col relative
    shared_rgce.push(0x1E); // PtgInt
    shared_rgce.extend_from_slice(&1u16.to_le_bytes());
    shared_rgce.push(0x03); // PtgAdd

    // SHRFMLA record defining shared rgce for range B65535:B65536.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(BASE_ROW, FOLLOW_ROW, COL_B as u8, COL_B as u8, &shared_rgce),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_area3d_oob_shared_formula_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula in B65535:B65536:
    // - B65535: Sheet1!A65535:A65536+1
    // - B65536: #REF!+1 (because Sheet1!A65536:A65537 is out of BIFF8 bounds)
    const BASE_ROW: u16 = u16::MAX - 1; // 65534 => row 65535 (1-based)
    const FOLLOW_ROW: u16 = u16::MAX; // 65535 => row 65536 (1-based)
    const COL_B: u16 = 1;
    let grbit_shared: u16 = 0x0008;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [65534, 65536) cols [1, 2) => B65535:B65536.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(BASE_ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(FOLLOW_ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&COL_B.to_le_bytes()); // first col (B)
    dims.extend_from_slice(&(COL_B + 1).to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Cells store PtgExp pointing at the base cell; the shared rgce is stored in SHRFMLA.
    let ptgexp = ptg_exp(BASE_ROW, COL_B);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(BASE_ROW, COL_B, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    // Shared formula rgce: Sheet1!A65535:A65536 + 1, where PtgArea3d carries relative flags so
    // filling down shifts both endpoints by +1 row.
    let mut shared_rgce = Vec::<u8>::new();
    let col_with_flags: u16 = 0xC000; // col=0 (A) + rowRel + colRel
    shared_rgce.extend_from_slice(&ptg_area3d(
        0,
        BASE_ROW,
        FOLLOW_ROW,
        col_with_flags,
        col_with_flags,
    ));
    shared_rgce.push(0x1E); // PtgInt
    shared_rgce.extend_from_slice(&1u16.to_le_bytes());
    shared_rgce.push(0x03); // PtgAdd

    // SHRFMLA record defining shared rgce for range B65535:B65536.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(BASE_ROW, FOLLOW_ROW, COL_B as u8, COL_B as u8, &shared_rgce),
    );

    // Follower B65536: PtgExp referencing base cell B65535.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(FOLLOW_ROW, COL_B, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_area3d_oob_shrfmla_only_sheet_stream(_xf_cell: u16) -> Vec<u8> {
    // Shared formula definition in SHRFMLA for range B65535:B65536, with no cell-level FORMULA
    // records. The importer must expand/materialize the shared rgce itself.
    //
    // Materialized formulas should match `build_shared_area3d_oob_shared_formula_sheet_stream`:
    // - B65535: Sheet1!A65535:A65536+1
    // - B65536: #REF!+1
    const BASE_ROW: u16 = u16::MAX - 1; // 65534 => row 65535 (1-based)
    const FOLLOW_ROW: u16 = u16::MAX; // 65535 => row 65536 (1-based)
    const COL_B: u16 = 1;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [65534, 65536) cols [1, 2) => B65535:B65536.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(BASE_ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(FOLLOW_ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&COL_B.to_le_bytes()); // first col (B)
    dims.extend_from_slice(&(COL_B + 1).to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Shared formula rgce: Sheet1!A65535:A65536 + 1, where PtgArea3d carries relative flags so
    // filling down shifts both endpoints by +1 row.
    let mut shared_rgce = Vec::<u8>::new();
    let col_with_flags: u16 = 0xC000; // col=0 (A) + rowRel + colRel
    shared_rgce.extend_from_slice(&ptg_area3d(
        0,
        BASE_ROW,
        FOLLOW_ROW,
        col_with_flags,
        col_with_flags,
    ));
    shared_rgce.push(0x1E); // PtgInt
    shared_rgce.extend_from_slice(&1u16.to_le_bytes());
    shared_rgce.push(0x03); // PtgAdd

    // SHRFMLA record defining shared rgce for range B65535:B65536.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(BASE_ROW, FOLLOW_ROW, COL_B as u8, COL_B as u8, &shared_rgce),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
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

fn build_page_setup_malformed_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table: 16 style XFs + 1 cell XF.
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
    write_short_unicode_string(&mut boundsheet, "PageSetupMalformed");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    let sheet_offset = globals.len();
    let sheet = build_page_setup_malformed_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());

    globals.extend_from_slice(&sheet);
    globals
}

fn build_page_setup_malformed_sheet_stream(xf_general: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) (A1).
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Truncated SETUP record payload (<34 bytes).
    //
    // Use an empty payload so the parser can't accidentally apply partially-present fields (e.g.
    // SETUP.grbit orientation), keeping the resulting page setup at defaults.
    push_record(&mut sheet, RECORD_SETUP, &[]);
    // Truncated margin record payload (<8 bytes).
    push_record(&mut sheet, RECORD_LEFTMARGIN, &[0u8; 4]);
    // Truncated HORIZONTALPAGEBREAKS: cbrk=2 but no Brk entries present.
    let breaks = 2u16.to_le_bytes();
    push_record(&mut sheet, RECORD_HORIZONTALPAGEBREAKS, &breaks);

    // A1: a single cell so calamine reports a non-empty range.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

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
    push_record(
        &mut sheet,
        RECORD_LEFTMARGIN,
        &cfg.left_margin.to_le_bytes(),
    );
    push_record(
        &mut sheet,
        RECORD_RIGHTMARGIN,
        &cfg.right_margin.to_le_bytes(),
    );
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
        &hpagebreaks_record(&[cfg.row_break_after]),
    );
    push_record(
        &mut sheet,
        RECORD_VPAGEBREAKS,
        &vpagebreaks_record(&[cfg.col_break_after]),
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

fn build_margins_without_setup_sheet_stream(
    xf_cell: u16,
    left: f64,
    right: f64,
    top: f64,
    bottom: f64,
) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Page margin records (no SETUP record).
    push_record(&mut sheet, RECORD_LEFTMARGIN, &left.to_le_bytes());
    push_record(&mut sheet, RECORD_RIGHTMARGIN, &right.to_le_bytes());
    push_record(&mut sheet, RECORD_TOPMARGIN, &top.to_le_bytes());
    push_record(&mut sheet, RECORD_BOTTOMMARGIN, &bottom.to_le_bytes());

    // A1: a single cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

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

fn name_record_malformed_claimed_cce(name: &str, claimed_cce: u16) -> Vec<u8> {
    // NAME record payload (BIFF8) with an intentionally inconsistent `cce` field.
    //
    // The `cce` header claims there are `claimed_cce` bytes of `rgce`, but we do not include any
    // `rgce` bytes in the payload. This is used to exercise calamine hardening and our
    // `sanitize_biff8_continued_name_records_for_calamine()` workaround, which patches `cce` to 0
    // to prevent panics on out-of-bounds `rgce` slicing.
    let mut out = Vec::<u8>::new();

    out.extend_from_slice(&0u16.to_le_bytes()); // grbit
    out.push(0); // chKey

    let cch: u8 = name
        .len()
        .try_into()
        .expect("defined name too long for u8 length");
    out.push(cch);

    out.extend_from_slice(&claimed_cce.to_le_bytes()); // cce (malformed)
    out.extend_from_slice(&0u16.to_le_bytes()); // ixals
    out.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook scoped)

    out.push(0); // cchCustMenu
    out.push(0); // cchDescription
    out.push(0); // cchHelpTopic
    out.push(0); // cchStatusText

    // Name string (XLUnicodeStringNoCch).
    write_unicode_string_no_cch(&mut out, name);

    // Intentionally omit rgce bytes.
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

fn ptg_exp_row_u32_col_u16(row: u32, col: u16) -> Vec<u8> {
    // Non-standard PtgExp payload (seen in the wild): [rw: u32][col: u16]
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

fn ptg_area3d_with_rel_flags(
    ixti: u16,
    row1: u16,
    row2: u16,
    col1: u16,
    col2: u16,
    row1_rel: bool,
    col1_rel: bool,
    row2_rel: bool,
    col2_rel: bool,
) -> Vec<u8> {
    // PtgArea3d (0x3B) payload with independent relative flags for each endpoint:
    //   [ixti: u16]
    //   [rowFirst: u16][rowLast: u16]
    //   [colFirst+flags: u16][colLast+flags: u16]
    //
    // BIFF8 stores relative row/col flags in the high bits of the column fields:
    // - bit 14: row-relative
    // - bit 15: col-relative
    let col1_u16 = pack_biff8_col_flags(col1, row1_rel, col1_rel);
    let col2_u16 = pack_biff8_col_flags(col2, row2_rel, col2_rel);
    ptg_area3d(ixti, row1, row2, col1_u16, col2_u16)
}

fn pack_biff8_col_flags(col: u16, row_rel: bool, col_rel: bool) -> u16 {
    let mut out = col & 0x3FFF;
    if row_rel {
        out |= 0x4000;
    }
    if col_rel {
        out |= 0x8000;
    }
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

fn build_shared_formula_2d_workbook_stream() -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: 16 style XFs + one cell XF used by our formula cells.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    let xf_cell = 16u16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

    // Single worksheet containing a shared formula group.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Shared");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    let sheet_offset = globals.len();
    let sheet = build_shared_formula_2d_sheet_stream(xf_cell);

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

fn build_shared_formula_master_not_top_left_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF: worksheet

    // DIMENSIONS: rows [0, 2) cols [0, 2) => A1:B2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2()); // WINDOW2

    // Values in A1/A2.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));

    // Shared formula range is B1:B2, but both FORMULA records reference B2 as the PtgExp master.
    // PtgExp payload: [ptg=0x01][rw:u16][col:u16] where rw/col are 0-based indices.
    let ptgexp_master_b2: [u8; 5] = [0x01, 0x01, 0x00, 0x01, 0x00];

    // B1 formula record (row=0,col=1) uses PtgExp.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 1.0, &ptgexp_master_b2),
    );

    // SHRFMLA definition for range B1:B2 (rwFirst=0,rwLast=1,colFirst=1,colLast=1) with an `rgce`
    // that references the cell to the left (`PtgRefN` col_off=-1).
    let rgce_ref_left: [u8; 5] = [0x2C, 0x00, 0x00, 0xFF, 0xFF]; // PtgRefN(row_off=0,col_off=-1)
    let shrfmla = shrfmla_record(0, 1, 1, 1, &rgce_ref_left);
    push_record(&mut sheet, RECORD_SHRFMLA, &shrfmla);

    // B2 formula record (row=1,col=1) uses PtgExp.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 2.0, &ptgexp_master_b2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn build_shared_formula_master_not_top_left_workbook_stream() -> Vec<u8> {
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

    // Single worksheet.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Shared");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    let sheet_offset = globals.len();
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_formula_master_not_top_left_sheet_stream(
        xf_cell,
    ));

    globals
}

fn build_shared_formula_sheet_scoped_name_sanitization_workbook_stream() -> Vec<u8> {
    // This workbook contains:
    // - Sheet 0: `Bad:Name` (invalid; will be sanitized to `Bad_Name` on import).
    // - Sheet 1: `Ref`, with a **shared formula** in A1:A2 whose shared rgce is `PtgName` pointing
    //   at a sheet-scoped defined name on `Bad:Name`.
    //
    // This exercises BIFF-decoded shared formula rendering that must resolve:
    // - sheet-scoped `PtgName` indices (NAME record order), and
    // - sheet prefixes using the final imported (sanitized) sheet name.
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

    // Sheet-scoped defined name on `Bad:Name` (itab=1, since itab-1 is BIFF sheet index).
    //
    // The rgce payload is arbitrary for this fixture; we use a simple 2D `PtgRef` to A1.
    let name_rgce: Vec<u8> = vec![
        0x24, // PtgRef
        0x00, 0x00, // row = 0
        0x00, 0x00, // col = 0 (absolute)
    ];
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record(
            "LocalName",
            /*itab*/ 1,
            /*hidden*/ false,
            None,
            &name_rgce,
        ),
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
    globals.extend_from_slice(&build_shared_ptgname_shrfmla_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_sheet_scoped_name_simple_workbook_stream(scoped_sheet_name: &str) -> Vec<u8> {
    // Shared formula (`SHRFMLA` + `PtgExp`) whose shared rgce is `PtgName(1)`, referencing a
    // sheet-scoped defined name on `scoped_sheet_name`.
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
    for name in [scoped_sheet_name, "Ref"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // Sheet-scoped defined name on `scoped_sheet_name` (itab=1 => sheet index 0).
    let name_rgce: Vec<u8> = vec![
        0x24, // PtgRef
        0x00, 0x00, // row = 0
        0x00, 0x00, // col = 0 (absolute)
    ];
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record(
            "LocalName",
            /*itab*/ 1,
            /*hidden*/ false,
            None,
            &name_rgce,
        ),
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
    globals.extend_from_slice(&build_shared_ptgname_shrfmla_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_sheet_scoped_name_apostrophe_workbook_stream() -> Vec<u8> {
    build_shared_formula_sheet_scoped_name_simple_workbook_stream("O'Brien")
}

fn build_shared_formula_sheet_scoped_name_true_sheet_workbook_stream() -> Vec<u8> {
    build_shared_formula_sheet_scoped_name_simple_workbook_stream("TRUE")
}

fn build_shared_formula_sheet_scoped_name_a1_sheet_workbook_stream() -> Vec<u8> {
    build_shared_formula_sheet_scoped_name_simple_workbook_stream("A1")
}

fn build_shared_formula_sheet_scoped_name_unicode_sheet_workbook_stream() -> Vec<u8> {
    // Use a name that exercises non-ASCII handling while still being representable in BIFF8's
    // "compressed" string form (all UTF-16 code units <= 0x00FF) so it is robust across BIFF8
    // decoders.
    build_shared_formula_sheet_scoped_name_simple_workbook_stream("Résumé")
}

fn build_shared_formula_sheet_scoped_name_dedup_collision_workbook_stream() -> Vec<u8> {
    // This workbook contains:
    // - Sheet 0: `Bad:Name` (invalid; will be sanitized on import)
    // - Sheet 1: `Bad_Name` (valid, collides with the sanitized form)
    // - Sheet 2: `Ref`, with a shared formula (A1:A2) whose rgce is `PtgName` referencing a
    //   sheet-scoped defined name on sheet 0.
    //
    // This exercises the case where sheet-name sanitization causes a name collision and one of
    // the sheets must be deduped. The `PtgName` scope sheet index must still resolve to the
    // correct output sheet name.
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
    for name in ["Bad:Name", "Bad_Name", "Ref"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // Sheet-scoped defined name on `Bad:Name` (itab=1 => sheet index 0).
    let name_rgce: Vec<u8> = vec![
        0x24, // PtgRef
        0x00, 0x00, // row = 0
        0x00, 0x00, // col = 0 (absolute)
    ];
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record(
            "LocalName",
            /*itab*/ 1,
            /*hidden*/ false,
            None,
            &name_rgce,
        ),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0 ------------------------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 111.0));

    // -- Sheet 1 ------------------------------------------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 222.0));

    // -- Sheet 2 ------------------------------------------------------------------
    let sheet2_offset = globals.len();
    globals[boundsheet_offset_positions[2]..boundsheet_offset_positions[2] + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_ptgname_shrfmla_sheet_stream(xf_cell));

    globals
}

fn build_shared_formula_sheet_scoped_name_dedup_collision_invalid_second_workbook_stream() -> Vec<u8>
{
    // Similar to `build_shared_formula_sheet_scoped_name_dedup_collision_workbook_stream`, but with
    // the colliding *valid* sheet name first. This means the invalid sheet's sanitized name must be
    // deduped (e.g. `Bad_Name (2)`), and `PtgName` resolution must use that final deduped name.
    //
    // This workbook contains:
    // - Sheet 0: `Bad_Name` (valid)
    // - Sheet 1: `Bad:Name` (invalid; sanitizes to `Bad_Name` and must be deduped)
    // - Sheet 2: `Ref`, with a shared formula (A1:A2) whose rgce is `PtgName` referencing a
    //   sheet-scoped defined name on sheet 1.
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
    for name in ["Bad_Name", "Bad:Name", "Ref"] {
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, name);
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        boundsheet_offset_positions.push(boundsheet_start + 4);
    }

    // Sheet-scoped defined name on `Bad:Name` (itab=2 => sheet index 1).
    let name_rgce: Vec<u8> = vec![
        0x24, // PtgRef
        0x00, 0x00, // row = 0
        0x00, 0x00, // col = 0 (absolute)
    ];
    push_record(
        &mut globals,
        RECORD_NAME,
        &name_record(
            "LocalName",
            /*itab*/ 2,
            /*hidden*/ false,
            None,
            &name_rgce,
        ),
    );

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet 0 (valid name) -----------------------------------------------------
    let sheet0_offset = globals.len();
    globals[boundsheet_offset_positions[0]..boundsheet_offset_positions[0] + 4]
        .copy_from_slice(&(sheet0_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 222.0));

    // -- Sheet 1 (invalid name, will be deduped) ----------------------------------
    let sheet1_offset = globals.len();
    globals[boundsheet_offset_positions[1]..boundsheet_offset_positions[1] + 4]
        .copy_from_slice(&(sheet1_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_simple_number_sheet_stream(xf_cell, 111.0));

    // -- Sheet 2 ------------------------------------------------------------------
    let sheet2_offset = globals.len();
    globals[boundsheet_offset_positions[2]..boundsheet_offset_positions[2] + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());
    globals.extend_from_slice(&build_shared_ptgname_shrfmla_sheet_stream(xf_cell));

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

fn build_array_formula_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [0, 2) (A..B, rows 1..2).
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: NUMBER record (ensures calamine surfaces a non-empty range).
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // Array formula group: B1:B2, formula `A1:A2`.
    let base_row = 0u16;
    let base_col = 1u16; // B

    // Formula tokens for `A1:A2`: PtgArea with relative flags so the rendered formula is `A1:A2`
    // (no `$`).
    let col_with_flags: u16 = 0xC000; // col=0 (A) + rowRel + colRel
    let array_rgce = ptg_area(0, 1, col_with_flags, col_with_flags);

    // B1 FORMULA: PtgExp -> base cell (B1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(
            base_row,
            base_col,
            xf_cell,
            0.0,
            &ptg_exp(base_row, base_col),
        ),
    );

    // ARRAY record stores the formula token stream for the whole group.
    push_record(
        &mut sheet,
        RECORD_ARRAY,
        &array_record_refu(base_row, 1, base_col as u8, base_col as u8, &array_rgce),
    );

    // B2 FORMULA: PtgExp -> base cell (B1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, base_col, xf_cell, 0.0, &ptg_exp(base_row, base_col)),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_array_formula_ptgarray_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Array formula range B1:B2.
    //
    // Both cells contain PtgExp, and the array formula body is stored in the `ARRAY` record along
    // with trailing `rgcb` bytes needed to decode a `PtgArray` constant.
    //
    // Expected decoded formula (same for B1 and B2):
    //   `A1+SUM({1,2;3,4})`
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
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));

    // Set FORMULA.grbit.fArray (0x0010) so parsers recognize the array-formula membership.
    let grbit_array: u16 = 0x0010;

    let base_row = 0u16;
    let base_col = 1u16; // B

    // B1 formula: PtgExp pointing to itself (rw=0,col=1).
    let ptgexp = ptg_exp(base_row, base_col);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(base_row, base_col, xf_cell, 0.0, grbit_array, &ptgexp),
    );

    // Array formula body stored in ARRAY record.
    let array_rgce = {
        let mut v = Vec::new();
        // PtgRefN: row_off=0, col_off=-1 relative to the array base cell (B1) => A1.
        v.push(0x2C);
        v.extend_from_slice(&0u16.to_le_bytes()); // row_off = 0
        v.extend_from_slice(&0xFFFFu16.to_le_bytes()); // col_off = -1 (14-bit), row+col relative

        // PtgArray (array constant; data stored in trailing rgcb).
        v.push(0x20);
        v.extend_from_slice(&[0u8; 7]); // reserved

        // PtgFuncVar: SUM(argc=1).
        v.push(0x22);
        v.push(1);
        v.extend_from_slice(&4u16.to_le_bytes());

        // PtgAdd.
        v.push(0x03);
        v
    };

    let rgcb = rgcb_array_constant_numbers_2x2(&[1.0, 2.0, 3.0, 4.0]);
    let mut array_payload =
        array_record_refu(base_row, 1, base_col as u8, base_col as u8, &array_rgce);
    array_payload.extend_from_slice(&rgcb);
    push_record(&mut sheet, RECORD_ARRAY, &array_payload);

    // B2 formula: PtgExp pointing to base cell B1.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, base_col, xf_cell, 0.0, grbit_array, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_array_formula_range_flags_ambiguity_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 10) cols [0, 3) => A1:C10.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&10u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide a value cell so calamine reports a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 2, xf_cell, 1.0)); // C1

    // Set FORMULA.grbit.fArray (0x0010) to mark the cell as part of an array formula group.
    let grbit_array: u16 = 0x0010;

    // -- ARRAY #1 --------------------------------------------------------------
    // Range A1:A2 encoded as RefU + flags + cce, where flags is a small non-zero value.
    let rgce_left: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x17); // PtgStr
        v.push(4); // cch
        v.push(0); // flags (compressed)
        v.extend_from_slice(b"LEFT");
        v
    };
    let mut array_left = Vec::<u8>::new();
    array_left.extend_from_slice(&0u16.to_le_bytes()); // rwFirst = 0
    array_left.extend_from_slice(&1u16.to_le_bytes()); // rwLast = 1
    array_left.push(0); // colFirst = A
    array_left.push(0); // colLast = A
    array_left.extend_from_slice(&2u16.to_le_bytes()); // flags/reserved (non-zero)
    array_left.extend_from_slice(&(rgce_left.len() as u16).to_le_bytes()); // cce
    array_left.extend_from_slice(&rgce_left);
    push_record(&mut sheet, RECORD_ARRAY, &array_left);

    // -- ARRAY #2 --------------------------------------------------------------
    // Array formula range B1:B10 whose body is `C1+1` (decoded relative to base cell B1).
    // PtgRefN(row_off=0,col_off=+1) + PtgInt(1) + PtgAdd
    let rgce_right: Vec<u8> = vec![
        0x2C, // PtgRefN
        0x00, 0x00, // row_off = 0
        0x01, 0xC0, // col_off = +1 with row+col relative flags
        0x1E, // PtgInt
        0x01, 0x00, // 1
        0x03, // PtgAdd
    ];
    push_record(
        &mut sheet,
        RECORD_ARRAY,
        &array_record_refu(0, 9, 1, 1, &rgce_right),
    );

    // B2 FORMULA record: PtgExp(B2) (self-reference).
    let ptgexp_b2 = ptg_exp(1, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, 1, xf_cell, 0.0, grbit_array, &ptgexp_b2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_array_formula_ptgname_sheet_stream(xf_cell: u16) -> Vec<u8> {
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

    // A1: NUMBER record (ensures calamine surfaces a non-empty range).
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // Set FORMULA.grbit.fArray (0x0010) so parsers recognize the array-formula membership.
    let grbit_array: u16 = 0x0010;

    // Array formula group: B1:B2, formula `MyName+1`.
    let base_row = 0u16;
    let base_col = 1u16; // B

    // B1/B2 formulas: PtgExp pointing at the array base cell (B1).
    let ptgexp = ptg_exp(base_row, base_col);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(base_row, base_col, xf_cell, 0.0, grbit_array, &ptgexp),
    );

    // Formula tokens for `MyName+1`: PtgName(1) + PtgInt(1) + PtgAdd.
    let array_rgce = [ptg_name(1), vec![0x1E, 0x01, 0x00], vec![0x03]].concat();
    push_record(
        &mut sheet,
        RECORD_ARRAY,
        &array_record_refu(base_row, 1, base_col as u8, base_col as u8, &array_rgce),
    );

    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, base_col, xf_cell, 0.0, grbit_array, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_array_formula_ptgexp_missing_array_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Malformed array formula range B1:B2:
    // - B1 stores a full token stream (`A1+1`) with relative flags set.
    // - B2 stores only `PtgExp(B1)` and sets FORMULA.grbit.fArray, but the `ARRAY` record is
    //   missing.
    //
    // Expected recovered formula for B2: `A1+1` (anchored at the base cell, no shifting).
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

    // Provide at least one numeric cell so calamine surfaces a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 0.0));

    let base_row = 0u16;
    let base_col = 1u16; // B

    // Base formula in B1: `A1+1` with relative row/col flags set (so shared-formula materialization
    // would become `A2+1` when filled down).
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
        &formula_cell(base_row, base_col, xf_cell, 0.0, &rgce_base),
    );

    // Follower formula in B2: `PtgExp` pointing at base cell B1 (row=0, col=1) with fArray set.
    let grbit_array: u16 = 0x0010;
    let rgce_ptgexp = ptg_exp(base_row, base_col);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, base_col, xf_cell, 0.0, grbit_array, &rgce_ptgexp),
    );

    // Intentionally omit the ARRAY definition record.

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_array_formula_external_refs_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [0, 3) => A1:C2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: NUMBER record (ensures calamine surfaces a non-empty range).
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));

    // Array formula group #1: B1:B2, formula `'[Book1.xlsx]ExtSheet'!$A$1+1` (PtgRef3d + 1).
    let base_b_row = 0u16;
    let base_b_col = 1u16; // B
    let array_b_rgce = [ptg_ref3d(0, 0, 0), vec![0x1E, 0x01, 0x00], vec![0x03]].concat();

    // B1 FORMULA: PtgExp -> base cell (B1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(
            base_b_row,
            base_b_col,
            xf_cell,
            0.0,
            &ptg_exp(base_b_row, base_b_col),
        ),
    );
    // ARRAY record stores the shared rgce.
    push_record(
        &mut sheet,
        RECORD_ARRAY,
        &array_record_refu(
            base_b_row,
            1,
            base_b_col as u8,
            base_b_col as u8,
            &array_b_rgce,
        ),
    );
    // B2 FORMULA: PtgExp -> base cell (B1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(
            1,
            base_b_col,
            xf_cell,
            0.0,
            &ptg_exp(base_b_row, base_b_col),
        ),
    );

    // Array formula group #2: C1:C2, formula `'[Book1.xlsx]ExtSheet'!ExtDefined+1` (PtgNameX + 1).
    let base_c_row = 0u16;
    let base_c_col = 2u16; // C
    let array_c_rgce = [ptg_namex(0, 1), vec![0x1E, 0x01, 0x00], vec![0x03]].concat();

    // C1 FORMULA: PtgExp -> base cell (C1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(
            base_c_row,
            base_c_col,
            xf_cell,
            0.0,
            &ptg_exp(base_c_row, base_c_col),
        ),
    );
    push_record(
        &mut sheet,
        RECORD_ARRAY,
        &array_record_refu(
            base_c_row,
            1,
            base_c_col as u8,
            base_c_col as u8,
            &array_c_rgce,
        ),
    );
    // C2 FORMULA: PtgExp -> base cell (C1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(
            1,
            base_c_col,
            xf_cell,
            0.0,
            &ptg_exp(base_c_row, base_c_col),
        ),
    );

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

fn build_shared_area3d_shared_formula_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // SharedArea3D!B1:B2:
    //   B1: Sheet1!A1:A2
    //   B2: Sheet1!A2:A3
    //
    // Shared formula definition (SHRFMLA) uses PtgArea3d with relative flags so materialization
    // must shift both area endpoints by +1 row for B2.
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [1, 2) => B1:B2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&1u16.to_le_bytes()); // first col (B)
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Shared-formula base rgce: PtgArea3d Sheet1!A1:A2 (ixti=0), with both endpoints row/col-relative.
    let base_rgce = ptg_area3d_with_rel_flags(0, 0, 1, 0, 0, true, true, true, true);

    // Base cell B1: full formula token stream.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &base_rgce),
    );

    // Shared formula definition for B1:B2.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &base_rgce),
    );

    // B2: PtgExp referencing base cell B1 (row=0, col=1).
    let ptgexp = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_area3d_shared_formula_mixed_flags_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // SharedArea3D!B1:B2:
    //   B1: Sheet1!A1:$A2
    //   B2: Sheet1!A2:$A3
    //
    // This uses different relative flags for the two area endpoints:
    // - start endpoint: row+col relative (no '$') => `A1`
    // - end endpoint: row relative, col absolute => `$A2`
    //
    // Materialization across the shared formula must shift both rows and preserve the per-endpoint
    // `$` flags.
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [1, 2) => B1:B2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&1u16.to_le_bytes()); // first col (B)
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Shared-formula base rgce: PtgArea3d Sheet1!A1:$A2 (ixti=0).
    // Endpoint 1: rowRel + colRel (A1)
    // Endpoint 2: rowRel + colAbs ($A2)
    let base_rgce = ptg_area3d_with_rel_flags(0, 0, 1, 0, 0, true, true, true, false);

    // Base cell B1: full formula token stream.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &base_rgce),
    );

    // Shared formula definition for B1:B2.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &base_rgce),
    );

    // B2: PtgExp referencing base cell B1 (row=0, col=1).
    let ptgexp = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &ptgexp),
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
    //
    // Set FORMULA.grbit.fShrFmla (0x0008) so parsers recognize the shared-formula membership.
    let grbit_shared: u16 = 0x0008;
    let ptgexp = ptg_exp(0, 0);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(0, 0, xf_cell, 0.0, grbit_shared, &ptgexp),
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
        &formula_cell_with_grbit(1, 0, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_ptgname_shrfmla_sheet_stream(xf_cell: u16) -> Vec<u8> {
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

    // Shared formula anchor: A1 formula token stream is PtgExp(A1), followed by SHRFMLA containing
    // the shared rgce.
    //
    // Set FORMULA.grbit.fShrFmla (0x0008) so parsers recognize the shared-formula membership.
    let grbit_shared: u16 = 0x0008;
    let ptgexp = ptg_exp(0, 0);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(0, 0, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    // Shared rgce: `PtgName` referencing NAME record index 1 (LocalName, sheet-scoped to Bad:Name).
    let shared_rgce = ptg_name(1);
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 0, 0, &shared_rgce),
    );

    // Second cell in the shared range: A2 formula record containing PtgExp(A1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(1, 0, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_ptgmemarean_sheet_stream(xf_cell: u16) -> Vec<u8> {
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

    // Provide inputs in A1/A2 (not strictly required for formula decoding).
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 1.0));
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(1, 0, xf_cell, 2.0));

    // Shared formula over B1:B2:
    //   B1: A1+1
    //   B2: A2+1 (via PtgExp)
    //
    // Base (full) formula for B1: PtgRef(A1, relative) + PtgInt(1) + PtgAdd
    let full_rgce: Vec<u8> = vec![
        0x24, // PtgRef
        0x00, 0x00, // row = 0 (A1)
        0x00, 0xC0, // col = 0 (A), row+col relative
        0x1E, // PtgInt
        0x01, 0x00, // 1
        0x03, // PtgAdd
    ];

    // Mark the base cell as a shared-formula anchor via FORMULA.grbit.fShrFmla (0x0008).
    let mut b1 = formula_cell(0, 1, xf_cell, 0.0, &full_rgce);
    b1[14..16].copy_from_slice(&0x0008u16.to_le_bytes());
    push_record(&mut sheet, RECORD_FORMULA, &b1);

    // Shared SHRFMLA record containing the base rgce in relative form:
    //   PtgRefN(row_off=0,col_off=-1) + PtgMemAreaN(cce=0) + PtgMemAreaN(cce=3, rgce=PtgInt(0)) + PtgInt(1) + PtgAdd
    //
    // Note: PtgMemAreaN is a no-op for printing but carries a payload; decoders must still skip the
    // `cce` field to keep the token stream aligned.
    let shared_rgce: Vec<u8> = vec![
        0x2C, // PtgRefN
        0x00, 0x00, // row_off = 0
        0xFF, 0xFF, // col_off = -1 (14-bit two's complement) + row/col relative flags
        0x2E, // PtgMemAreaN
        0x00, 0x00, // cce = 0
        0x2E, // PtgMemAreaN (with a nested rgce payload)
        0x03, 0x00, // cce = 3
        0x1E, // PtgInt
        0x00, 0x00, // 0
        0x1E, // PtgInt
        0x01, 0x00, // 1
        0x03, // PtgAdd
    ];
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &shared_rgce),
    );

    // B2 uses PtgExp to reference base cell B1 (row=0,col=1).
    let exp_rgce: Vec<u8> = vec![0x01, 0x00, 0x00, 0x01, 0x00];
    let mut b2 = formula_cell(1, 1, xf_cell, 0.0, &exp_rgce);
    b2[14..16].copy_from_slice(&0x0008u16.to_le_bytes());
    push_record(&mut sheet, RECORD_FORMULA, &b2);

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_out_of_bounds_relative_refs_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // BIFF8 max row index (0-based).
    const MAX_ROW0: u32 = u16::MAX as u32;
    let base_row: u32 = MAX_ROW0 - 1;
    let base_col: u16 = 1; // column B

    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: cover rows [base_row, MAX_ROW0+1) and cols [0, 2) (A..B).
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&base_row.to_le_bytes()); // first row
    dims.extend_from_slice(&(MAX_ROW0 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    let base_row_u16: u16 = base_row as u16;
    let max_row_u16: u16 = MAX_ROW0 as u16;
    let base_col_u8: u8 = base_col as u8;

    // Shared formula definition (SHRFMLA) for B65535:B65536:
    //   B65535: A65536+1
    //   B65536: #REF!+1 (because A65537 is out of bounds in BIFF8)
    //
    // Base rgce stored in SHRFMLA: PtgRefN(row_off=+1,col_off=-1) + 1 + +
    let shared_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x2C); // PtgRefN
        v.extend_from_slice(&1u16.to_le_bytes()); // row_off=+1 (stored in rw when row-relative)
        v.extend_from_slice(&0xFFFFu16.to_le_bytes()); // col_off=-1 (0x3FFF) with row+col relative bits set
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };

    // Base cell (B65535) formula record, marked as a shared formula anchor.
    //
    // Use a full formula rgce so calamine can decode B65535 even if SHRFMLA resolution fails.
    // `A65536+1` => PtgRef(A65536, relative row/col) + 1 + +
    let base_full_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x24); // PtgRef
        v.extend_from_slice(&max_row_u16.to_le_bytes()); // row = 65535 (A65536)
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, row+col relative
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(
            base_row_u16,
            base_col,
            xf_cell,
            0.0,
            FORMULA_FLAG_SHARED,
            &base_full_rgce,
        ),
    );

    // SHRFMLA record must immediately follow the base FORMULA record.
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(
            base_row_u16,
            max_row_u16,
            base_col_u8,
            base_col_u8,
            &shared_rgce,
        ),
    );

    // Follower cell (B65536) uses PtgExp to reference the shared formula base cell (B65535).
    let ptgexp: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x01); // PtgExp
        v.extend_from_slice(&base_row_u16.to_le_bytes()); // base row
        v.extend_from_slice(&base_col.to_le_bytes()); // base col
        v
    };
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(max_row_u16, base_col, xf_cell, 0.0, &ptgexp),
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

fn build_shared_formula_2d_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [0, 3) => A1:C2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Shared formula range B1:C2.
    //
    // - Anchor/base cell: B1.
    // - Shared rgce (stored in SHRFMLA): `PtgRefN(col_off=-1) + 1 + +`
    //   so the materialized formula in each cell references the cell to the left.
    //
    // Expected materialization:
    //   B1 = A1+1, C1 = B1+1, B2 = A2+1, C2 = B2+1
    let shared_rgce: Vec<u8> = {
        let row_off: i16 = 0;
        let col_off: i16 = -1;
        let col14: u16 = (col_off as u16) & 0x3FFF;
        let col_field: u16 = col14 | 0xC000; // row+col relative

        let mut v = Vec::new();
        v.push(0x2C); // PtgRefN
        v.extend_from_slice(&(row_off as u16).to_le_bytes());
        v.extend_from_slice(&col_field.to_le_bytes());
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };

    // Anchor cell B1: store a PtgExp(B1) formula, followed by SHRFMLA containing the shared rgce.
    let ptgexp = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &ptgexp),
    );
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 2, &shared_rgce),
    );

    // Remaining cells in the shared range store PtgExp(B1).
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 2, xf_cell, 0.0, &ptgexp),
    ); // C1
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &ptgexp),
    ); // B2
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 2, xf_cell, 0.0, &ptgexp),
    ); // C2

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_external_refs_sheet_stream(xf_cell: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();

    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 2) cols [0, 3) => A1:C2.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&2u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide a value cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 0.0));

    // Shared formula range #1: B1:B2.
    // Base cell B1 stores PtgExp -> (B1) and SHRFMLA stores the shared rgce.
    let exp_b1 = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 1, xf_cell, 0.0, &exp_b1),
    );
    let shrf_rgce = [ptg_ref3d(0, 0, 0), vec![0x1E, 0x01, 0x00], vec![0x03]].concat(); // ExtSheet!$A$1 + 1
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 1, 1, &shrf_rgce),
    );
    // Follower cell B2 stores PtgExp -> (B1).
    let exp_b2 = ptg_exp(0, 1);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 1, xf_cell, 0.0, &exp_b2),
    );

    // Shared formula range #2: C1:C2 using PtgNameX.
    let exp_c1 = ptg_exp(0, 2);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(0, 2, xf_cell, 0.0, &exp_c1),
    );
    let namex_rgce = [ptg_namex(0, 1), vec![0x1E, 0x01, 0x00], vec![0x03]].concat(); // ExtDefined + 1
    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 1, 2, 2, &namex_rgce),
    );
    let exp_c2 = ptg_exp(0, 2);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(1, 2, xf_cell, 0.0, &exp_c2),
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

fn build_sheet_protection_allow_flags_mask_offset_sheet_stream() -> Vec<u8> {
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

    // FEATHEADR header data with prefix bytes: [0xFF,0xFF] + allow_mask. The importer should scan
    // past the prefix and recover the allow-mask.
    push_record(
        &mut sheet,
        RECORD_FEATHEADR,
        &feat_hdr_record_sheet_protection_allow_mask_prefixed(allow_mask),
    );

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

/// Build a BIFF8 `.xls` fixture containing a hyperlink whose display string is encoded as a BIFF8
/// `XLUnicodeString` (u16 length + flags) rather than the standard hyperlink u32-length prefix.
///
/// `parse_hyperlink_string` should fall back to BIFF8 string decoding and still import the link.
pub fn build_biff8_unicode_string_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "Biff8Display",
        hlink_internal_location_biff8_unicode_display(
            0,
            0,
            0,
            0,
            "Biff8Display!B2",
            "Foo\u{0}Bar",
            "Tip",
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

/// Build a BIFF8 `.xls` fixture containing a shared formula (`SHRFMLA`) whose follower `PtgExp`
/// token references a master cell that is *not* the shared range's top-left cell.
///
/// This is used to validate our best-effort shared-formula association logic (range containment
/// match rather than "key by range.start only").
pub fn build_shared_formula_master_not_top_left_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_master_not_top_left_workbook_stream();

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

fn build_shared_formula_ptgref_row_oob_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet = build_shared_formula_ptgref_row_oob_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("Shared", &sheet, 1252)
}

fn build_shared_formula_ptgarea_row_oob_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet = build_shared_formula_ptgarea_row_oob_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedArea", &sheet, 1252)
}

fn build_shared_formula_ptgref_row_oob_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula over the last two BIFF8 rows:
    //   B65535: A65536+1
    //   B65536: #REF!+1 (because A65537 is out of BIFF8 bounds)
    //
    // The cells themselves contain only PtgExp; the shared rgce body is stored in SHRFMLA. This
    // forces the importer to materialize the shared rgce into per-cell formulas.
    const BASE_ROW: u16 = 65_534; // 0-based row for Excel row 65535
    const FOLLOWER_ROW: u16 = 65_535; // 0-based row for Excel row 65536
    const BASE_COL: u16 = 1; // column B

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [BASE_ROW, FOLLOWER_ROW + 1) cols [0, 2) => A65535:B65536.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(BASE_ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(FOLLOWER_ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Set FORMULA.grbit.fShrFmla (0x0008) so parsers recognize the shared-formula membership.
    let grbit_shared: u16 = 0x0008;

    // Base cell B65535: PtgExp pointing to itself.
    let ptgexp = ptg_exp(BASE_ROW, BASE_COL);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(BASE_ROW, BASE_COL, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    // Shared rgce body stored in SHRFMLA. This uses a PtgRef (not PtgRefN) with relative flags so
    // materialization must shift the row.
    let rgce_shared: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x24); // PtgRef
        v.extend_from_slice(&FOLLOWER_ROW.to_le_bytes()); // rw = 65535 (A65536)
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, row+col relative flags
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(
            BASE_ROW,
            FOLLOWER_ROW,
            BASE_COL as u8,
            BASE_COL as u8,
            &rgce_shared,
        ),
    );

    // Follower cell B65536: PtgExp referencing base cell B65535.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(FOLLOWER_ROW, BASE_COL, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

/// Build a BIFF8 `.xls` fixture containing a shared formula that overflows a `PtgRef` row index
/// during materialization near the BIFF8 row limit.
pub fn build_shared_formula_ptgref_row_oob_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgref_row_oob_workbook_stream();

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

fn build_shared_formula_ptgarea_row_oob_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula over the last two BIFF8 rows:
    //   B65535: SUM(A65535:A65536)+1
    //   B65536: SUM(#REF!)+1 (because A65536:A65537 is out of BIFF8 bounds)
    //
    // The cells themselves contain only PtgExp; the shared rgce body is stored in SHRFMLA. This
    // forces the importer to materialize the shared rgce into per-cell formulas.
    const BASE_ROW: u16 = 65_534; // 0-based row for Excel row 65535
    const FOLLOWER_ROW: u16 = 65_535; // 0-based row for Excel row 65536
    const BASE_COL: u16 = 1; // column B
    let col_with_flags: u16 = 0xC000; // col=A + row+col relative flags

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [BASE_ROW, FOLLOWER_ROW + 1) cols [0, 2) => A65535:B65536.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(BASE_ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(FOLLOWER_ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Set FORMULA.grbit.fShrFmla (0x0008) so parsers recognize the shared-formula membership.
    let grbit_shared: u16 = 0x0008;

    // Base cell B65535: PtgExp pointing to itself.
    let ptgexp = ptg_exp(BASE_ROW, BASE_COL);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(BASE_ROW, BASE_COL, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    // Shared rgce body stored in SHRFMLA. This uses a PtgArea (not PtgAreaN) with relative flags
    // on both endpoints so materialization must shift both row indices.
    let rgce_shared: Vec<u8> = {
        let mut v = Vec::new();
        // Base formula area: A65535:A65536
        v.extend_from_slice(&ptg_area(
            BASE_ROW,
            FOLLOWER_ROW,
            col_with_flags,
            col_with_flags,
        ));
        // SUM(...)
        v.push(0x22); // PtgFuncVar
        v.push(1); // argc=1
        v.extend_from_slice(&0x0004u16.to_le_bytes()); // iftab=4 (SUM)
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(
            BASE_ROW,
            FOLLOWER_ROW,
            BASE_COL as u8,
            BASE_COL as u8,
            &rgce_shared,
        ),
    );

    // Follower cell B65536: PtgExp referencing base cell B65535.
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell_with_grbit(FOLLOWER_ROW, BASE_COL, xf_cell, 0.0, grbit_shared, &ptgexp),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

/// Build a BIFF8 `.xls` fixture containing a shared formula that overflows a `PtgArea` row index
/// during materialization near the BIFF8 row limit.
pub fn build_shared_formula_ptgarea_row_oob_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgarea_row_oob_workbook_stream();

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

fn build_shared_formula_ptgref_row_oob_shrfmla_only_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet = build_shared_formula_ptgref_row_oob_shrfmla_only_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedRefRowOOB_ShrFmlaOnly", &sheet, 1252)
}

fn build_shared_formula_ptgarea_row_oob_shrfmla_only_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet = build_shared_formula_ptgarea_row_oob_shrfmla_only_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedAreaRowOOB_ShrFmlaOnly", &sheet, 1252)
}

fn build_shared_formula_ptgref_col_oob_shrfmla_only_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet = build_shared_formula_ptgref_col_oob_shrfmla_only_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedRefColOOB_ShrFmlaOnly", &sheet, 1252)
}

fn build_shared_formula_ptgarea_col_oob_shrfmla_only_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet = build_shared_formula_ptgarea_col_oob_shrfmla_only_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedAreaColOOB_ShrFmlaOnly", &sheet, 1252)
}

fn build_shared_formula_ptgref_row_oob_shrfmla_only_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula definition stored only in SHRFMLA for range B65535:B65536 (no FORMULA records).
    //
    // The shared rgce uses a PtgRef (not PtgRefN) with row/col-relative flags and a row coordinate
    // at the BIFF8 max (0xFFFF => Excel row 65536). Filling down by 1 shifts the row to 65537,
    // which is out-of-bounds and should materialize as `#REF!`.
    //
    // Expected decoded formulas:
    // - B65535: `A65536+1`
    // - B65536: `#REF!+1`
    const BASE_ROW: u16 = 65_534; // 0-based row for Excel row 65535
    const FOLLOWER_ROW: u16 = 65_535; // 0-based row for Excel row 65536
    const BASE_COL: u16 = 1; // column B

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [BASE_ROW, FOLLOWER_ROW + 1) cols [0, 2) => A65535:B65536.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(BASE_ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(FOLLOWER_ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide one numeric cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(BASE_ROW, 0, xf_cell, 1.0)); // A65535

    // Shared rgce body stored in SHRFMLA. This uses a PtgRef (not PtgRefN) with relative flags so
    // materialization must shift the row.
    let rgce_shared: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x24); // PtgRef
        v.extend_from_slice(&FOLLOWER_ROW.to_le_bytes()); // rw = 65535 (Excel row 65536)
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, row+col relative flags
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(
            BASE_ROW,
            FOLLOWER_ROW,
            BASE_COL as u8,
            BASE_COL as u8,
            &rgce_shared,
        ),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_ptgarea_row_oob_shrfmla_only_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula definition stored only in SHRFMLA for range B65535:B65536 (no FORMULA records).
    //
    // The shared rgce uses a PtgArea (not PtgAreaN) with relative flags on both endpoints.
    // Filling down by 1 shifts the area to A65536:A65537 which is out-of-bounds.
    //
    // Expected decoded formulas:
    // - B65535: `SUM(A65535:A65536)+1`
    // - B65536: `SUM(#REF!)+1`
    const BASE_ROW: u16 = 65_534; // 0-based row for Excel row 65535
    const FOLLOWER_ROW: u16 = 65_535; // 0-based row for Excel row 65536
    const BASE_COL: u16 = 1; // column B
    let col_with_flags: u16 = 0xC000; // col=A + row+col relative flags

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [BASE_ROW, FOLLOWER_ROW + 1) cols [0, 2) => A65535:B65536.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(BASE_ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(FOLLOWER_ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide one numeric cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(BASE_ROW, 0, xf_cell, 1.0)); // A65535

    // Shared rgce body stored in SHRFMLA. This uses a PtgArea (not PtgAreaN) with relative flags so
    // materialization must shift both row indices.
    let rgce_shared: Vec<u8> = {
        let mut v = Vec::new();
        // A65535:A65536
        v.extend_from_slice(&ptg_area(
            BASE_ROW,
            FOLLOWER_ROW,
            col_with_flags,
            col_with_flags,
        ));
        // SUM(...)
        v.push(0x22); // PtgFuncVar
        v.push(1); // argc=1
        v.extend_from_slice(&0x0004u16.to_le_bytes()); // iftab=4 (SUM)
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(
            BASE_ROW,
            FOLLOWER_ROW,
            BASE_COL as u8,
            BASE_COL as u8,
            &rgce_shared,
        ),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_ptgref_col_oob_shrfmla_only_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula definition stored only in SHRFMLA for range A1:B1 (no FORMULA records).
    //
    // The shared rgce uses a PtgRef with the col-relative flag set and an in-bounds column index at
    // the BIFF8 14-bit max (0x3FFF => XFD). Filling right by 1 column produces an out-of-bounds
    // reference (XFE), which Excel represents as `#REF!`.
    //
    // Expected decoded formulas:
    // - A1: `XFD1+1`
    // - B1: `#REF!+1`
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 3) => A1:C1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1 (A..C)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide one numeric cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 2, xf_cell, 1.0)); // C1

    // Shared formula rgce: XFD1 + 1, encoded using PtgRef with row/col-relative flags.
    let rgce_shared: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x24); // PtgRef
        v.extend_from_slice(&0u16.to_le_bytes()); // row = 0 (row 1)
        let col_field = pack_biff8_col_flags(0x3FFF, true, true); // col=XFD + rowRel + colRel
        v.extend_from_slice(&col_field.to_le_bytes());
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 0, 0, 1, &rgce_shared),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

fn build_shared_formula_ptgarea_col_oob_shrfmla_only_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Shared formula definition stored only in SHRFMLA for range A1:B1 (no FORMULA records).
    //
    // The shared rgce uses a PtgArea with col-relative flags set and an endpoint at the BIFF8
    // 14-bit max column (0x3FFF => XFD). Filling right by 1 column shifts the second endpoint to
    // XFE, which is out-of-bounds and should materialize as `#REF!`.
    //
    // Expected decoded formulas:
    // - A1: `SUM(XFC1:XFD1)+1`
    // - B1: `SUM(#REF!)+1`
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 3) => A1:C1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&3u16.to_le_bytes()); // last col + 1 (A..C)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Provide one numeric cell so calamine returns a non-empty range.
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 2, xf_cell, 1.0)); // C1

    // Shared formula rgce: SUM(XFC1:XFD1) + 1, encoded using PtgArea with row/col-relative flags.
    let col_first = pack_biff8_col_flags(0x3FFE, true, true); // XFC
    let col_last = pack_biff8_col_flags(0x3FFF, true, true); // XFD
    let rgce_shared: Vec<u8> = {
        let mut v = Vec::new();
        v.extend_from_slice(&ptg_area(0, 0, col_first, col_last));
        v.push(0x22); // PtgFuncVar
        v.push(1); // argc=1
        v.extend_from_slice(&0x0004u16.to_le_bytes()); // SUM
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };

    push_record(
        &mut sheet,
        RECORD_SHRFMLA,
        &shrfmla_record(0, 0, 0, 1, &rgce_shared),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

/// Build a BIFF8 `.xls` fixture containing a SHRFMLA-only shared formula whose `PtgRef` shifts
/// out-of-bounds in the row direction during materialization.
pub fn build_shared_formula_ptgref_row_oob_shrfmla_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgref_row_oob_shrfmla_only_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a SHRFMLA-only shared formula whose `PtgArea` shifts
/// out-of-bounds in the row direction during materialization.
pub fn build_shared_formula_ptgarea_row_oob_shrfmla_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgarea_row_oob_shrfmla_only_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a SHRFMLA-only shared formula whose `PtgRef` shifts
/// out-of-bounds in the column direction during materialization.
pub fn build_shared_formula_ptgref_col_oob_shrfmla_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgref_col_oob_shrfmla_only_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture containing a SHRFMLA-only shared formula whose `PtgArea` shifts
/// out-of-bounds in the column direction during materialization.
pub fn build_shared_formula_ptgarea_col_oob_shrfmla_only_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgarea_col_oob_shrfmla_only_workbook_stream();

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
/// importer, and a **shared formula** (SHRFMLA + PtgExp) references a *sheet-scoped* defined name
/// via `PtgName`.
///
/// This validates that BIFF-decoded worksheet formulas have access to workbook NAME metadata and
/// render the sheet prefix using the final imported (sanitized) sheet name.
pub fn build_shared_formula_sheet_scoped_name_sanitization_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_sheet_scoped_name_sanitization_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture where the sheet containing the sheet-scoped `PtgName` has an
/// apostrophe (`'`) in its name.
///
/// This validates that sheet-qualified `PtgName` references escape apostrophes correctly
/// (`'O''Brien'!LocalName`).
pub fn build_shared_formula_sheet_scoped_name_apostrophe_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_sheet_scoped_name_apostrophe_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture where the sheet containing the sheet-scoped `PtgName` is named
/// `TRUE`, which Excel allows but requires quoting in formulas to avoid ambiguity with the boolean
/// literal.
pub fn build_shared_formula_sheet_scoped_name_true_sheet_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_sheet_scoped_name_true_sheet_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture where the sheet containing the sheet-scoped `PtgName` is named
/// `A1`, which Excel allows but requires quoting in formulas to avoid ambiguity with an A1-style
/// cell reference.
pub fn build_shared_formula_sheet_scoped_name_a1_sheet_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_sheet_scoped_name_a1_sheet_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture where the sheet containing the sheet-scoped `PtgName` has a
/// non-ASCII (Unicode) name and therefore must be quoted in formulas for `formula-engine` to parse
/// it.
pub fn build_shared_formula_sheet_scoped_name_unicode_sheet_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_sheet_scoped_name_unicode_sheet_workbook_stream();

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

fn build_shared_formula_ptgexp_missing_shrfmla_row_oob_workbook_stream() -> Vec<u8> {
    let xf_cell = 16u16;
    let sheet_stream = build_shared_formula_ptgexp_missing_shrfmla_row_oob_sheet_stream(xf_cell);
    build_single_sheet_workbook_stream("SharedFallback_OOB", &sheet_stream, 1252)
}

fn build_shared_formula_ptgexp_missing_shrfmla_row_oob_sheet_stream(xf_cell: u16) -> Vec<u8> {
    // Malformed shared-formula pattern near the BIFF8 row limit:
    // - Base cell stores a full formula rgce using PtgRef with relative flags.
    // - Follower cell stores only PtgExp referencing the base.
    // - SHRFMLA definition record is intentionally missing.
    //
    // Base: B65535 = A65536+1
    // Follower: B65536 should materialize as #REF!+1 (A65537 is out of BIFF8 bounds).
    const BASE_ROW: u16 = 65_534; // 0-based row for Excel row 65535
    const FOLLOWER_ROW: u16 = 65_535; // 0-based row for Excel row 65536
    const BASE_COL: u16 = 1; // column B

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [BASE_ROW, FOLLOWER_ROW + 1) cols [0, 2) => A65535:B65536.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&(BASE_ROW as u32).to_le_bytes()); // first row
    dims.extend_from_slice(&(FOLLOWER_ROW as u32 + 1).to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col (A)
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1 (A..B)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // Base formula in B65535: `A65536+1` encoded using PtgRef with relative flags.
    let rgce_base: Vec<u8> = vec![
        0x24, // PtgRef
        FOLLOWER_ROW.to_le_bytes()[0],
        FOLLOWER_ROW.to_le_bytes()[1],
        0x00,
        0xC0, // col = A (0) + row_rel + col_rel
        0x1E, // PtgInt
        0x01,
        0x00, // 1
        0x03, // PtgAdd
    ];
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(BASE_ROW, BASE_COL, xf_cell, 0.0, &rgce_base),
    );

    // Follower formula in B65536: `PtgExp(B65535)`.
    let rgce_ptgexp = ptg_exp(BASE_ROW, BASE_COL);
    push_record(
        &mut sheet,
        RECORD_FORMULA,
        &formula_cell(FOLLOWER_ROW, BASE_COL, xf_cell, 0.0, &rgce_ptgexp),
    );

    // Intentionally omit SHRFMLA/ARRAY definition records.

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
}

/// Build a BIFF8 `.xls` fixture containing a malformed shared formula near the BIFF8 row limit:
/// `PtgExp` follower cell references a base `FORMULA.rgce` that uses `PtgRef` with relative flags,
/// but the `SHRFMLA` definition record is missing.
pub fn build_shared_formula_ptgexp_missing_shrfmla_row_oob_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_ptgexp_missing_shrfmla_row_oob_workbook_stream();

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

/// Build a BIFF8 `.xls` fixture where sheet-name sanitization causes a name collision (one sheet
/// must be deduped), and a **shared formula** references a sheet-scoped defined name via `PtgName`.
///
/// The shared formula must resolve the sheet-scoped name to the correct final sheet name (the one
/// corresponding to the `NAME.itab` scope), not the deduped collision sheet.
pub fn build_shared_formula_sheet_scoped_name_dedup_collision_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_shared_formula_sheet_scoped_name_dedup_collision_workbook_stream();

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

pub fn build_shared_formula_sheet_scoped_name_dedup_collision_invalid_second_fixture_xls() -> Vec<u8>
{
    let workbook_stream =
        build_shared_formula_sheet_scoped_name_dedup_collision_invalid_second_workbook_stream();

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

fn hlink_internal_location_biff8_unicode_display(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    location: &str,
    display: &str,
    tooltip: &str,
) -> Vec<u8> {
    // Like `hlink_internal_location`, but encodes the display string as a BIFF8 XLUnicodeString
    // (u16 length + flags) so the importer exercises the fallback string decoding path.
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

    write_hyperlink_string_biff8_unicode(&mut out, display);
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

fn hlink_external_url_len_as_char_count(
    rw_first: u16,
    rw_last: u16,
    col_first: u16,
    col_last: u16,
    url: &str,
    display: &str,
    tooltip: &str,
) -> Vec<u8> {
    // Like `hlink_external_url`, but stores the URL moniker length as a UTF-16 character count
    // rather than a byte count.
    const STREAM_VERSION: u32 = 2;
    const LINK_OPTS_HAS_MONIKER: u32 = 0x0000_0001;
    const LINK_OPTS_HAS_DISPLAY: u32 = 0x0000_0010;
    const LINK_OPTS_HAS_TOOLTIP: u32 = 0x0000_0020;

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

    out.extend_from_slice(&CLSID_URL_MONIKER);
    let mut url_utf16: Vec<u16> = url.encode_utf16().collect();
    url_utf16.push(0); // NUL terminator
                       // Length field stored as code unit count rather than byte length.
    out.extend_from_slice(&(url_utf16.len() as u32).to_le_bytes());
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

fn write_hyperlink_string_biff8_unicode(out: &mut Vec<u8>, s: &str) {
    // Intentionally encode a hyperlink string field as a BIFF8 XLUnicodeString:
    // [cch:u16][flags:u8][chars...]
    //
    // Some producers appear to store hyperlink strings in this form. The importer’s
    // `parse_hyperlink_string` should fall back to BIFF8 string decoding when the u32
    // HyperlinkString layout doesn’t fit.
    let mut u16s: Vec<u16> = s.encode_utf16().collect();
    u16s.push(0);
    let cch: u16 = u16s
        .len()
        .try_into()
        .expect("fixture strings should fit in u16");
    out.extend_from_slice(&cch.to_le_bytes());
    out.push(0x01); // fHighByte=1 (UTF-16LE)
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

fn build_setup_fnopls_fit_to_page_workbook_stream() -> Vec<u8> {
    build_single_sheet_workbook_stream(
        "Sheet1",
        &build_setup_fnopls_fit_to_page_sheet_stream(),
        1252,
    )
}

fn build_setup_fnopls_fit_to_page_sheet_stream() -> Vec<u8> {
    // The workbook globals built by `build_single_sheet_workbook_stream` contain 16 style XFs + 1
    // cell XF (General), so the first usable cell XF index is 16.
    const XF_GENERAL_CELL: u16 = 16;

    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1), cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // WSBOOL: enable FitToPage.
    let wsbool = 0x0C01u16 | WSBOOL_OPTION_FIT_TO_PAGE;
    push_record(&mut sheet, RECORD_WSBOOL, &wsbool.to_le_bytes());

    // SETUP record: fNoPls=1 but fit-to width/height and header/footer margins should still apply.
    let mut setup = Vec::<u8>::new();
    setup.extend_from_slice(&9u16.to_le_bytes()); // iPaperSize (A4) => should be ignored
    setup.extend_from_slice(&80u16.to_le_bytes()); // iScale => should be ignored
    setup.extend_from_slice(&0i16.to_le_bytes()); // iPageStart (ignored)
    setup.extend_from_slice(&2u16.to_le_bytes()); // iFitWidth
    setup.extend_from_slice(&3u16.to_le_bytes()); // iFitHeight
                                                  // grbit: fNoPls=1 and fPortrait=0 (landscape). Paper size / percent scale / orientation should
                                                  // be ignored due to fNoPls, but FitTo + numHdr/numFtr should still be imported.
    setup.extend_from_slice(&0x0004u16.to_le_bytes());
    setup.extend_from_slice(&0u16.to_le_bytes()); // iRes (ignored)
    setup.extend_from_slice(&0u16.to_le_bytes()); // iVRes (ignored)
    setup.extend_from_slice(&0.9f64.to_le_bytes()); // numHdr
    setup.extend_from_slice(&1.1f64.to_le_bytes()); // numFtr
    setup.extend_from_slice(&1u16.to_le_bytes()); // iCopies (ignored)
    push_record(&mut sheet, RECORD_SETUP, &setup);

    // Provide at least one cell so calamine returns a non-empty range.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, XF_GENERAL_CELL, 1.0),
    );

    push_record(&mut sheet, RECORD_EOF, &[]);
    sheet
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

fn build_autofilter_criteria_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteria");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$C$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 2);
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

    // AUTOFILTER record: simple text equality filter on column A for "Alice".
    //
    // Note: do NOT emit FILTERMODE for this fixture; we want to exercise criteria parsing without
    // the separate "filtered rows may not round-trip" warning.
    let doper1 = autofilter_doper_string(AUTOFILTER_OP_EQUAL, "Alice");
    let doper2 = autofilter_doper_none();
    let autofilter = autofilter_record(0, false, &doper1, &doper2);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);

    // Second AUTOFILTER record: numeric BETWEEN filter on column B.
    // Criterion: 10 <= value <= 20.
    let doper1 = autofilter_doper_number(AUTOFILTER_OP_BETWEEN, 10.0);
    let doper2 = autofilter_doper_number(AUTOFILTER_OP_NONE, 20.0);
    let autofilter = autofilter_record(1, false, &doper1, &doper2);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_join_all_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaJoinAll");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$A$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 0);
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

    // DIMENSIONS: rows [0, 5) cols [0, 1) so A1:A5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 1 (A).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &1u16.to_le_bytes());

    // AUTOFILTER record: numeric range filter (A > 10 AND A < 20).
    // This exercises the `AUTOFILTER.grbit` AND join bit.
    let doper1 = autofilter_doper_number(AUTOFILTER_OP_GREATER_THAN, 10.0);
    let doper2 = autofilter_doper_number(AUTOFILTER_OP_LESS_THAN, 20.0);
    let autofilter = autofilter_record(0, true, &doper1, &doper2);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_between_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaBetween");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$B$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 1);
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

    // DIMENSIONS: rows [0, 5) cols [0, 2) so A1:B5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 2 (A..B).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &2u16.to_le_bytes());

    // AUTOFILTER records: BETWEEN / NOT BETWEEN operator codes.
    //
    // Use reversed operands (20, 10) to ensure the importer normalizes min/max.
    //
    // Column A: BETWEEN 10..20.
    let doper1 = autofilter_doper_number(1, 20.0); // op=Between
    let doper2 = autofilter_doper_number(AUTOFILTER_OP_NONE, 10.0);
    let autofilter = autofilter_record(0, false, &doper1, &doper2);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);

    // Column B: NOT BETWEEN 10..20.
    let doper1 = autofilter_doper_number(2, 20.0); // op=NotBetween
    let doper2 = autofilter_doper_number(AUTOFILTER_OP_NONE, 10.0);
    let autofilter = autofilter_record(1, false, &doper1, &doper2);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_blanks_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaBlanks");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$D$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 3);
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

    // DIMENSIONS: rows [0, 5) cols [0, 4) so A1:D5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&4u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 4 (A..D).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &4u16.to_le_bytes());

    // AUTOFILTER records:
    // - Column A: blanks via vt=empty + operator equal.
    // - Column B: non-blanks via vt=empty + operator notEqual.
    // - Column C: blanks via vt=string + empty trailing string.
    // - Column D: non-blanks via vt=string + empty trailing string + notEqual.
    let mut doper1_bytes = [0u8; 8];
    doper1_bytes[0] = AUTOFILTER_VT_EMPTY;
    doper1_bytes[2..4].copy_from_slice(&(AUTOFILTER_OP_EQUAL as u16).to_le_bytes());
    let doper1 = AutoFilterDoper {
        bytes: doper1_bytes,
        trailing_string: None,
    };
    let doper2 = autofilter_doper_none();
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(0, false, &doper1, &doper2),
    );

    let mut doper1_bytes = [0u8; 8];
    doper1_bytes[0] = AUTOFILTER_VT_EMPTY;
    doper1_bytes[2..4].copy_from_slice(&(4u16).to_le_bytes()); // op=NotEqual
    let doper1 = AutoFilterDoper {
        bytes: doper1_bytes,
        trailing_string: None,
    };
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(1, false, &doper1, &doper2),
    );

    let doper1 = autofilter_doper_string(AUTOFILTER_OP_EQUAL, "");
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(2, false, &doper1, &doper2),
    );

    let doper1 = autofilter_doper_string(4, ""); // op=NotEqual
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(3, false, &doper1, &doper2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_text_ops_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaTextOps");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$C$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 2);
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

    // AUTOFILTER records: contains / beginsWith / endsWith.
    // These are surfaced as `FilterCriterion::OpaqueCustom`.
    let doper2 = autofilter_doper_none();
    let doper1 = autofilter_doper_string(9, "Al"); // op=Contains
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(0, false, &doper1, &doper2),
    );
    let doper1 = autofilter_doper_string(10, "B"); // op=BeginsWith
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(1, false, &doper1, &doper2),
    );
    let doper1 = autofilter_doper_string(11, "z"); // op=EndsWith
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(2, false, &doper1, &doper2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_text_ops_negative_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaTextOpsNeg");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$C$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 2);
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

    // AUTOFILTER records: doesNotContain / doesNotBeginWith / doesNotEndWith.
    // These are surfaced as `FilterCriterion::OpaqueCustom`.
    let doper2 = autofilter_doper_none();
    let doper1 = autofilter_doper_string(12, "Al"); // op=DoesNotContain
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(0, false, &doper1, &doper2),
    );

    // Encode op=DoesNotBeginWith using the "operator stored in second byte" variant to exercise
    // that importer path.
    let mut bytes = [0u8; 8];
    bytes[0] = AUTOFILTER_VT_STRING;
    bytes[1] = 13; // op=DoesNotBeginWith
    let doper1 = AutoFilterDoper {
        bytes,
        trailing_string: Some("B".to_string()),
    };
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(1, false, &doper1, &doper2),
    );

    let doper1 = autofilter_doper_string(14, "z"); // op=DoesNotEndWith
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(2, false, &doper1, &doper2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_bool_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaBool");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$B$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 1);
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

    // DIMENSIONS: rows [0, 5) cols [0, 2) so A1:B5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&2u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 2 (A..B).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &2u16.to_le_bytes());

    // AUTOFILTER records: boolean equality filters.
    let doper2 = autofilter_doper_none();
    let doper1 = autofilter_doper_bool(AUTOFILTER_OP_EQUAL, true);
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(0, false, &doper1, &doper2),
    );
    let doper1 = autofilter_doper_bool(AUTOFILTER_OP_EQUAL, false);
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(1, false, &doper1, &doper2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_top10_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaTop10");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$A$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 0);
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

    // DIMENSIONS: rows [0, 5) cols [0, 1) so A1:A5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 1 (A).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &1u16.to_le_bytes());

    // AUTOFILTER record: Top10 filter.
    // BIFF grbit flags: fTop10 (0x0008), fTop (0x0010), fPercent (0x0020).
    let grbit_top10: u16 = 0x0008 | 0x0010 | 0x0020;
    let doper1 = autofilter_doper_number(AUTOFILTER_OP_NONE, 5.0);
    let doper2 = autofilter_doper_none();
    let mut autofilter = Vec::new();
    autofilter.extend_from_slice(&0u16.to_le_bytes()); // iEntry (A)
    autofilter.extend_from_slice(&grbit_top10.to_le_bytes());
    autofilter.extend_from_slice(&doper1.bytes);
    autofilter.extend_from_slice(&doper2.bytes);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_operator_byte1_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaOpByte1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$A$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 0);
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

    // DIMENSIONS: rows [0, 5) cols [0, 1) so A1:A5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: a single General cell so calamine populates a range for the sheet.
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 0, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 1 (A).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &1u16.to_le_bytes());

    // AUTOFILTER record: equals "Alice", but with the operator code stored in the byte1/grbit
    // field instead of `wOper` (which is left as 0). This exercises the parser's operator fallback.
    let mut doper1_bytes = [0u8; 8];
    doper1_bytes[0] = AUTOFILTER_VT_STRING;
    doper1_bytes[1] = AUTOFILTER_OP_EQUAL;
    let doper1 = AutoFilterDoper {
        bytes: doper1_bytes,
        trailing_string: Some("Alice".to_string()),
    };
    let doper2 = autofilter_doper_none();
    push_record(
        &mut sheet,
        RECORD_AUTOFILTER,
        &autofilter_record(0, false, &doper1, &doper2),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_continued_string_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaContinue");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`): $A$1:$C$5.
    let filter_db_rgce = ptg_area(0, 4, 0, 2);
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

    // AUTOFILTER record with a trailing string ("Alice") split across a CONTINUE record.
    //
    // Split after the XLUnicodeString header + first 2 characters ("Al") so the importer must read
    // the continuation flag byte and remaining character bytes from the CONTINUE fragment.
    let doper1 = autofilter_doper_string(AUTOFILTER_OP_EQUAL, "Alice");
    let doper2 = autofilter_doper_none();
    let full = autofilter_record(0, false, &doper1, &doper2);

    const FIXED_PREFIX_LEN: usize = 20; // iEntry+grbit+DOPER1+DOPER2
    let split_at = (FIXED_PREFIX_LEN + 3 + 2).min(full.len());
    let first = &full[..split_at];
    let rest = &full[split_at..];

    push_record(&mut sheet, RECORD_AUTOFILTER, first);
    let mut cont = Vec::<u8>::new();
    cont.push(0); // continued segment flags (compressed)
    cont.extend_from_slice(rest);
    push_record(&mut sheet, RECORD_CONTINUE, &cont);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_criteria_absolute_entry_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterCriteriaAbsEntry");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    // `_xlnm._FilterDatabase` (built-in name id 0x0D) scoped to the sheet (`itab=1`).
    // Range: $D$1:$F$5.
    let filter_db_rgce = ptg_area(0, 4, 3, 5);
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

    // DIMENSIONS: rows [0, 5) cols [3, 6) so D1:F5 exists.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&5u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&3u16.to_le_bytes()); // first col (D)
    dims.extend_from_slice(&6u16.to_le_bytes()); // last col + 1 (F)
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // D1: seed calamine's range (use a General NUMBER cell).
    push_record(
        &mut sheet,
        RECORD_NUMBER,
        &number_cell(0, 3, xf_general, 1.0),
    );

    // AUTOFILTERINFO: cEntries = 3 (D..F).
    push_record(&mut sheet, RECORD_AUTOFILTERINFO, &3u16.to_le_bytes());
    // FILTERMODE: present (no payload) to indicate an active filter.
    push_record(&mut sheet, RECORD_FILTERMODE, &[]);

    // AUTOFILTER records: encode iEntry as an *absolute* worksheet column index.
    // Column D (col=3): equals "Alice".
    let doper1 = autofilter_doper_string(AUTOFILTER_OP_EQUAL, "Alice");
    let doper2 = autofilter_doper_none();
    let autofilter = autofilter_record(3, false, &doper1, &doper2);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);

    // Column F (col=5): numeric comparison > 1.
    let doper1 = autofilter_doper_number(AUTOFILTER_OP_GREATER_THAN, 1.0);
    let doper2 = autofilter_doper_none();
    let autofilter = autofilter_record(5, false, &doper1, &doper2);
    push_record(&mut sheet, RECORD_AUTOFILTER, &autofilter);

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

fn build_autofilter_sort12_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterSort12");
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

    // Sort12/SortData12 future records: sort the filtered range A1:C5 by column B descending,
    // with a header row.
    let sort_payload = sort_record_payload(
        0,
        4, // rows (rwFirst..rwLast) => A1:C5
        0,
        2,               // cols (colFirst..colLast)
        true,            // has header row
        &[(1u16, true)], // key: column B descending
    );

    let mut sort12 = Vec::<u8>::new();
    sort12.extend_from_slice(&RECORD_SORT12.to_le_bytes()); // FrtHeader.rt
    sort12.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    sort12.extend_from_slice(&0u32.to_le_bytes()); // reserved
    sort12.extend_from_slice(&sort_payload);
    push_record(&mut sheet, RECORD_SORT12, &sort12);

    let mut sortdata12 = Vec::<u8>::new();
    sortdata12.extend_from_slice(&RECORD_SORTDATA12.to_le_bytes()); // FrtHeader.rt
    sortdata12.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    sortdata12.extend_from_slice(&0u32.to_le_bytes()); // reserved
    sortdata12.extend_from_slice(&sort_payload);
    push_record(&mut sheet, RECORD_SORTDATA12, &sortdata12);

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet

    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_autofilter_sort12_continuefrt12_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "FilterSort12Cont");
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

    let sort_payload = sort_record_payload(
        0,
        4, // rows (rwFirst..rwLast) => A1:C5
        0,
        2,               // cols (colFirst..colLast)
        true,            // has header row
        &[(1u16, true)], // key: column B descending
    );

    // Split the Sort12 payload (after `FrtHeader`) across two records to force a ContinueFrt12
    // continuation path.
    let split_at = 12usize.min(sort_payload.len());

    let mut sort12 = Vec::<u8>::new();
    sort12.extend_from_slice(&RECORD_SORT12.to_le_bytes()); // FrtHeader.rt
    sort12.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    sort12.extend_from_slice(&0u32.to_le_bytes()); // reserved
    sort12.extend_from_slice(&sort_payload[..split_at]);
    push_record(&mut sheet, RECORD_SORT12, &sort12);

    let mut cont = Vec::<u8>::new();
    cont.extend_from_slice(&RECORD_CONTINUEFRT12.to_le_bytes()); // FrtHeader.rt
    cont.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    cont.extend_from_slice(&0u32.to_le_bytes()); // reserved
    cont.extend_from_slice(&sort_payload[split_at..]);
    push_record(&mut sheet, RECORD_CONTINUEFRT12, &cont);

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

fn build_autofilter12_continuefrt12_workbook_stream() -> Vec<u8> {
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
    write_short_unicode_string(&mut boundsheet, "Filter12Cont");
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

    // AutoFilter12 record split across AutoFilter12 + ContinueFrt12.
    //
    // Layout (best-effort):
    //   AutoFilter12:
    //     FrtHeader (8 bytes): rt, grbitFrt, reserved
    //     colId (u16)
    //     cVals (u16)
    //     first XLUnicodeString
    //   ContinueFrt12:
    //     FrtHeader (8 bytes): rt=ContinueFrt12, grbitFrt, reserved
    //     remaining XLUnicodeString bytes
    let mut af12 = Vec::<u8>::new();
    af12.extend_from_slice(&RECORD_AUTOFILTER12.to_le_bytes()); // FrtHeader.rt
    af12.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    af12.extend_from_slice(&0u32.to_le_bytes()); // reserved
    af12.extend_from_slice(&0u16.to_le_bytes()); // colId
    af12.extend_from_slice(&2u16.to_le_bytes()); // cVals
    write_unicode_string(&mut af12, "Alice");
    push_record(&mut sheet, RECORD_AUTOFILTER12, &af12);

    let mut cont = Vec::<u8>::new();
    cont.extend_from_slice(&RECORD_CONTINUEFRT12.to_le_bytes()); // FrtHeader.rt
    cont.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    cont.extend_from_slice(&0u32.to_le_bytes()); // reserved
    write_unicode_string(&mut cont, "Bob");
    push_record(&mut sheet, RECORD_CONTINUEFRT12, &cont);

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

fn array_record_refu(
    rw_first: u16,
    rw_last: u16,
    col_first: u8,
    col_last: u8,
    rgce: &[u8],
) -> Vec<u8> {
    // ARRAY record payload (best-effort BIFF8 encoding):
    //   [ref: RefU (6 bytes)]
    //   [flags/reserved: u16] (commonly present; importer is permissive about its presence)
    //   [cce: u16]
    //   [rgce: cce bytes]
    //
    // RefU: [rwFirst:u16][rwLast:u16][colFirst:u8][colLast:u8]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.push(col_first);
    out.push(col_last);
    out.extend_from_slice(&0u16.to_le_bytes()); // flags/reserved
    out.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
    out.extend_from_slice(rgce);
    out
}

fn doper(vt: u8, grbit: u8, w_oper: u16, value_raw: u32) -> [u8; 8] {
    // DOPER (best-effort BIFF8 encoding) [MS-XLS 2.5.69]:
    // [vt:u8][grbit:u8][wOper:u16 LE][value_raw:u32 LE]
    let mut out = [0u8; 8];
    out[0] = vt;
    out[1] = grbit;
    out[2..4].copy_from_slice(&w_oper.to_le_bytes());
    out[4..8].copy_from_slice(&value_raw.to_le_bytes());
    out
}

fn autofilter_record_payload(
    i_entry: u16,
    grbit: u16,
    doper1: [u8; 8],
    doper2: [u8; 8],
    strings: &[&str],
) -> Vec<u8> {
    // AUTOFILTER (worksheet substream) payload [MS-XLS 2.4.31] (best-effort):
    // [iEntry:u16][grbit:u16][DOPER1:8][DOPER2:8][optional XLUnicodeString payloads...]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&i_entry.to_le_bytes());
    out.extend_from_slice(&grbit.to_le_bytes());
    out.extend_from_slice(&doper1);
    out.extend_from_slice(&doper2);
    for s in strings {
        write_unicode_string(&mut out, s);
    }
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

fn labelsst_cell(row: u16, col: u16, xf: u16, sst_index: u32) -> [u8; 10] {
    let mut out = [0u8; 10];
    out[0..2].copy_from_slice(&row.to_le_bytes());
    out[2..4].copy_from_slice(&col.to_le_bytes());
    out[4..6].copy_from_slice(&xf.to_le_bytes());
    out[6..10].copy_from_slice(&sst_index.to_le_bytes());
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

#[derive(Debug, Clone)]
struct AutoFilterDoper {
    bytes: [u8; 8],
    trailing_string: Option<String>,
}

fn autofilter_record(
    col: u16,
    join_all: bool,
    doper1: &AutoFilterDoper,
    doper2: &AutoFilterDoper,
) -> Vec<u8> {
    // AUTOFILTER record payload [MS-XLS 2.4.31], minimal encoding compatible with
    // `biff/autofilter_criteria.rs`:
    // - iEntry (u16): 0-based column index in the sheet (not offset within the filter range)
    // - grbit (u16): flags (bit 0 indicates AND)
    // - DOPER1, DOPER2: 8 bytes each
    // - optional trailing XLUnicodeString payloads for string DOPER values
    let mut out = Vec::new();
    out.extend_from_slice(&col.to_le_bytes());
    let mut grbit = 0u16;
    if join_all {
        grbit |= AUTOFILTER_GRBIT_AND;
    }
    out.extend_from_slice(&grbit.to_le_bytes());
    out.extend_from_slice(&doper1.bytes);
    out.extend_from_slice(&doper2.bytes);
    if let Some(s) = doper1.trailing_string.as_ref() {
        write_unicode_string(&mut out, s);
    }
    if let Some(s) = doper2.trailing_string.as_ref() {
        write_unicode_string(&mut out, s);
    }
    out
}

fn autofilter_doper_none() -> AutoFilterDoper {
    // DOPER with op=None (no criterion).
    let mut bytes = [0u8; 8];
    bytes[0] = AUTOFILTER_VT_EMPTY;
    bytes[2..4].copy_from_slice(&(AUTOFILTER_OP_NONE as u16).to_le_bytes());
    AutoFilterDoper {
        bytes,
        trailing_string: None,
    }
}

fn autofilter_doper_number(operator: u8, value: f64) -> AutoFilterDoper {
    let mut bytes = [0u8; 8];
    bytes[0] = AUTOFILTER_VT_NUMBER;
    bytes[2..4].copy_from_slice(&(operator as u16).to_le_bytes());
    bytes[4..8].copy_from_slice(&rk_number(value).to_le_bytes());
    AutoFilterDoper {
        bytes,
        trailing_string: None,
    }
}

fn autofilter_doper_bool(operator: u8, value: bool) -> AutoFilterDoper {
    let mut bytes = [0u8; 8];
    bytes[0] = AUTOFILTER_VT_BOOL;
    bytes[2..4].copy_from_slice(&(operator as u16).to_le_bytes());
    let raw: u32 = if value { 1 } else { 0 };
    bytes[4..8].copy_from_slice(&raw.to_le_bytes());
    AutoFilterDoper {
        bytes,
        trailing_string: None,
    }
}

fn autofilter_doper_string(operator: u8, value: &str) -> AutoFilterDoper {
    let mut bytes = [0u8; 8];
    bytes[0] = AUTOFILTER_VT_STRING;
    bytes[2..4].copy_from_slice(&(operator as u16).to_le_bytes());
    AutoFilterDoper {
        bytes,
        trailing_string: Some(value.to_string()),
    }
}

fn rk_number(value: f64) -> u32 {
    // Best-effort RK encoding used by BIFF DOPER values.
    //
    // For integers in the signed 30-bit range, use the integer RK form. Otherwise, store the high
    // 30 bits of the f64 and drop the remaining bits (matching the lossy RK float form).
    if value.fract() == 0.0 && value.is_finite() {
        let i = value as i32;
        if i >= -0x2000_0000 && i <= 0x1FFF_FFFF {
            return ((i as u32) << 2) | 0x02;
        }
    }

    let bits = value.to_bits();
    let high30 = (bits >> 34) as u32;
    high30 << 2
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
    formula_cell_with_grbit(row, col, xf, cached_result, 0, rgce)
}

fn formula_cell_with_grbit(
    row: u16,
    col: u16,
    xf: u16,
    cached_result: f64,
    grbit: u16,
    rgce: &[u8],
) -> Vec<u8> {
    // FORMULA record payload (BIFF8) [MS-XLS 2.4.127].
    //
    // This is a minimal encoding sufficient for calamine to surface the formula text, while
    // allowing callers to specify the raw `grbit` flags (e.g. shared formulas).
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    out.extend_from_slice(&xf.to_le_bytes());
    out.extend_from_slice(&cached_result.to_le_bytes()); // cached formula result (IEEE f64)
    out.extend_from_slice(&grbit.to_le_bytes()); // grbit
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
    // Layout (RefU + cUse + cce + rgce):
    //   [rwFirst: u16][rwLast: u16][colFirst: u8][colLast: u8]  // RefU range
    //   [cUse: u16]
    //   [cce: u16][rgce bytes]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&rw_first.to_le_bytes());
    out.extend_from_slice(&rw_last.to_le_bytes());
    out.push(col_first);
    out.push(col_last);
    out.extend_from_slice(&0u16.to_le_bytes()); // cUse
    out.extend_from_slice(&(rgce.len() as u16).to_le_bytes());
    out.extend_from_slice(rgce);
    out
}

fn shrfmla_record_with_rgcb(
    rw_first: u16,
    rw_last: u16,
    col_first: u8,
    col_last: u8,
    rgce: &[u8],
    rgcb: &[u8],
) -> Vec<u8> {
    // SHRFMLA record payload (BIFF8) [MS-XLS 2.4.277].
    //
    // Some ptgs (notably `PtgArray`) reference additional data blocks serialized after the rgce
    // token stream. BIFF8 stores these blocks as trailing bytes inside the same record (commonly
    // referred to as `rgcb`).
    let mut out = shrfmla_record(rw_first, rw_last, col_first, col_last, rgce);
    out.extend_from_slice(rgcb);
    out
}

fn rgcb_array_constant_numbers_2x2(values: &[f64; 4]) -> Vec<u8> {
    // Serialize a 2x2 numeric array constant payload for BIFF8 `rgcb` trailing bytes.
    //
    // BIFF8/BIFF12 use a similar array-constant structure. We write a minimal subset that matches
    // what our BIFF8 array-constant decoder expects:
    //   [cols_minus1: u16][rows_minus1: u16]
    // followed by row-major elements, each encoded as:
    //   [xltypeNum: 0x01][f64]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&1u16.to_le_bytes()); // cols_minus1 (2 cols)
    out.extend_from_slice(&1u16.to_le_bytes()); // rows_minus1 (2 rows)
    for v in *values {
        out.push(0x01); // xltypeNum
        out.extend_from_slice(&v.to_le_bytes());
    }
    out
}

fn rgcb_array_constant_string_1x1(value: &str) -> Vec<u8> {
    // Serialize a 1x1 string array constant payload for BIFF8 `rgcb` trailing bytes.
    //
    // This matches the simplified structure expected by our BIFF8 array-constant decoder:
    //   [cols_minus1: u16][rows_minus1: u16]
    // followed by a single element encoded as:
    //   [xltypeStr: 0x02][cch: u16][utf16 chars...]
    let units: Vec<u16> = value.encode_utf16().collect();
    let cch: u16 = units.len().try_into().unwrap_or(u16::MAX);
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1 (1 col)
    out.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1 (1 row)
    out.push(0x02); // xltypeStr
    out.extend_from_slice(&cch.to_le_bytes());
    for unit in units {
        out.extend_from_slice(&unit.to_le_bytes());
    }
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

fn formula_cell_with_rgcb(
    row: u16,
    col: u16,
    xf: u16,
    cached_result: f64,
    rgce: &[u8],
    rgcb: &[u8],
) -> Vec<u8> {
    let mut out = formula_cell(row, col, xf, cached_result, rgce);
    out.extend_from_slice(rgcb);
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

fn feat_hdr_record_sheet_protection_allow_mask_prefixed(allow_mask: u16) -> Vec<u8> {
    // Like `feat_hdr_record_sheet_protection_allow_mask`, but stores the allow mask after a 2-byte
    // prefix in the header data so parsers must scan for it.
    const ISF_SHEET_PROTECTION: u16 = 0x0002;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&RECORD_FEATHEADR.to_le_bytes()); // rt
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved

    out.extend_from_slice(&ISF_SHEET_PROTECTION.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out.extend_from_slice(&4u32.to_le_bytes()); // cbHdrData
    out.extend_from_slice(&[0xFF, 0xFF]); // prefix bytes
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

fn build_sst_phonetic_workbook_stream(phonetic_text: &str) -> Vec<u8> {
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // Minimal XF table (16 style XFs + 1 cell XF).
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

    // SST with one string that carries an ExtRst phonetic block.
    let sst = sst_record_single_string_with_phonetic("Base", phonetic_text);
    push_record(&mut globals, RECORD_SST, &sst);

    push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

    // -- Sheet -------------------------------------------------------------------
    let sheet_offset = globals.len();
    let sheet = build_sst_phonetic_sheet_stream(xf_general);

    // Patch BoundSheet offset.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn build_sst_phonetic_sheet_stream(xf_general: u16) -> Vec<u8> {
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET)); // BOF worksheet

    // DIMENSIONS: rows [0, 1) cols [0, 1)
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);

    push_record(&mut sheet, RECORD_WINDOW2, &window2());

    // A1: LABELSST referencing the first SST entry (index 0).
    push_record(
        &mut sheet,
        RECORD_LABELSST,
        &labelsst_cell(0, 0, xf_general, 0),
    );

    push_record(&mut sheet, RECORD_EOF, &[]); // EOF worksheet
    sheet
}

fn sst_record_single_string_with_phonetic(base_text: &str, phonetic_text: &str) -> Vec<u8> {
    // SST record payload [MS-XLS 2.4.261]:
    //   [cstTotal:u32][cstUnique:u32][rgb: XLUnicodeRichExtendedString[]]
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&1u32.to_le_bytes()); // cstTotal
    out.extend_from_slice(&1u32.to_le_bytes()); // cstUnique

    // ExtRst: TLV stream containing a single phonetic block.
    let mut phonetic_string = Vec::<u8>::new();
    write_unicode_string(&mut phonetic_string, phonetic_text);

    let mut ext = Vec::<u8>::new();
    ext.extend_from_slice(&EXT_RST_TYPE_PHONETIC.to_le_bytes());
    ext.extend_from_slice(&(phonetic_string.len() as u16).to_le_bytes());
    ext.extend_from_slice(&phonetic_string);

    // XLUnicodeRichExtendedString [MS-XLS 2.5.296]:
    //   [cch:u16][flags:u8][cbExtRst:u32][chars...][ExtRst bytes...]
    let utf16: Vec<u16> = base_text.encode_utf16().collect();
    let cch: u16 = utf16
        .len()
        .try_into()
        .expect("base string too long for u16 length");

    // Use compressed 8-bit storage for ASCII base strings.
    // flags: fHighByte=0, fExtSt=1.
    out.extend_from_slice(&cch.to_le_bytes());
    out.push(0x04); // STR_FLAG_EXT
    out.extend_from_slice(&(ext.len() as u32).to_le_bytes());

    // Character bytes (compressed).
    out.extend_from_slice(base_text.as_bytes());
    out.extend_from_slice(&ext);
    out
}

// -- Test-only BIFF8 RC4 CryptoAPI encryption helpers ---------------------------------------------
//
// These helpers are intentionally minimal and only cover the subset needed to generate encrypted
// fixtures for integration tests. They are *not* a general `.xls` encryption implementation.
//
// The decryptor lives in `crates/formula-xls/src/decrypt.rs`.
const ENCRYPTION_TYPE_RC4: u16 = 0x0001;
const ENCRYPTION_SUBTYPE_CRYPTOAPI: u16 = 0x0002;

const CALG_RC4: u32 = 0x0000_6801;
const CALG_SHA1: u32 = 0x0000_8004;

const PAYLOAD_BLOCK_SIZE: usize = 1024;
const PASSWORD_HASH_ITERATIONS: u32 = 50_000;

fn utf16le_bytes(s: &str) -> Vec<u8> {
    s.encode_utf16().flat_map(|c| c.to_le_bytes()).collect()
}

fn sha1_bytes(chunks: &[&[u8]]) -> [u8; 20] {
    let mut hasher = Sha1::new();
    for chunk in chunks {
        hasher.update(chunk);
    }
    let digest = hasher.finalize();
    let mut out = [0u8; 20];
    out.copy_from_slice(&digest);
    out
}

fn derive_key_material(password: &str, salt: &[u8]) -> [u8; 20] {
    // Match `crates/formula-xls/src/decrypt.rs` (CryptoAPI RC4):
    //   H0 = SHA1(salt || UTF16LE(password))
    //   for i in 0..49999: H0 = SHA1(i_le32 || H0)
    let pw_bytes = utf16le_bytes(password);
    let mut hash = sha1_bytes(&[salt, &pw_bytes]);
    for i in 0..PASSWORD_HASH_ITERATIONS {
        let iter = i.to_le_bytes();
        hash = sha1_bytes(&[&iter, &hash]);
    }
    hash
}

fn derive_block_key(key_material: &[u8; 20], block: u32, key_len: usize) -> Vec<u8> {
    let block_bytes = block.to_le_bytes();
    let digest = sha1_bytes(&[key_material, &block_bytes]);
    digest[..key_len].to_vec()
}

#[derive(Debug, Clone)]
struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        let mut s = [0u8; 256];
        for (i, v) in s.iter_mut().enumerate() {
            *v = i as u8;
        }

        let mut j: u8 = 0;
        for i in 0..256usize {
            j = j.wrapping_add(s[i]).wrapping_add(key[i % key.len()]);
            s.swap(i, j as usize);
        }

        Self { s, i: 0, j: 0 }
    }

    fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data.iter_mut() {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let t = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
            let k = self.s[t as usize];
            *b ^= k;
        }
    }
}

struct PayloadRc4 {
    key_material: [u8; 20],
    key_len: usize,
    block: u32,
    pos_in_block: usize,
    rc4: Rc4,
}

impl PayloadRc4 {
    fn new(key_material: [u8; 20], key_len: usize) -> Self {
        let key = derive_block_key(&key_material, 0, key_len);
        let rc4 = Rc4::new(&key);
        Self {
            key_material,
            key_len,
            block: 0,
            pos_in_block: 0,
            rc4,
        }
    }

    fn rekey(&mut self) {
        self.block = self.block.wrapping_add(1);
        let key = derive_block_key(&self.key_material, self.block, self.key_len);
        self.rc4 = Rc4::new(&key);
        self.pos_in_block = 0;
    }

    fn apply_keystream(&mut self, mut data: &mut [u8]) {
        while !data.is_empty() {
            if self.pos_in_block == PAYLOAD_BLOCK_SIZE {
                self.rekey();
            }

            let remaining_in_block = PAYLOAD_BLOCK_SIZE.saturating_sub(self.pos_in_block);
            let chunk_len = data.len().min(remaining_in_block);
            let (chunk, rest) = data.split_at_mut(chunk_len);
            self.rc4.apply_keystream(chunk);
            self.pos_in_block += chunk_len;
            data = rest;
        }
    }
}

struct CryptoApiFilepassRecord {
    record_bytes: Vec<u8>,
    key_material: [u8; 20],
    key_len: usize,
}

fn build_filepass_record_rc4_cryptoapi(password: &str) -> CryptoApiFilepassRecord {
    // Deterministic salt/verifier so the generated fixture is stable and diffable.
    let salt: [u8; 16] = [
        0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD, 0xAE,
        0xAF,
    ];
    let verifier: [u8; 16] = [
        0xF0, 0xE1, 0xD2, 0xC3, 0xB4, 0xA5, 0x96, 0x87, 0x78, 0x69, 0x5A, 0x4B, 0x3C, 0x2D, 0x1E,
        0x0F,
    ];

    let verifier_hash = sha1_bytes(&[&verifier]);

    let key_size_bits: u32 = 128;
    let key_len: usize = (key_size_bits / 8) as usize;

    let key_material = derive_key_material(password, &salt);
    let key0 = derive_block_key(&key_material, 0, key_len);
    let mut rc4 = Rc4::new(&key0);

    let mut encrypted_verifier = verifier;
    rc4.apply_keystream(&mut encrypted_verifier);
    let mut encrypted_verifier_hash = verifier_hash;
    rc4.apply_keystream(&mut encrypted_verifier_hash);

    // Build an [MS-OFFCRYPTO] Standard `EncryptionInfo` payload that matches the subset parsed by
    // `formula-xls`'s RC4 CryptoAPI decryptor.
    //
    // EncryptionHeader (fixed 8 DWORDs, no CSPName).
    let mut encryption_header = Vec::with_capacity(32);
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // Flags
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    encryption_header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
    encryption_header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // AlgIDHash
    encryption_header.extend_from_slice(&key_size_bits.to_le_bytes()); // KeySize
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // ProviderType (ignored by decryptor)
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2

    // EncryptionVerifier.
    let mut encryption_verifier = Vec::new();
    encryption_verifier.extend_from_slice(&(salt.len() as u32).to_le_bytes()); // SaltSize
    encryption_verifier.extend_from_slice(&salt);
    encryption_verifier.extend_from_slice(&encrypted_verifier);
    encryption_verifier.extend_from_slice(&(encrypted_verifier_hash.len() as u32).to_le_bytes()); // VerifierHashSize
    encryption_verifier.extend_from_slice(&encrypted_verifier_hash);

    // EncryptionInfo.
    let header_size = encryption_header.len() as u32;
    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // MajorVersion (opaque to decryptor)
    encryption_info.extend_from_slice(&2u16.to_le_bytes()); // MinorVersion (opaque to decryptor)
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // Flags
    encryption_info.extend_from_slice(&header_size.to_le_bytes()); // HeaderSize
    encryption_info.extend_from_slice(&encryption_header);
    encryption_info.extend_from_slice(&encryption_verifier);

    // FILEPASS payload (BIFF8 RC4 CryptoAPI).
    let mut filepass_payload = Vec::new();
    filepass_payload.extend_from_slice(&ENCRYPTION_TYPE_RC4.to_le_bytes());
    filepass_payload.extend_from_slice(&ENCRYPTION_SUBTYPE_CRYPTOAPI.to_le_bytes());
    filepass_payload.extend_from_slice(&(encryption_info.len() as u32).to_le_bytes());
    filepass_payload.extend_from_slice(&encryption_info);

    assert!(
        filepass_payload.len() <= u16::MAX as usize,
        "FILEPASS payload too large for BIFF record: len={}",
        filepass_payload.len()
    );

    let mut record = Vec::with_capacity(4 + filepass_payload.len());
    record.extend_from_slice(&RECORD_FILEPASS.to_le_bytes());
    record.extend_from_slice(&(filepass_payload.len() as u16).to_le_bytes());
    record.extend_from_slice(&filepass_payload);

    CryptoApiFilepassRecord {
        record_bytes: record,
        key_material,
        key_len,
    }
}

fn patch_boundsheet_offsets_for_inserted_prefix(workbook_stream: &mut [u8], delta: u32) {
    // BoundSheet8 record payload begins with `lbPlyPos` (u32), which is an absolute offset from the
    // start of the workbook stream. When we insert a FILEPASS record near the start of the stream,
    // all sheet substreams shift by `delta`.
    let mut offset = 0usize;
    while offset + 4 <= workbook_stream.len() {
        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start + len;
        assert!(
            data_end <= workbook_stream.len(),
            "truncated record while patching boundsheets"
        );

        if record_id == RECORD_BOUNDSHEET {
            assert!(len >= 4, "BOUNDSHEET record payload too short (len={len})");
            let lb_ply_pos = u32::from_le_bytes([
                workbook_stream[data_start],
                workbook_stream[data_start + 1],
                workbook_stream[data_start + 2],
                workbook_stream[data_start + 3],
            ]);
            let patched = lb_ply_pos
                .checked_add(delta)
                .expect("lbPlyPos overflow while patching boundsheets");
            workbook_stream[data_start..data_start + 4].copy_from_slice(&patched.to_le_bytes());
        }

        if record_id == RECORD_EOF {
            break;
        }
        offset = data_end;
    }
}

fn encrypt_biff8_workbook_stream_rc4_cryptoapi(workbook_stream: &[u8], password: &str) -> Vec<u8> {
    // Insert FILEPASS after the workbook globals BOF record.
    assert!(workbook_stream.len() >= 4, "workbook stream too short");
    let record_id = u16::from_le_bytes([workbook_stream[0], workbook_stream[1]]);
    assert_eq!(
        record_id, RECORD_BOF,
        "expected workbook stream to start with BOF"
    );
    let bof_len = u16::from_le_bytes([workbook_stream[2], workbook_stream[3]]) as usize;
    let bof_end = 4 + bof_len;
    assert!(bof_end <= workbook_stream.len(), "truncated BOF record");

    let filepass = build_filepass_record_rc4_cryptoapi(password);
    let delta = filepass.record_bytes.len() as u32;

    let mut out = Vec::with_capacity(workbook_stream.len() + filepass.record_bytes.len());
    out.extend_from_slice(&workbook_stream[..bof_end]);
    let filepass_offset = out.len();
    out.extend_from_slice(&filepass.record_bytes);
    out.extend_from_slice(&workbook_stream[bof_end..]);

    patch_boundsheet_offsets_for_inserted_prefix(&mut out, delta);

    // Encrypt record payload bytes after FILEPASS using the same record-payload-only stream model
    // used by `decrypt_biff8_workbook_stream_rc4_cryptoapi`.
    let mut cipher = PayloadRc4::new(filepass.key_material, filepass.key_len);
    let mut offset = filepass_offset + filepass.record_bytes.len();
    while offset < out.len() {
        assert!(
            offset + 4 <= out.len(),
            "truncated BIFF record header while encrypting"
        );
        let len = u16::from_le_bytes([out[offset + 2], out[offset + 3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start + len;
        assert!(
            data_end <= out.len(),
            "truncated BIFF record payload while encrypting"
        );

        cipher.apply_keystream(&mut out[data_start..data_end]);
        offset = data_end;
    }

    out
}
