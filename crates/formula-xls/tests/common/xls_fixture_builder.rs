#![allow(dead_code)]

use std::io::{Cursor, Write};

// This fixture builder writes just enough BIFF8 to exercise the importer. Keep record ids and
// commonly-used BIFF constants named so the intent stays readable.
const RECORD_BOF: u16 = 0x0809;
const RECORD_EOF: u16 = 0x000A;
const RECORD_CODEPAGE: u16 = 0x0042;
const RECORD_DATEMODE: u16 = 0x0022;
const RECORD_WINDOW1: u16 = 0x003D;
const RECORD_FONT: u16 = 0x0031;
const RECORD_CALCCOUNT: u16 = 0x000C;
const RECORD_CALCMODE: u16 = 0x000D;
const RECORD_PRECISION: u16 = 0x000E;
const RECORD_DELTA: u16 = 0x0010;
const RECORD_ITERATION: u16 = 0x0011;
const RECORD_FORMAT: u16 = 0x041E;
const RECORD_CONTINUE: u16 = 0x003C;
const RECORD_XF: u16 = 0x00E0;
const RECORD_BOUNDSHEET: u16 = 0x0085;
const RECORD_SAVERECALC: u16 = 0x005F;
const RECORD_SHEETEXT: u16 = 0x0862;
const RECORD_WINDOW2: u16 = 0x023E;
const RECORD_SCL: u16 = 0x00A0;
const RECORD_PANE: u16 = 0x0041;
const RECORD_SELECTION: u16 = 0x001D;
const RECORD_DIMENSIONS: u16 = 0x0200;
const RECORD_MERGEDCELLS: u16 = 0x00E5;
const RECORD_BLANK: u16 = 0x0201;
const RECORD_NUMBER: u16 = 0x0203;
const RECORD_HLINK: u16 = 0x01B8;
const RECORD_ROW: u16 = 0x0208;
const RECORD_COLINFO: u16 = 0x007D;

const ROW_OPTION_HIDDEN: u16 = 0x0020;
const ROW_OPTION_COLLAPSED: u16 = 0x1000;
const COLINFO_OPTION_HIDDEN: u16 = 0x0001;
const COLINFO_OPTION_COLLAPSED: u16 = 0x1000;

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
    let workbook_stream = build_workbook_stream(date_1904);

    let cursor = Cursor::new(Vec::new());
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

/// Build a BIFF8 `.xls` fixture containing a single external hyperlink on `A1`.
///
/// This is used to ensure we preserve BIFF `HLINK` records when importing `.xls` workbooks.
pub fn build_hyperlink_fixture_xls() -> Vec<u8> {
    let workbook_stream = build_hyperlink_workbook_stream(
        "Links",
        hlink_external_url(0, 0, 0, 0, "https://example.com", "Example", "Example tooltip"),
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

fn build_workbook_stream(date_1904: bool) -> Vec<u8> {
    // -- Globals -----------------------------------------------------------------
    let mut globals = Vec::<u8>::new();

    push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS)); // BOF: workbook globals
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
    push_record(&mut globals, RECORD_SHEETEXT, &sheetext_record_rgb(0x11, 0x22, 0x33));

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
        stream
            .write_all(&globals)
            .expect("write Workbook stream");
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
        &hlink_external_url(0, 0, 0, 0, "https://example.com", "Example", "Example tooltip"),
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
        0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B,
        0xA9, 0x0B,
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

fn write_hyperlink_string(out: &mut Vec<u8>, s: &str) {
    // HyperlinkString: u32 cch + UTF-16LE (including trailing NUL).
    let mut u16s: Vec<u16> = s.encode_utf16().collect();
    u16s.push(0);
    out.extend_from_slice(&(u16s.len() as u32).to_le_bytes());
    for code_unit in u16s {
        out.extend_from_slice(&code_unit.to_le_bytes());
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

    // Outline rows:
    // - Rows 2-3 (1-based) are detail rows: outline level 1 and hidden (collapsed).
    // - Row 4 (1-based) is the collapsed summary row (level 0, collapsed).
    push_record(&mut sheet, RECORD_ROW, &row_record(1, true, 1, false));
    push_record(&mut sheet, RECORD_ROW, &row_record(2, true, 1, false));
    push_record(&mut sheet, RECORD_ROW, &row_record(3, false, 0, true));

    // Outline columns:
    // - Columns B-C (1-based) are detail columns: outline level 1 and hidden (collapsed).
    // - Column D (1-based) is the collapsed summary column.
    push_record(&mut sheet, RECORD_COLINFO, &colinfo_record(1, 2, true, 1, false));
    push_record(&mut sheet, RECORD_COLINFO, &colinfo_record(3, 3, false, 0, true));

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
        push_record(&mut sheet, RECORD_SELECTION, &selection_single_cell(0, 2, 2));

        push_record(&mut sheet, RECORD_EOF, &[]);
        sheet
    };
    globals[boundsheet2_offset_pos..boundsheet2_offset_pos + 4]
        .copy_from_slice(&(sheet2_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet2);

    globals
}

fn window2() -> [u8; 18] {
    // WINDOW2 record payload (BIFF8). Most fields can be zero for our fixtures.
    let mut out = [0u8; 18];
    let grbit: u16 = 0x02B6;
    out[0..2].copy_from_slice(&grbit.to_le_bytes());
    out
}

fn font(name: &str) -> Vec<u8> {
    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&200u16.to_le_bytes()); // height = 10pt
    out.extend_from_slice(&0u16.to_le_bytes()); // option flags
    out.extend_from_slice(&COLOR_AUTOMATIC.to_le_bytes()); // color: automatic
    out.extend_from_slice(&400u16.to_le_bytes()); // weight: normal
    out.extend_from_slice(&0u16.to_le_bytes()); // escapement
    out.push(0); // underline
    out.push(0); // family
    out.push(0); // charset
    out.push(0); // reserved
    write_short_unicode_string(&mut out, name);
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
    let bytes = s.as_bytes();
    let len: u8 = bytes
        .len()
        .try_into()
        .expect("string too long for u8 length");
    out.push(len);
    out.push(0); // compressed (8-bit)
    out.extend_from_slice(bytes);
}

fn write_unicode_string(out: &mut Vec<u8>, s: &str) {
    // BIFF8 XLUnicodeString: [cch: u16][flags: u8][chars]
    let bytes = s.as_bytes();
    let len: u16 = bytes
        .len()
        .try_into()
        .expect("string too long for u16 length");
    out.extend_from_slice(&len.to_le_bytes());
    out.push(0); // compressed (8-bit)
    out.extend_from_slice(bytes);
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
