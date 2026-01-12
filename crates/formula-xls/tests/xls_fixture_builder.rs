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
const RECORD_FORMAT: u16 = 0x041E;
const RECORD_CONTINUE: u16 = 0x003C;
const RECORD_XF: u16 = 0x00E0;
const RECORD_BOUNDSHEET: u16 = 0x0085;
const RECORD_WINDOW2: u16 = 0x023E;
const RECORD_DIMENSIONS: u16 = 0x0200;
const RECORD_MERGEDCELLS: u16 = 0x00E5;
const RECORD_BLANK: u16 = 0x0201;
const RECORD_NUMBER: u16 = 0x0203;

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

fn window2() -> [u8; 18] {
    // WINDOW2 record payload (BIFF8). Most fields can be zero for our fixtures.
    [0u8; 18]
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
