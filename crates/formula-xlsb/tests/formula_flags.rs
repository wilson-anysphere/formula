use std::io::{Read, Seek, Write};
use std::path::Path;

use formula_xlsb::{biff12_varint, patch_sheet_bin, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;

const WORKSHEET: u32 = 0x0081;
const SHEETDATA: u32 = 0x0091;
const SHEETDATA_END: u32 = 0x0092;
const ROW: u32 = 0x0000;
const FORMULA_STRING: u32 = 0x0008;
const FORMULA_FLOAT: u32 = 0x0009;
const FORMULA_BOOL: u32 = 0x000A;
const FORMULA_BOOLERR: u32 = 0x000B;

fn push_record(stream: &mut Vec<u8>, id: u32, data: &[u8]) {
    biff12_varint::write_record_id(stream, id).expect("write record id");
    let len = u32::try_from(data.len()).expect("record too large");
    biff12_varint::write_record_len(stream, len).expect("write record len");
    stream.extend_from_slice(data);
}

fn synthetic_sheet_brt_fmla_num(flags: u16, cached: f64, extra: &[u8]) -> Vec<u8> {
    let mut sheet = Vec::new();

    push_record(&mut sheet, WORKSHEET, &[]);
    push_record(&mut sheet, SHEETDATA, &[]);
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // BrtFmlaNum (FORMULA_FLOAT) record payload:
    //   [col: u32][style: u32][value: f64][flags: u16][cce: u32][rgce bytes...][extra...]
    let mut rec = Vec::new();
    rec.extend_from_slice(&0u32.to_le_bytes()); // col
    rec.extend_from_slice(&0u32.to_le_bytes()); // style
    rec.extend_from_slice(&cached.to_le_bytes());
    rec.extend_from_slice(&flags.to_le_bytes());
    rec.extend_from_slice(&0u32.to_le_bytes()); // cce = 0, rgce empty
    rec.extend_from_slice(extra);

    push_record(&mut sheet, FORMULA_FLOAT, &rec);
    push_record(&mut sheet, SHEETDATA_END, &[]);
    sheet
}

fn push_utf16_chars(out: &mut Vec<u8>, s: &str) {
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
}

fn synthetic_sheet_brt_fmla_bool(flags: u16, cached: bool, extra: &[u8]) -> Vec<u8> {
    let mut sheet = Vec::new();

    push_record(&mut sheet, WORKSHEET, &[]);
    push_record(&mut sheet, SHEETDATA, &[]);
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // BrtFmlaBool (FORMULA_BOOL) record payload:
    //   [col: u32][style: u32][value: u8][flags: u16][cce: u32][rgce bytes...][extra...]
    let mut rec = Vec::new();
    rec.extend_from_slice(&0u32.to_le_bytes()); // col
    rec.extend_from_slice(&0u32.to_le_bytes()); // style
    rec.push(u8::from(cached));
    rec.extend_from_slice(&flags.to_le_bytes());
    rec.extend_from_slice(&0u32.to_le_bytes()); // cce = 0, rgce empty
    rec.extend_from_slice(extra);

    push_record(&mut sheet, FORMULA_BOOL, &rec);
    push_record(&mut sheet, SHEETDATA_END, &[]);
    sheet
}

fn synthetic_sheet_brt_fmla_error(flags: u16, cached: u8, extra: &[u8]) -> Vec<u8> {
    let mut sheet = Vec::new();

    push_record(&mut sheet, WORKSHEET, &[]);
    push_record(&mut sheet, SHEETDATA, &[]);
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // BrtFmlaError (FORMULA_BOOLERR) record payload:
    //   [col: u32][style: u32][value: u8][flags: u16][cce: u32][rgce bytes...][extra...]
    let mut rec = Vec::new();
    rec.extend_from_slice(&0u32.to_le_bytes()); // col
    rec.extend_from_slice(&0u32.to_le_bytes()); // style
    rec.push(cached);
    rec.extend_from_slice(&flags.to_le_bytes());
    rec.extend_from_slice(&0u32.to_le_bytes()); // cce = 0, rgce empty
    rec.extend_from_slice(extra);

    push_record(&mut sheet, FORMULA_BOOLERR, &rec);
    push_record(&mut sheet, SHEETDATA_END, &[]);
    sheet
}

fn synthetic_sheet_brt_fmla_string(flags: u16, cached: &str, extra: &[u8]) -> Vec<u8> {
    let mut sheet = Vec::new();

    push_record(&mut sheet, WORKSHEET, &[]);
    push_record(&mut sheet, SHEETDATA, &[]);
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // BrtFmlaString (FORMULA_STRING) record payload:
    //   [col: u32][style: u32][cch: u32][flags: u16][utf16 chars...][cce: u32][rgce bytes...][extra...]
    let mut rec = Vec::new();
    rec.extend_from_slice(&0u32.to_le_bytes()); // col
    rec.extend_from_slice(&0u32.to_le_bytes()); // style
    let cch = cached.encode_utf16().count() as u32;
    rec.extend_from_slice(&cch.to_le_bytes());
    rec.extend_from_slice(&flags.to_le_bytes());
    push_utf16_chars(&mut rec, cached);
    rec.extend_from_slice(&0u32.to_le_bytes()); // cce = 0, rgce empty
    rec.extend_from_slice(extra);

    push_record(&mut sheet, FORMULA_STRING, &rec);
    push_record(&mut sheet, SHEETDATA_END, &[]);
    sheet
}

fn write_fixture_like_xlsb(sheet1_bin: &[u8]) -> tempfile::NamedTempFile {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/simple.xlsb"
    ));
    let file = std::fs::File::open(fixture_path).expect("open xlsb fixture");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");

    let mut out_file = tempfile::NamedTempFile::new().expect("create temp file");
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);
    {
        let mut writer = zip::ZipWriter::new(out_file.as_file_mut());
        for i in 0..zip.len() {
            let mut entry = zip.by_index(i).expect("fixture entry");
            if entry.is_dir() {
                continue;
            }
            let name = entry.name().to_string();
            let mut bytes = Vec::with_capacity(entry.size() as usize);
            entry.read_to_end(&mut bytes).expect("read fixture entry");
            if name == "xl/worksheets/sheet1.bin" {
                bytes.clear();
                bytes.extend_from_slice(sheet1_bin);
            }
            writer
                .start_file(name, options.clone())
                .expect("start zip entry");
            writer.write_all(&bytes).expect("write zip entry");
        }
        writer.finish().expect("finish zip");
    }

    // Ensure the file cursor is reset for any future readers.
    out_file
        .as_file_mut()
        .seek(std::io::SeekFrom::Start(0))
        .unwrap();
    out_file
}

#[test]
fn parses_and_preserves_brt_fmla_flags() {
    let flags = 0x1234;
    let extra = [0xDE, 0xAD, 0xBE, 0xEF];
    let sheet_bin = synthetic_sheet_brt_fmla_num(flags, 1.0, &extra);

    let tmp = write_fixture_like_xlsb(&sheet_bin);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.row, 0);
    assert_eq!(cell.col, 0);
    assert_eq!(cell.value, CellValue::Number(1.0));

    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_updates_cached_value_without_changing_flags() {
    let flags = 0x2222;
    let extra = [0xAA, 0xBB];
    let sheet_bin = synthetic_sheet_brt_fmla_num(flags, 10.0, &extra);

    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(99.5),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
    };
    let patched_sheet = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    let tmp = write_fixture_like_xlsb(&patched_sheet);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.value, CellValue::Number(99.5));
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_can_override_existing_brt_fmla_flags() {
    let flags = 0x2222;
    let new_flags = 0x3333;
    let extra = [0xAA, 0xBB];
    let sheet_bin = synthetic_sheet_brt_fmla_num(flags, 10.0, &extra);

    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(99.5),
        new_style: None,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: Some(new_flags),
        shared_string_index: None,
    };
    let patched_sheet = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    let tmp = write_fixture_like_xlsb(&patched_sheet);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.value, CellValue::Number(99.5));
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, new_flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_can_insert_formula_cells_with_explicit_flags() {
    let mut sheet_bin = Vec::new();
    push_record(&mut sheet_bin, WORKSHEET, &[]);
    push_record(&mut sheet_bin, SHEETDATA, &[]);
    push_record(&mut sheet_bin, SHEETDATA_END, &[]);

    let flags = 0x1234;
    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(1.0),
        new_style: None,
        new_formula: Some(Vec::new()), // empty rgce
        new_rgcb: None,
        new_formula_flags: Some(flags),
        shared_string_index: None,
    };
    let patched_sheet = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    let tmp = write_fixture_like_xlsb(&patched_sheet);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.row, 0);
    assert_eq!(cell.col, 0);
    assert_eq!(cell.value, CellValue::Number(1.0));
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
}

#[test]
fn parses_and_preserves_brt_fmla_bool_flags() {
    let flags = 0x1234;
    let extra = [0xDE, 0xAD, 0xBE, 0xEF];
    let sheet_bin = synthetic_sheet_brt_fmla_bool(flags, true, &extra);

    let tmp = write_fixture_like_xlsb(&sheet_bin);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.row, 0);
    assert_eq!(cell.col, 0);
    assert_eq!(cell.value, CellValue::Bool(true));

    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_updates_cached_bool_without_changing_flags() {
    let flags = 0x2222;
    let extra = [0xAA, 0xBB];
    let sheet_bin = synthetic_sheet_brt_fmla_bool(flags, true, &extra);

    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Bool(false),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
    };
    let patched_sheet = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    let tmp = write_fixture_like_xlsb(&patched_sheet);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.value, CellValue::Bool(false));
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn parses_and_preserves_brt_fmla_error_flags() {
    let flags = 0x1234;
    let extra = [0xDE, 0xAD];
    let sheet_bin = synthetic_sheet_brt_fmla_error(flags, 0x07, &extra);

    let tmp = write_fixture_like_xlsb(&sheet_bin);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.value, CellValue::Error(0x07));

    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_updates_cached_error_without_changing_flags() {
    let flags = 0x2222;
    let extra = [0xAA];
    let sheet_bin = synthetic_sheet_brt_fmla_error(flags, 0x07, &extra);

    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Error(0x2A),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
    };
    let patched_sheet = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    let tmp = write_fixture_like_xlsb(&patched_sheet);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.value, CellValue::Error(0x2A));
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn parses_and_preserves_brt_fmla_string_flags() {
    let flags = 0x1234;
    let extra = [0x10, 0x20, 0x30];
    let sheet_bin = synthetic_sheet_brt_fmla_string(flags, "Hello", &extra);

    let tmp = write_fixture_like_xlsb(&sheet_bin);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.value, CellValue::Text("Hello".to_string()));

    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_updates_cached_string_without_changing_flags() {
    // BrtFmlaString reuses the BIFF12 wide-string flag bits (0x0001 rich runs, 0x0002 phonetic)
    // to signal the presence of cached formatting payloads. Use a value that avoids those bits so
    // the synthetic record layout stays valid.
    let flags = 0x2220;
    let extra = [0xAA, 0xBB, 0xCC];
    let sheet_bin = synthetic_sheet_brt_fmla_string(flags, "Hello", &extra);

    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text("World".to_string()),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
    };
    let patched_sheet = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    let tmp = write_fixture_like_xlsb(&patched_sheet);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.value, CellValue::Text("World".to_string()));
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_updates_cached_string_with_reserved_flags_and_4byte_extra() {
    let flags = 0x2222;
    let extra = [0xAA, 0xBB, 0xCC, 0xDD];
    let sheet_bin = synthetic_sheet_brt_fmla_string(flags, "Hello", &extra);

    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text("World".to_string()),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
    };
    let patched_sheet = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    let tmp = write_fixture_like_xlsb(&patched_sheet);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    assert_eq!(sheet.cells.len(), 1);
    let cell = &sheet.cells[0];
    assert_eq!(cell.value, CellValue::Text("World".to_string()));
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.flags, flags);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_is_byte_identical_for_noop_brt_fmla_bool_with_extra_bytes() {
    let flags = 0x2222;
    let extra = [0xAA, 0xBB];
    let sheet_bin = synthetic_sheet_brt_fmla_bool(flags, true, &extra);

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Bool(true),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet");

    assert_eq!(patched, sheet_bin);
}

#[test]
fn patcher_is_byte_identical_for_noop_brt_fmla_error_with_extra_bytes() {
    let flags = 0x2222;
    let extra = [0xAA, 0xBB];
    let sheet_bin = synthetic_sheet_brt_fmla_error(flags, 0x07, &extra);

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Error(0x07),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet");

    assert_eq!(patched, sheet_bin);
}

#[test]
fn patcher_is_byte_identical_for_noop_brt_fmla_string_with_extra_bytes() {
    // Use flags with reserved bits + phonetic bit set, but no actual phonetic payload. This is a
    // common real-world shape and is handled by the cached-string offset parser.
    let flags = 0x2222;
    let extra = [0xAA, 0xBB, 0xCC, 0xDD];
    let sheet_bin = synthetic_sheet_brt_fmla_string(flags, "Hello", &extra);

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello".to_string()),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet");

    assert_eq!(patched, sheet_bin);
}

#[test]
fn patcher_requires_new_rgcb_when_replacing_rgce_for_brt_fmla_string_with_existing_extra() {
    fn ptg_str(s: &str) -> Vec<u8> {
        let mut out = vec![0x17]; // PtgStr
        let units: Vec<u16> = s.encode_utf16().collect();
        out.extend_from_slice(&(units.len() as u16).to_le_bytes());
        for u in units {
            out.extend_from_slice(&u.to_le_bytes());
        }
        out
    }

    let flags = 0x2220;
    let extra = [0xAA, 0xBB, 0xCC];
    let sheet_bin = synthetic_sheet_brt_fmla_string(flags, "Hello", &extra);
    let new_rgce = ptg_str("Hello");

    let err = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello".to_string()),
            new_formula: Some(new_rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect_err("expected InvalidInput when changing rgce without supplying new_rgcb");
    match err {
        formula_xlsb::Error::Io(io_err) => {
            assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidInput)
        }
        other => panic!("expected InvalidInput, got {other:?}"),
    }

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello".to_string()),
            new_formula: Some(new_rgce.clone()),
            new_rgcb: Some(extra.to_vec()),
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet bin with explicit rgcb");

    let tmp = write_fixture_like_xlsb(&patched);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");
    let cell = &sheet.cells[0];
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.rgce, new_rgce);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_requires_new_rgcb_when_replacing_rgce_for_brt_fmla_bool_with_existing_extra() {
    let flags = 0x2222;
    let extra = [0xAA, 0xBB];
    let sheet_bin = synthetic_sheet_brt_fmla_bool(flags, false, &extra);
    let new_rgce = vec![0x1D, 0x01]; // PtgBool TRUE

    let err = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Bool(true),
            new_formula: Some(new_rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect_err("expected InvalidInput when changing rgce without supplying new_rgcb");
    match err {
        formula_xlsb::Error::Io(io_err) => {
            assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidInput)
        }
        other => panic!("expected InvalidInput, got {other:?}"),
    }

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Bool(true),
            new_formula: Some(new_rgce.clone()),
            new_rgcb: Some(extra.to_vec()),
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet bin with explicit rgcb");

    let tmp = write_fixture_like_xlsb(&patched);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");
    let cell = &sheet.cells[0];
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.rgce, new_rgce);
    assert_eq!(formula.extra, extra.to_vec());
}

#[test]
fn patcher_requires_new_rgcb_when_replacing_rgce_for_brt_fmla_error_with_existing_extra() {
    let flags = 0x2222;
    let extra = [0xAA, 0xBB];
    let sheet_bin = synthetic_sheet_brt_fmla_error(flags, 0x07, &extra);
    let new_rgce = vec![0x1C, 0x2A]; // PtgErr #N/A

    let err = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Error(0x2A),
            new_formula: Some(new_rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect_err("expected InvalidInput when changing rgce without supplying new_rgcb");
    match err {
        formula_xlsb::Error::Io(io_err) => {
            assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidInput)
        }
        other => panic!("expected InvalidInput, got {other:?}"),
    }

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Error(0x2A),
            new_formula: Some(new_rgce.clone()),
            new_rgcb: Some(extra.to_vec()),
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet bin with explicit rgcb");

    let tmp = write_fixture_like_xlsb(&patched);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open patched xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");
    let cell = &sheet.cells[0];
    let formula = cell.formula.as_ref().expect("formula expected");
    assert_eq!(formula.rgce, new_rgce);
    assert_eq!(formula.extra, extra.to_vec());
}
