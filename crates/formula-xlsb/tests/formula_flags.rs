use std::io::{Read, Seek, Write};
use std::path::Path;

use formula_xlsb::{patch_sheet_bin, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;

const WORKSHEET: u32 = 0x0181;
const SHEETDATA: u32 = 0x0191;
const SHEETDATA_END: u32 = 0x0192;
const ROW: u32 = 0x0000;
const FORMULA_FLOAT: u32 = 0x0009;

fn write_record_id(out: &mut Vec<u8>, mut id: u32) {
    // XLSB record ids use a variable-length encoding where the high bit of each
    // byte indicates continuation (and is included in the value). This matches
    // `Biff12Reader::read_id` in `parser.rs`.
    let mut bytes = Vec::new();
    loop {
        bytes.push((id & 0xFF) as u8);
        id >>= 8;
        if id == 0 {
            break;
        }
    }

    // Ensure the last byte terminates (msb cleared).
    if bytes.last().copied().unwrap_or(0) & 0x80 != 0 {
        bytes.push(0);
    }

    assert!(bytes.len() <= 4, "record id too large for XLSB encoding");
    for (idx, b) in bytes.iter().enumerate() {
        if idx + 1 == bytes.len() {
            assert_eq!(b & 0x80, 0, "last record id byte must have msb cleared");
        } else {
            assert_ne!(b & 0x80, 0, "intermediate record id byte must have msb set");
        }
        out.push(*b);
    }
}

fn write_record_len(out: &mut Vec<u8>, mut len: u32) {
    // Standard 7-bit LEB128 encoding (matches `Biff12Reader::read_len`).
    loop {
        let mut byte = (len & 0x7F) as u8;
        len >>= 7;
        if len != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if len == 0 {
            break;
        }
    }
}

fn push_record(stream: &mut Vec<u8>, id: u32, data: &[u8]) {
    write_record_id(stream, id);
    write_record_len(stream, data.len() as u32);
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

fn write_fixture_like_xlsb(sheet1_bin: &[u8]) -> tempfile::NamedTempFile {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/simple.xlsb"
    ));
    let file = std::fs::File::open(fixture_path).expect("open xlsb fixture");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");

    let mut out_file = tempfile::NamedTempFile::new().expect("create temp file");
    let options =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
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
            writer.start_file(name, options).expect("start zip entry");
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
