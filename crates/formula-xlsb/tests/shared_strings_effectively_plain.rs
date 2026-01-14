use std::fs::File;
use std::io::{Cursor, Read};

use formula_xlsb::{biff12_varint, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use tempfile::tempdir;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

const SST: u32 = 0x009F;
const SI: u32 = 0x0013;
const SST_END: u32 = 0x00A0;
const SHEETDATA: u32 = 0x0091;
const SHEETDATA_END: u32 = 0x0092;
const ROW: u32 = 0x0000;

const CELL_ISST: u32 = 0x0007;

fn write_record(out: &mut Vec<u8>, id: u32, data: &[u8]) {
    biff12_varint::write_record_id(out, id).expect("write record id");
    let len = u32::try_from(data.len()).expect("record too large");
    biff12_varint::write_record_len(out, len).expect("write record len");
    out.extend_from_slice(data);
}

fn write_utf16_string(out: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    let len = u32::try_from(units.len()).expect("string too large");
    out.extend_from_slice(&len.to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
}

fn build_shared_strings_bin_effectively_plain(rich_text: &str, phonetic_text: &str) -> Vec<u8> {
    // Build:
    //   BrtSST(totalCount=0, uniqueCount=2)
    //   BrtSI(flags=RICH, text, cRun=0)
    //   BrtSI(flags=PHONETIC, text, cb=0)
    //   BrtSSTEnd
    let mut out = Vec::new();

    let mut sst = Vec::new();
    sst.extend_from_slice(&0u32.to_le_bytes()); // totalCount
    sst.extend_from_slice(&2u32.to_le_bytes()); // uniqueCount
    write_record(&mut out, SST, &sst);

    // BrtSI rich-but-empty.
    let mut si = Vec::new();
    si.push(0x01); // rich flag
    write_utf16_string(&mut si, rich_text);
    si.extend_from_slice(&0u32.to_le_bytes()); // cRun = 0
    write_record(&mut out, SI, &si);

    // BrtSI phonetic-but-empty.
    let mut si = Vec::new();
    si.push(0x02); // phonetic flag
    write_utf16_string(&mut si, phonetic_text);
    si.extend_from_slice(&0u32.to_le_bytes()); // cb = 0
    write_record(&mut out, SI, &si);

    write_record(&mut out, SST_END, &[]);

    out
}

fn build_shared_strings_bin_rich(text: &str) -> Vec<u8> {
    // Build:
    //   BrtSST(totalCount=0, uniqueCount=1)
    //   BrtSI(flags=RICH, text, cRun=1, StrRun[0])
    //   BrtSSTEnd
    let mut out = Vec::new();

    let mut sst = Vec::new();
    sst.extend_from_slice(&0u32.to_le_bytes()); // totalCount
    sst.extend_from_slice(&1u32.to_le_bytes()); // uniqueCount
    write_record(&mut out, SST, &sst);

    let mut si = Vec::new();
    si.push(0x01); // rich flag
    write_utf16_string(&mut si, text);
    si.extend_from_slice(&1u32.to_le_bytes()); // cRun = 1
                                               // StrRun (8 bytes): [ich:u32][ifnt:u16][reserved:u16]
    si.extend_from_slice(&[0u8; 8]);
    write_record(&mut out, SI, &si);

    write_record(&mut out, SST_END, &[]);

    out
}

fn read_zip_part(path: &str, part_path: &str) -> Vec<u8> {
    let file = File::open(path).expect("open xlsb");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

fn find_cell_record(sheet_bin: &[u8], target_row: u32, target_col: u32) -> Option<(u32, Vec<u8>)> {
    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;

    loop {
        let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
            Some(id) => id,
            None => break,
        };
        let len = match biff12_varint::read_record_len(&mut cursor).ok().flatten() {
            Some(len) => len as usize,
            None => return None,
        };
        let mut payload = vec![0u8; len];
        cursor.read_exact(&mut payload).ok()?;

        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            _ if in_sheet_data => {
                if payload.len() < 8 {
                    continue;
                }
                let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                if current_row == target_row && col == target_col {
                    return Some((id, payload));
                }
            }
            _ => {}
        }
    }
    None
}

#[derive(Debug)]
struct SharedStringsCounts {
    total: u32,
    unique: u32,
    si_records: usize,
}

fn read_shared_strings_counts(shared_strings_bin: &[u8]) -> SharedStringsCounts {
    let mut cursor = Cursor::new(shared_strings_bin);
    let mut total = None;
    let mut unique = None;
    let mut si_records = 0usize;

    loop {
        let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
            Some(id) => id,
            None => break,
        };
        let len = match biff12_varint::read_record_len(&mut cursor).ok().flatten() {
            Some(len) => len as usize,
            None => break,
        };
        let mut payload = vec![0u8; len];
        cursor
            .read_exact(&mut payload)
            .expect("read record payload");

        match id {
            SST if payload.len() >= 8 => {
                total = Some(u32::from_le_bytes(payload[0..4].try_into().unwrap()));
                unique = Some(u32::from_le_bytes(payload[4..8].try_into().unwrap()));
            }
            SI => {
                si_records += 1;
            }
            SST_END => break,
            _ => {}
        }
    }

    SharedStringsCounts {
        total: total.expect("missing BrtSST totalCount"),
        unique: unique.expect("missing BrtSST uniqueCount"),
        si_records,
    }
}

#[test]
fn shared_strings_writer_reuses_effectively_plain_flagged_si_records() {
    let rich_text = "RichFlagButNoRuns";
    let phonetic_text = "PhoneticFlagButNoBytes";

    let shared_strings_bin = build_shared_strings_bin_effectively_plain(rich_text, phonetic_text);

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_shared_strings_bin_override(shared_strings_bin);
    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, &bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open input workbook");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
        0,
        &[
            CellEdit {
                row: 0,
                col: 0,
                new_value: CellValue::Text(rich_text.to_string()),
                new_style: None,
                clear_formula: false,
                new_formula: None,
                new_formula_flags: None,
                new_rgcb: None,
                shared_string_index: None,
                clear_formula: false,
            },
            CellEdit {
                row: 0,
                col: 1,
                new_value: CellValue::Text(phonetic_text.to_string()),
                new_style: None,
                clear_formula: false,
                new_formula: None,
                new_formula_flags: None,
                new_rgcb: None,
                shared_string_index: None,
                clear_formula: false,
            },
        ],
    )
    .expect("save_with_cell_edits_shared_strings");

    // Verify that the worksheet cells reference the *existing* shared string indices (0 and 1),
    // rather than appending new plain strings.
    let sheet_bin = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");

    let (id_a1, payload_a1) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id_a1, CELL_ISST, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload_a1[8..12].try_into().unwrap()),
        0,
        "expected A1 to reuse shared string index 0"
    );

    let (id_b1, payload_b1) = find_cell_record(&sheet_bin, 0, 1).expect("find B1 record");
    assert_eq!(id_b1, CELL_ISST, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload_b1[8..12].try_into().unwrap()),
        1,
        "expected B1 to reuse shared string index 1"
    );

    let shared_strings_out = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    let counts = read_shared_strings_counts(&shared_strings_out);

    // Two cells now reference the SST; no new SI records should be appended.
    assert_eq!(counts.total, 2);
    assert_eq!(counts.unique, 2);
    assert_eq!(counts.si_records, 2);

    // Spot-check that the output workbook can still be opened and read.
    let wb2 = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    assert_eq!(
        sheet
            .cells
            .iter()
            .find(|c| c.row == 0 && c.col == 0)
            .expect("A1 exists")
            .value,
        CellValue::Text(rich_text.to_string())
    );
    assert_eq!(
        sheet
            .cells
            .iter()
            .find(|c| c.row == 0 && c.col == 1)
            .expect("B1 exists")
            .value,
        CellValue::Text(phonetic_text.to_string())
    );
}

#[test]
fn shared_strings_writer_does_not_reuse_true_rich_si_records_as_plain() {
    let text = "RichFlagWithRuns";
    let shared_strings_bin = build_shared_strings_bin_rich(text);

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_shared_strings_bin_override(shared_strings_bin);
    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, &bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open input workbook");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text(text.to_string()),
            new_style: None,
            clear_formula: false,
            new_formula: None,
            new_formula_flags: None,
            new_rgcb: None,
            shared_string_index: None,
            clear_formula: false,
        }],
    )
    .expect("save_with_cell_edits_shared_strings");

    let sheet_bin = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, CELL_ISST, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        1,
        "expected A1 to reference an appended plain shared string, not rich index 0"
    );

    let shared_strings_out = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    let counts = read_shared_strings_counts(&shared_strings_out);
    assert_eq!(counts.total, 1);
    assert_eq!(counts.unique, 2);
    assert_eq!(counts.si_records, 2);
}
