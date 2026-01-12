use std::io::{Cursor, Read};

use formula_model::Range;

mod common;

use common::xls_fixture_builder;

const RECORD_BOUNDSHEET: u16 = 0x0085;

fn read_workbook_stream_from_xls_bytes(data: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(data.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).expect("open xls cfb");

    for candidate in ["/Workbook", "/Book", "Workbook", "Book"] {
        if let Ok(mut stream) = ole.open_stream(candidate) {
            let mut buf = Vec::new();
            stream.read_to_end(&mut buf).expect("read workbook stream");
            return buf;
        }
    }

    panic!("fixture missing Workbook/Book stream");
}

fn find_first_boundsheet_offset(workbook_stream: &[u8]) -> Option<u32> {
    let mut offset = 0usize;
    while offset + 4 <= workbook_stream.len() {
        let header = &workbook_stream[offset..offset + 4];
        let record_id = u16::from_le_bytes([header[0], header[1]]);
        let len = u16::from_le_bytes([header[2], header[3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start.checked_add(len)?;
        let Some(data) = workbook_stream.get(data_start..data_end) else {
            break;
        };

        if record_id == RECORD_BOUNDSHEET && data.len() >= 4 {
            return Some(u32::from_le_bytes([data[0], data[1], data[2], data[3]]));
        }

        offset = data_end;
    }
    None
}

#[test]
fn parses_biff_mergedcells_records_from_fixture_workbook_stream() {
    let bytes = xls_fixture_builder::build_merged_formatted_blank_fixture_xls();
    let workbook_stream = read_workbook_stream_from_xls_bytes(&bytes);

    let sheet_offset = find_first_boundsheet_offset(&workbook_stream).expect("sheet offset");

    // Call the BIFF parser directly to simulate the calamine merge-cells path being absent.
    let ranges =
        formula_xls::parse_biff_sheet_merged_cells(&workbook_stream, sheet_offset as usize)
            .expect("parse merged cells");

    let expected = Range::from_a1("A1:B1").unwrap();
    assert!(
        ranges.iter().any(|r| *r == expected),
        "expected merged range {expected}, got {ranges:?}"
    );
}

