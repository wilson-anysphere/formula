use std::path::Path;

use formula_io::{open_workbook, save_workbook};
use formula_model::CellRef;
use std::collections::HashMap;
use std::io::Cursor;

#[test]
fn xlsb_export_preserves_date_system_1904() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/date1904.xlsb"
    ));
    let wb = open_workbook(fixture_path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");
    assert_eq!(
        doc.workbook.date_system,
        formula_model::DateSystem::Excel1904
    );
}

#[test]
fn xlsb_export_preserves_date_number_format_style() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures_styles/date.xlsb"
    ));
    let wb = open_workbook(fixture_path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");

    let sheet = doc
        .workbook
        .sheet_by_name("Sheet1")
        .unwrap_or_else(|| &doc.workbook.sheets[0]);
    let cell_ref = CellRef::from_a1("A1").expect("valid cell ref");
    let cell = sheet.cell(cell_ref).expect("expected A1 to exist");
    assert_ne!(cell.style_id, 0, "expected A1 to have a non-default style");

    let style = doc
        .workbook
        .styles
        .get(cell.style_id)
        .expect("expected style to exist in workbook style table");
    assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));
}

#[test]
fn xlsb_export_preserves_sheet_visibility() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    let wb = formula_xlsb::XlsbWorkbook::open(fixture_path).expect("open xlsb fixture");

    let workbook_bin = wb
        .preserved_parts()
        .get("xl/workbook.bin")
        .expect("expected xl/workbook.bin to be preserved");
    let patched_workbook_bin =
        patch_workbook_bin_set_first_sheet_visibility(workbook_bin, 1 /* hidden */);

    let dir = tempfile::tempdir().expect("temp dir");
    let hidden_xlsb_path = dir.path().join("hidden.xlsb");
    wb.save_with_part_overrides(
        &hidden_xlsb_path,
        &HashMap::from([("xl/workbook.bin".to_string(), patched_workbook_bin)]),
    )
    .expect("write modified xlsb");

    let wb = open_workbook(&hidden_xlsb_path).expect("open modified xlsb via formula-io");

    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");

    assert_eq!(doc.workbook.sheets.len(), 1);
    assert_eq!(
        doc.workbook.sheets[0].visibility,
        formula_model::SheetVisibility::Hidden
    );
}

#[test]
fn xlsb_export_preserves_sheet_very_hidden_visibility() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    let wb = formula_xlsb::XlsbWorkbook::open(fixture_path).expect("open xlsb fixture");

    let workbook_bin = wb
        .preserved_parts()
        .get("xl/workbook.bin")
        .expect("expected xl/workbook.bin to be preserved");
    let patched_workbook_bin =
        patch_workbook_bin_set_first_sheet_visibility(workbook_bin, 2 /* veryHidden */);

    let dir = tempfile::tempdir().expect("temp dir");
    let very_hidden_xlsb_path = dir.path().join("very_hidden.xlsb");
    wb.save_with_part_overrides(
        &very_hidden_xlsb_path,
        &HashMap::from([("xl/workbook.bin".to_string(), patched_workbook_bin)]),
    )
    .expect("write modified xlsb");

    let wb = open_workbook(&very_hidden_xlsb_path).expect("open modified xlsb via formula-io");

    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");

    assert_eq!(doc.workbook.sheets.len(), 1);
    assert_eq!(
        doc.workbook.sheets[0].visibility,
        formula_model::SheetVisibility::VeryHidden
    );
}

fn patch_workbook_bin_set_first_sheet_visibility(
    workbook_bin: &[u8],
    visibility_state: u32,
) -> Vec<u8> {
    const SHEET_RECORD_ID: u32 = 0x019C;
    const SHEET_STATE_MASK: u32 = 0x0003;

    let mut out = workbook_bin.to_vec();
    let mut cursor = Cursor::new(workbook_bin);
    loop {
        let Some(id) = formula_xlsb::biff12_varint::read_record_id(&mut cursor)
            .expect("read BIFF12 record id")
        else {
            break;
        };
        let Some(len) = formula_xlsb::biff12_varint::read_record_len(&mut cursor)
            .expect("read BIFF12 record len")
        else {
            break;
        };

        let payload_start = cursor.position() as usize;
        let payload_end = payload_start
            .checked_add(len as usize)
            .expect("record length should not overflow");

        if id == SHEET_RECORD_ID {
            let current =
                u32::from_le_bytes(out[payload_start..payload_start + 4].try_into().unwrap());
            let next = (current & !SHEET_STATE_MASK) | (visibility_state & SHEET_STATE_MASK);
            out[payload_start..payload_start + 4].copy_from_slice(&next.to_le_bytes());
            return out;
        }

        cursor.set_position(payload_end as u64);
    }

    panic!("expected workbook.bin to contain a BrtSheet record");
}
