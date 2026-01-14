use std::path::Path;

use formula_io::{open_workbook, save_workbook};
use formula_model::{CalculationMode, CellRef, DefinedNameScope};
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

#[test]
fn xlsb_export_preserves_calc_mode_manual() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    let wb = formula_xlsb::XlsbWorkbook::open(fixture_path).expect("open xlsb fixture");

    let workbook_bin = wb
        .preserved_parts()
        .get("xl/workbook.bin")
        .expect("expected xl/workbook.bin to be preserved");

    // CALC_PROP flags: low 2 bits are calc mode. 0 = Manual.
    let patched_workbook_bin = patch_workbook_bin_append_calc_prop(workbook_bin, 0u16);

    let dir = tempfile::tempdir().expect("temp dir");
    let manual_xlsb_path = dir.path().join("manual_calc.xlsb");
    wb.save_with_part_overrides(
        &manual_xlsb_path,
        &HashMap::from([("xl/workbook.bin".to_string(), patched_workbook_bin)]),
    )
    .expect("write modified xlsb");

    let wb = open_workbook(&manual_xlsb_path).expect("open modified xlsb via formula-io");

    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("load exported xlsx package");
    let settings = pkg.calc_settings().expect("read exported calc settings");
    assert_eq!(settings.calculation_mode, CalculationMode::Manual);
}

#[test]
fn xlsb_export_preserves_full_calc_on_load() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    let wb = formula_xlsb::XlsbWorkbook::open(fixture_path).expect("open xlsb fixture");

    let workbook_bin = wb
        .preserved_parts()
        .get("xl/workbook.bin")
        .expect("expected xl/workbook.bin to be preserved");

    // CALC_PROP flags:
    // - low 2 bits are calc mode (0 = Manual)
    // - bit 0x0004 is fullCalcOnLoad.
    let patched_workbook_bin = patch_workbook_bin_append_calc_prop(workbook_bin, 0x0004u16);

    let dir = tempfile::tempdir().expect("temp dir");
    let xlsb_path = dir.path().join("full_calc_on_load.xlsb");
    wb.save_with_part_overrides(
        &xlsb_path,
        &HashMap::from([("xl/workbook.bin".to_string(), patched_workbook_bin)]),
    )
    .expect("write modified xlsb");

    let wb = open_workbook(&xlsb_path).expect("open modified xlsb via formula-io");

    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("load exported xlsx package");
    let settings = pkg.calc_settings().expect("read exported calc settings");
    assert!(
        settings.full_calc_on_load,
        "expected fullCalcOnLoad to be preserved"
    );
}

#[test]
fn xlsb_export_preserves_defined_names_and_ptgname_formulas() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    let wb = formula_xlsb::XlsbWorkbook::open_with_options(
        fixture_path,
        formula_xlsb::OpenOptions {
            preserve_worksheets: true,
            ..Default::default()
        },
    )
    .expect("open xlsb fixture");

    let workbook_bin = wb
        .preserved_parts()
        .get("xl/workbook.bin")
        .expect("expected xl/workbook.bin to be preserved");
    let sheet_bin = wb
        .preserved_parts()
        .get("xl/worksheets/sheet1.bin")
        .expect("expected xl/worksheets/sheet1.bin to be preserved");

    let patched_workbook_bin = patch_workbook_bin_append_defined_name(workbook_bin, "MyName", 42);

    // Insert a new numeric formula cell at D1 that references the defined name via `PtgName`.
    // rgce: PtgName (ref class) + nameId=1 + reserved.
    let mut rgce = vec![0x23];
    rgce.extend_from_slice(&1u32.to_le_bytes());
    rgce.extend_from_slice(&0u16.to_le_bytes());
    let patched_sheet_bin = formula_xlsb::patch_sheet_bin(
        sheet_bin,
        &[formula_xlsb::CellEdit {
            row: 0,
            col: 3,
            new_value: formula_xlsb::CellValue::Number(0.0),
            new_formula: Some(rgce),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet");

    let dir = tempfile::tempdir().expect("temp dir");
    let xlsb_path = dir.path().join("defined_names.xlsb");
    wb.save_with_part_overrides(
        &xlsb_path,
        &HashMap::from([
            ("xl/workbook.bin".to_string(), patched_workbook_bin),
            ("xl/worksheets/sheet1.bin".to_string(), patched_sheet_bin),
        ]),
    )
    .expect("write modified xlsb");

    let wb = open_workbook(&xlsb_path).expect("open modified xlsb via formula-io");

    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");

    let name = doc
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "MyName")
        .expect("expected MyName to be preserved as a defined name");
    assert_eq!(name.scope, DefinedNameScope::Workbook);
    assert_eq!(name.refers_to, "42");

    let sheet = doc
        .workbook
        .sheet_by_name("Sheet1")
        .unwrap_or_else(|| &doc.workbook.sheets[0]);
    let formula_cell = CellRef::from_a1("D1").expect("valid cell ref");
    assert_eq!(sheet.formula(formula_cell), Some("MyName"));
}

fn patch_workbook_bin_set_first_sheet_visibility(
    workbook_bin: &[u8],
    visibility_state: u32,
) -> Vec<u8> {
    const SHEET_RECORD_ID: u32 = 0x009C;
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

fn patch_workbook_bin_append_calc_prop(workbook_bin: &[u8], flags: u16) -> Vec<u8> {
    const CALC_PROP_RECORD_ID: u32 = 0x009A;

    let mut out = workbook_bin.to_vec();
    let payload = {
        let mut buf = Vec::new();
        buf.extend_from_slice(&0u32.to_le_bytes()); // calcId (unused by parser)
        buf.extend_from_slice(&flags.to_le_bytes());
        buf
    };

    formula_xlsb::biff12_varint::write_record_id(&mut out, CALC_PROP_RECORD_ID)
        .expect("write calcProp record id");
    formula_xlsb::biff12_varint::write_record_len(&mut out, payload.len() as u32)
        .expect("write calcProp record len");
    out.extend_from_slice(&payload);
    out
}

fn patch_workbook_bin_append_defined_name(workbook_bin: &[u8], name: &str, value: u16) -> Vec<u8> {
    // BrtName record id in BIFF12 / MS-XLSB.
    const NAME_RECORD_ID: u32 = 0x0027;
    const WORKBOOK_SCOPE: u32 = 0xFFFF_FFFF;

    let rgce = {
        // PtgInt (0x1E) literal.
        let mut buf = vec![0x1E];
        buf.extend_from_slice(&value.to_le_bytes());
        buf
    };

    let payload = {
        let mut buf = Vec::new();
        buf.extend_from_slice(&0u32.to_le_bytes()); // flags (hidden/etc)
        buf.extend_from_slice(&WORKBOOK_SCOPE.to_le_bytes()); // scope sheet index
        buf.push(0); // reserved
        write_utf16_string(&mut buf, name);
        buf.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
        buf.extend_from_slice(&rgce);
        buf
    };

    let mut out = workbook_bin.to_vec();
    formula_xlsb::biff12_varint::write_record_id(&mut out, NAME_RECORD_ID)
        .expect("write name record id");
    formula_xlsb::biff12_varint::write_record_len(&mut out, payload.len() as u32)
        .expect("write name record len");
    out.extend_from_slice(&payload);
    out
}

fn write_utf16_string(out: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    out.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
}
