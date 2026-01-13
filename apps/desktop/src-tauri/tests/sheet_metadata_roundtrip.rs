use desktop::file_io::Workbook;
use desktop::persistence::{write_xlsx_from_storage, WorkbookPersistenceLocation};
use desktop::state::AppState;
use formula_model::{SheetVisibility, TabColor};
use tempfile::tempdir;

#[test]
fn read_xlsx_preserves_sheet_visibility_and_tab_color() {
    let mut model = formula_model::Workbook::new();
    model.add_sheet("Sheet1".to_string()).expect("add sheet1");
    let sheet2_id = model.add_sheet("Sheet2".to_string()).expect("add sheet2");
    let sheet3_id = model.add_sheet("Sheet3".to_string()).expect("add sheet3");

    {
        let sheet2 = model.sheet_mut(sheet2_id).expect("sheet2 exists");
        sheet2.visibility = SheetVisibility::Hidden;
        sheet2.tab_color = Some(TabColor::rgb("FF112233"));
    }
    {
        let sheet3 = model.sheet_mut(sheet3_id).expect("sheet3 exists");
        sheet3.visibility = SheetVisibility::VeryHidden;
        sheet3.tab_color = Some(TabColor {
            theme: Some(1),
            tint: Some(0.5),
            ..Default::default()
        });
    }

    let mut cursor = std::io::Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write xlsx bytes");
    let bytes = cursor.into_inner();

    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("input.xlsx");
    std::fs::write(&path, bytes).expect("write temp xlsx");

    let workbook = desktop::file_io::read_xlsx_blocking(&path).expect("read xlsx");
    let sheet2 = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet2")
        .expect("sheet2 in app workbook");
    assert_eq!(sheet2.visibility, SheetVisibility::Hidden);
    assert_eq!(sheet2.tab_color, Some(TabColor::rgb("FF112233")));

    let sheet3 = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet3")
        .expect("sheet3 in app workbook");
    assert_eq!(sheet3.visibility, SheetVisibility::VeryHidden);
    assert_eq!(
        sheet3.tab_color,
        Some(TabColor {
            theme: Some(1),
            tint: Some(0.5),
            ..Default::default()
        })
    );
}

#[test]
fn sheet_metadata_round_trips_through_persistence_and_xlsx_save() {
    let mut state = AppState::new();
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.add_sheet("Sheet2".to_string());
    workbook.add_sheet("Sheet3".to_string());
    // Simulate imported metadata that the UI doesn't mutate directly (eg: `veryHidden`).
    // This should survive persistence + XLSX export even if we only edit other sheets.
    if let Some(sheet3) = workbook.sheets.iter_mut().find(|s| s.id == "Sheet3") {
        sheet3.visibility = SheetVisibility::VeryHidden;
        sheet3.tab_color = Some(TabColor {
            theme: Some(1),
            tint: Some(0.5),
            ..Default::default()
        });
    }

    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
        .expect("load workbook");

    state
        .set_sheet_visibility("Sheet2", SheetVisibility::Hidden)
        .expect("set sheet2 hidden");
    state
        .set_sheet_tab_color("Sheet2", Some(TabColor::rgb("FF00FF00")))
        .expect("set sheet2 tab color");

    let info = state.workbook_info().expect("workbook_info");
    let sheet2 = info
        .sheets
        .iter()
        .find(|s| s.id == "Sheet2")
        .expect("sheet2 in workbook_info");
    assert_eq!(sheet2.visibility, SheetVisibility::Hidden);
    assert_eq!(sheet2.tab_color, Some(TabColor::rgb("FF00FF00")));

    let sheet3 = info
        .sheets
        .iter()
        .find(|s| s.id == "Sheet3")
        .expect("sheet3 in workbook_info");
    assert_eq!(sheet3.visibility, SheetVisibility::VeryHidden);
    assert_eq!(
        sheet3.tab_color,
        Some(TabColor {
            theme: Some(1),
            tint: Some(0.5),
            ..Default::default()
        })
    );

    let storage = state
        .persistent_storage()
        .expect("persistent storage should be available");
    let workbook_id = state
        .persistent_workbook_id()
        .expect("persistent workbook id should be available");
    let model = storage
        .export_model_workbook(workbook_id)
        .expect("export model workbook");
    let model_sheet2 = model
        .sheets
        .iter()
        .find(|s| s.name == "Sheet2")
        .expect("sheet2 in model workbook");
    assert_eq!(model_sheet2.visibility, SheetVisibility::Hidden);
    assert_eq!(model_sheet2.tab_color, Some(TabColor::rgb("FF00FF00")));
    let model_sheet3 = model
        .sheets
        .iter()
        .find(|s| s.name == "Sheet3")
        .expect("sheet3 in model workbook");
    assert_eq!(model_sheet3.visibility, SheetVisibility::VeryHidden);
    assert_eq!(
        model_sheet3.tab_color,
        Some(TabColor {
            theme: Some(1),
            tint: Some(0.5),
            ..Default::default()
        })
    );

    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("out.xlsx");
    let bytes = write_xlsx_from_storage(&storage, workbook_id, state.get_workbook().unwrap(), &path)
        .expect("write xlsx from storage");

    let doc = formula_xlsx::load_from_bytes(bytes.as_ref()).expect("load from bytes");
    let workbook_xml = doc
        .parts()
        .get("xl/workbook.xml")
        .expect("workbook.xml part");
    let workbook_xml = std::str::from_utf8(workbook_xml).expect("workbook.xml utf8");
    assert!(
        workbook_xml.contains("state=\"hidden\""),
        "expected workbook.xml to contain a hidden sheet state"
    );
    assert!(
        workbook_xml.contains("state=\"veryHidden\""),
        "expected workbook.xml to contain a veryHidden sheet state"
    );

    let sheet2_xml = doc
        .parts()
        .get("xl/worksheets/sheet2.xml")
        .expect("sheet2.xml part");
    let sheet2_xml = std::str::from_utf8(sheet2_xml).expect("sheet2.xml utf8");
    assert!(
        sheet2_xml.contains("tabColor"),
        "expected sheet2.xml to contain a tabColor element"
    );
    assert!(
        sheet2_xml.contains("rgb=\"FF00FF00\""),
        "expected sheet2.xml tabColor to contain the rgb ARGB payload"
    );

    let model_roundtrip =
        formula_xlsx::read_workbook_model_from_bytes(bytes.as_ref()).expect("read workbook model");
    let roundtrip_sheet3 = model_roundtrip
        .sheets
        .iter()
        .find(|s| s.name == "Sheet3")
        .expect("Sheet3 in round-tripped workbook");
    assert_eq!(roundtrip_sheet3.visibility, SheetVisibility::VeryHidden);
    assert_eq!(
        roundtrip_sheet3.tab_color,
        Some(TabColor {
            theme: Some(1),
            tint: Some(0.5),
            ..Default::default()
        })
    );
}

#[test]
fn sheet_metadata_edits_preserve_unknown_origin_parts_on_patch_save() {
    // Build a small XLSX workbook and inject a synthetic "unknown" part to prove the save path
    // preserves origin ZIP contents (i.e. it uses the patch-based save path rather than exporting
    // a fresh XLSX from storage/model).
    let mut model = formula_model::Workbook::new();
    model.add_sheet("Sheet1".to_string()).expect("add sheet1");
    model.add_sheet("Sheet2".to_string()).expect("add sheet2");

    let mut cursor = std::io::Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write base xlsx bytes");
    let base_bytes = cursor.into_inner();

    let mut pkg = formula_xlsx::XlsxPackage::from_bytes(&base_bytes).expect("parse base package");
    pkg.set_part("customXml/item1.xml", b"<custom>Hello</custom>".to_vec());
    let bytes = pkg.write_to_bytes().expect("write package with custom part");

    let dir = tempdir().expect("tempdir");
    let input_path = dir.path().join("input.xlsx");
    std::fs::write(&input_path, &bytes).expect("write input xlsx");

    let mut state = AppState::new();
    let workbook = desktop::file_io::read_xlsx_blocking(&input_path).expect("read input xlsx");
    assert!(
        workbook.origin_xlsx_bytes.is_some(),
        "expected read_xlsx_blocking to retain origin bytes"
    );
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
        .expect("load workbook");

    // Apply sheet metadata edits via the same AppState APIs the Tauri commands call.
    state
        .set_sheet_visibility("Sheet2", SheetVisibility::Hidden)
        .expect("set sheet2 hidden");
    state
        .set_sheet_tab_color("Sheet1", Some(TabColor::rgb("FFFF0000")))
        .expect("set sheet1 tab color");

    let workbook = state.get_workbook().expect("workbook should be loaded");
    assert!(
        workbook.origin_xlsx_bytes.is_some(),
        "expected sheet metadata edits to preserve origin_xlsx_bytes so save can patch in-place"
    );

    let out_path = dir.path().join("out.xlsx");
    let written =
        desktop::file_io::write_xlsx_blocking(&out_path, workbook).expect("write patched workbook");

    let out_pkg = formula_xlsx::XlsxPackage::from_bytes(written.as_ref()).expect("parse output pkg");
    assert!(
        out_pkg.part("customXml/item1.xml").is_some(),
        "expected custom part to be preserved by patch-based save path"
    );

    let roundtrip =
        formula_xlsx::read_workbook_from_reader(std::io::Cursor::new(written.as_ref()))
            .expect("read workbook model from patched bytes");
    assert_eq!(
        roundtrip
            .sheet_by_name("Sheet2")
            .expect("Sheet2 exists")
            .visibility,
        SheetVisibility::Hidden
    );
    assert_eq!(
        roundtrip
            .sheet_by_name("Sheet1")
            .expect("Sheet1 exists")
            .tab_color
            .as_ref()
            .and_then(|c| c.rgb.as_deref()),
        Some("FFFF0000")
    );
}

#[test]
fn sheet_metadata_restores_from_on_disk_autosave_db() {
    let dir = tempdir().expect("tempdir");
    let db_path = dir.path().join("autosave.sqlite");
    let location = WorkbookPersistenceLocation::OnDisk(db_path.clone());

    // First session: load a workbook into on-disk persistence and mutate sheet metadata.
    {
        let mut state = AppState::new();
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        // Simulate an imported tab color that the desktop UI/host APIs don't currently allow
        // setting directly (theme/tint-based colors can still be loaded and round-tripped).
        if let Some(sheet3) = workbook.sheet_mut("Sheet3") {
            sheet3.tab_color = Some(TabColor {
                theme: Some(1),
                tint: Some(0.5),
                ..Default::default()
            });
        }

        state
            .load_workbook_persistent(workbook, location.clone())
            .expect("load workbook");

        state
            .set_sheet_visibility("Sheet2", SheetVisibility::Hidden)
            .expect("set sheet2 hidden");
        state
            .set_sheet_tab_color("Sheet2", Some(TabColor::rgb("FF00FF00")))
            .expect("set sheet2 tab color");

        // `veryHidden` is not settable via the desktop UI, but the backend should be able to
        // persist + restore it for crash recovery / round-tripping.
        state
            .set_sheet_visibility("Sheet3", SheetVisibility::VeryHidden)
            .expect("set sheet3 veryHidden");
    }

    // Second session: load the same workbook against the existing autosave DB and verify the
    // metadata is restored from SQLite (not from the newly provided workbook template).
    let mut state = AppState::new();
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.add_sheet("Sheet2".to_string());
    workbook.add_sheet("Sheet3".to_string());

    state
        .load_workbook_persistent(workbook, location)
        .expect("load workbook (restore)");

    let info = state.workbook_info().expect("workbook_info");
    let sheet2 = info
        .sheets
        .iter()
        .find(|s| s.id == "Sheet2")
        .expect("sheet2 in workbook_info");
    assert_eq!(sheet2.visibility, SheetVisibility::Hidden);
    assert_eq!(sheet2.tab_color, Some(TabColor::rgb("FF00FF00")));

    let sheet3 = info
        .sheets
        .iter()
        .find(|s| s.id == "Sheet3")
        .expect("sheet3 in workbook_info");
    assert_eq!(sheet3.visibility, SheetVisibility::VeryHidden);
    assert_eq!(
        sheet3.tab_color,
        Some(TabColor {
            theme: Some(1),
            tint: Some(0.5),
            ..Default::default()
        })
    );
}
