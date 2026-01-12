use desktop::file_io::{read_xlsx_blocking, Workbook};
use desktop::persistence::{write_xlsx_from_storage, WorkbookPersistenceLocation};
use desktop::state::{AppState, CellScalar};
use formula_storage::{CellRange, CellValue, Storage};
use serde_json::json;
use std::io::Cursor;
use std::path::PathBuf;
use std::collections::BTreeSet;

#[tokio::test]
async fn autosave_persists_and_exports_round_trip() {
    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    let sheet_id = workbook.sheets[0].id.clone();

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path.clone()))
        .expect("load persistent workbook");

    state
        .set_cell(&sheet_id, 0, 0, Some(json!(42.0)), None)
        .expect("set A1");
    state
        .set_cell(&sheet_id, 0, 1, None, Some("=A1+1".to_string()))
        .expect("set B1 formula");
    assert_eq!(
        state.get_cell(&sheet_id, 0, 1).expect("get B1").value,
        CellScalar::Number(43.0)
    );

    let autosave = state.autosave_manager().expect("autosave manager");
    autosave.flush().await.expect("flush autosave");

    let storage = Storage::open_path(&db_path).expect("open storage");
    let workbooks = storage.list_workbooks().expect("list workbooks");
    assert_eq!(workbooks.len(), 1);
    let workbook_id = workbooks[0].id;

    let sheets = storage.list_sheets(workbook_id).expect("list sheets");
    let sheet_uuid = sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("sheet1 meta")
        .id;

    let persisted = storage
        .load_cells_in_range(sheet_uuid, CellRange::new(0, 0, 0, 1))
        .expect("load persisted cells");
    assert_eq!(persisted.len(), 2);
    assert!(
        persisted
            .iter()
            .any(|((r, c), snap)| *r == 0 && *c == 0 && snap.value == CellValue::Number(42.0)),
        "expected A1=42 in persisted cells, got {persisted:?}"
    );
    assert!(
        persisted.iter().any(|((r, c), snap)| {
            *r == 0 && *c == 1 && snap.formula.as_deref() == Some("A1+1")
        }),
        "expected B1 formula in persisted cells, got {persisted:?}"
    );

    // Clearing a cell should remove it from SQLite.
    state
        .set_cell(&sheet_id, 0, 0, None, None)
        .expect("clear A1");
    autosave.flush().await.expect("flush autosave after clear");
    assert_eq!(storage.cell_count(sheet_uuid).expect("cell count"), 1);

    // Recovery: load a new AppState from the same autosave DB.
    let mut template = Workbook::new_empty(None);
    template.add_sheet("Sheet1".to_string());
    let mut recovered = AppState::new();
    recovered
        .load_workbook_persistent(template, WorkbookPersistenceLocation::OnDisk(db_path.clone()))
        .expect("load recovered workbook");

    assert_eq!(
        recovered.get_cell(&sheet_id, 0, 0).expect("get A1").value,
        CellScalar::Empty
    );
    assert_eq!(
        recovered.get_cell(&sheet_id, 0, 1).expect("get B1").value,
        CellScalar::Number(1.0),
        "empty cells are treated as 0 in arithmetic"
    );

    // Export through the same storage→xlsx path as the desktop `save_workbook` command.
    let xlsx_path = tmp_dir.path().join("export.xlsx");
    let export_storage = recovered.persistent_storage().expect("storage handle");
    let export_id = recovered.persistent_workbook_id().expect("workbook id");
    let export_meta = recovered.get_workbook().expect("workbook").clone();
    recovered
        .autosave_manager()
        .expect("autosave")
        .flush()
        .await
        .expect("flush before export");

    write_xlsx_from_storage(&export_storage, export_id, &export_meta, &xlsx_path).expect("export xlsx");

    let reopened = read_xlsx_blocking(&xlsx_path).expect("read exported xlsx");
    let reopened_sheet_id = reopened.sheets[0].id.clone();
    let mut reopened_state = AppState::new();
    reopened_state.load_workbook(reopened);

    let cell = reopened_state
        .get_cell(&reopened_sheet_id, 0, 1)
        .expect("get B1 after reopen");
    assert_eq!(cell.formula.as_deref(), Some("=A1+1"));
    assert_eq!(cell.value, CellScalar::Number(1.0));
}

#[tokio::test]
async fn export_creates_parent_dirs_for_new_workbook() {
    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    let sheet_id = workbook.sheets[0].id.clone();
    assert!(
        workbook.origin_xlsx_bytes.is_none(),
        "expected new workbook to have no origin_xlsx_bytes"
    );

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path))
        .expect("load persistent workbook");

    state
        .set_cell(&sheet_id, 0, 0, Some(json!(1.0)), None)
        .expect("set A1");

    state
        .autosave_manager()
        .expect("autosave")
        .flush()
        .await
        .expect("flush autosave");

    let export_storage = state.persistent_storage().expect("storage handle");
    let export_id = state.persistent_workbook_id().expect("workbook id");
    let export_meta = state.get_workbook().expect("workbook").clone();
    assert!(
        export_meta.origin_xlsx_bytes.is_none(),
        "expected exported workbook metadata to have no origin_xlsx_bytes"
    );

    let xlsx_path = tmp_dir.path().join("nested").join("dir").join("export.xlsx");
    let parent = xlsx_path.parent().expect("has parent");
    assert!(!parent.exists(), "expected export parent dir to not exist yet");

    write_xlsx_from_storage(&export_storage, export_id, &export_meta, &xlsx_path).expect("export xlsx");

    assert!(parent.is_dir(), "expected export to create parent dirs");
    assert!(xlsx_path.is_file(), "expected export file to exist");
    let bytes = std::fs::read(&xlsx_path).expect("read exported xlsx bytes");
    assert!(
        bytes.starts_with(b"PK"),
        "expected exported xlsx to start with ZIP header, got: {:?}",
        bytes.get(..4)
    );
}

#[test]
fn xlsx_import_uses_origin_bytes_to_preserve_styles_in_storage() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../../fixtures/xlsx/styles/styles.xlsx");
    let workbook = read_xlsx_blocking(&fixture).expect("read fixture");
    assert!(
        workbook.origin_xlsx_bytes.is_some(),
        "fixture should load with origin_xlsx_bytes"
    );

    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path))
        .expect("load persistent workbook");

    let storage = state.persistent_storage().expect("storage");
    let workbook_id = state.persistent_workbook_id().expect("workbook id");
    let model = storage
        .export_model_workbook(workbook_id)
        .expect("export model");

    assert!(
        model.styles.len() > 1,
        "expected more than the default style in imported workbook"
    );
}

#[tokio::test]
async fn export_writes_cached_values_for_formula_cells() {
    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    let sheet_id = workbook.sheets[0].id.clone();

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path.clone()))
        .expect("load persistent workbook");

    state
        .set_cell(&sheet_id, 0, 0, Some(json!(2.0)), None)
        .expect("set A1");
    state
        .set_cell(&sheet_id, 0, 1, None, Some("=A1+1".to_string()))
        .expect("set B1 formula");

    state
        .autosave_manager()
        .expect("autosave")
        .flush()
        .await
        .expect("flush autosave");

    let export_storage = state.persistent_storage().expect("storage handle");
    let export_id = state.persistent_workbook_id().expect("workbook id");
    let export_meta = state.get_workbook().expect("workbook").clone();
    let xlsx_path = tmp_dir.path().join("export.xlsx");
    let bytes =
        write_xlsx_from_storage(&export_storage, export_id, &export_meta, &xlsx_path).expect("export xlsx");

    let model =
        formula_xlsx::read_workbook_from_reader(Cursor::new(bytes.as_ref())).expect("read exported model");
    let sheet = model.sheets.first().expect("sheet");
    let cell = sheet
        .cell(formula_model::CellRef::new(0, 1))
        .expect("B1 cell should exist");

    assert_eq!(
        cell.value,
        formula_model::CellValue::Number(3.0),
        "expected cached value for B1"
    );
}

#[tokio::test]
async fn export_xltx_enforces_template_content_type_and_writes_print_settings() {
    use formula_xlsx::print::{
        CellRange as PrintCellRange, ManualPageBreaks, Orientation, PageMargins, PageSetup, PaperSize,
        Scaling, SheetPrintSettings, WorkbookPrintSettings,
    };
    use formula_xlsx::WorkbookKind;

    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    // Simulate a workbook that currently has macros loaded in memory (e.g. from `.xltm`) but the
    // user saves as `.xltx`.
    workbook.vba_project_bin = Some(b"fake-vba-project".to_vec());

    workbook.print_settings = WorkbookPrintSettings {
        sheets: vec![SheetPrintSettings {
            sheet_name: "Sheet1".to_string(),
            print_area: Some(vec![PrintCellRange {
                start_row: 1,
                end_row: 2,
                start_col: 1,
                end_col: 2,
            }]),
            print_titles: None,
            page_setup: PageSetup {
                orientation: Orientation::Landscape,
                paper_size: PaperSize::A4,
                margins: PageMargins {
                    left: 0.5,
                    right: 0.5,
                    top: 1.0,
                    bottom: 1.0,
                    header: 0.25,
                    footer: 0.25,
                },
                scaling: Scaling::FitTo { width: 1, height: 2 },
            },
            manual_page_breaks: ManualPageBreaks {
                row_breaks_after: BTreeSet::from([1]),
                col_breaks_after: BTreeSet::new(),
            },
        }],
    };

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path.clone()))
        .expect("load persistent workbook");

    state
        .autosave_manager()
        .expect("autosave")
        .flush()
        .await
        .expect("flush autosave");

    let export_storage = state.persistent_storage().expect("storage handle");
    let export_id = state.persistent_workbook_id().expect("workbook id");
    let export_meta = state.get_workbook().expect("workbook").clone();

    let out_path = tmp_dir.path().join("export.xltx");
    let bytes =
        write_xlsx_from_storage(&export_storage, export_id, &export_meta, &out_path).expect("export xltx");

    let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref()).expect("parse package");
    assert!(
        pkg.part("xl/vbaProject.bin").is_none(),
        "expected .xltx export to not contain xl/vbaProject.bin"
    );

    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        ct.contains(WorkbookKind::Template.workbook_content_type()),
        "expected template workbook content type, got:\n{ct}"
    );

    let read_settings =
        formula_xlsx::print::read_workbook_print_settings(bytes.as_ref()).expect("read print settings");
    assert_eq!(
        read_settings, export_meta.print_settings,
        "expected print settings to round-trip through storage export"
    );
}

#[tokio::test]
async fn export_xltm_enforces_macro_enabled_template_content_type_and_preserves_vba() {
    use formula_xlsx::WorkbookKind;

    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.vba_project_bin = Some(b"fake-vba-project".to_vec());

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path.clone()))
        .expect("load persistent workbook");

    state
        .autosave_manager()
        .expect("autosave")
        .flush()
        .await
        .expect("flush autosave");

    let export_storage = state.persistent_storage().expect("storage handle");
    let export_id = state.persistent_workbook_id().expect("workbook id");
    let export_meta = state.get_workbook().expect("workbook").clone();

    let out_path = tmp_dir.path().join("export.xltm");
    let bytes =
        write_xlsx_from_storage(&export_storage, export_id, &export_meta, &out_path).expect("export xltm");

    let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref()).expect("parse package");
    assert!(
        pkg.part("xl/vbaProject.bin").is_some(),
        "expected .xltm export to contain xl/vbaProject.bin"
    );

    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        ct.contains(WorkbookKind::MacroEnabledTemplate.workbook_content_type()),
        "expected macro-enabled template workbook content type, got:\n{ct}"
    );
}

#[tokio::test]
async fn export_xlam_enforces_addin_content_type() {
    use formula_xlsx::WorkbookKind;

    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.vba_project_bin = Some(b"fake-vba-project".to_vec());

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path.clone()))
        .expect("load persistent workbook");

    state
        .autosave_manager()
        .expect("autosave")
        .flush()
        .await
        .expect("flush autosave");

    let export_storage = state.persistent_storage().expect("storage handle");
    let export_id = state.persistent_workbook_id().expect("workbook id");
    let export_meta = state.get_workbook().expect("workbook").clone();

    let out_path = tmp_dir.path().join("export.xlam");
    let bytes =
        write_xlsx_from_storage(&export_storage, export_id, &export_meta, &out_path).expect("export xlam");

    let pkg = formula_xlsx::XlsxPackage::from_bytes(bytes.as_ref()).expect("parse package");
    assert!(
        pkg.part("xl/vbaProject.bin").is_some(),
        "expected .xlam export to contain xl/vbaProject.bin"
    );

    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        ct.contains(WorkbookKind::MacroEnabledAddIn.workbook_content_type()),
        "expected add-in workbook content type, got:\n{ct}"
    );
}

#[tokio::test]
async fn export_preserves_vba_project_signature_part_when_present() {
    use formula_xlsx::WorkbookKind;

    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    // Start from a real XLSM fixture so the workbook model contains a VBA project.
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../../fixtures/xlsx/macros/basic.xlsm");
    let bytes = std::fs::read(&fixture).expect("read fixture bytes");

    // Inject a fake `xl/vbaProjectSignature.bin` part so we can validate it round-trips through
    // the storage-based export path.
    let signature_bytes = b"fake-vba-project-signature".to_vec();
    let mut pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("parse fixture package");
    pkg.set_part("xl/vbaProjectSignature.bin", signature_bytes.clone());
    let signed_bytes = pkg.write_to_bytes().expect("write signed package bytes");

    let signed_path = tmp_dir.path().join("signed.xlsm");
    std::fs::write(&signed_path, &signed_bytes).expect("write signed xlsm");

    let workbook = read_xlsx_blocking(&signed_path).expect("read signed workbook");
    assert_eq!(
        workbook.vba_project_signature_bin.as_deref(),
        Some(signature_bytes.as_slice()),
        "expected read_xlsx_blocking to preserve xl/vbaProjectSignature.bin"
    );

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path.clone()))
        .expect("load persistent workbook");

    state
        .autosave_manager()
        .expect("autosave")
        .flush()
        .await
        .expect("flush autosave");

    let export_storage = state.persistent_storage().expect("storage handle");
    let export_id = state.persistent_workbook_id().expect("workbook id");
    let export_meta = state.get_workbook().expect("workbook").clone();

    let out_path = tmp_dir.path().join("export.xltm");
    let out_bytes =
        write_xlsx_from_storage(&export_storage, export_id, &export_meta, &out_path).expect("export");

    let out_pkg = formula_xlsx::XlsxPackage::from_bytes(out_bytes.as_ref()).expect("parse output");
    assert_eq!(
        out_pkg.vba_project_signature_bin(),
        Some(signature_bytes.as_slice()),
        "expected storage export to preserve xl/vbaProjectSignature.bin"
    );

    let ct = std::str::from_utf8(out_pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        ct.contains(WorkbookKind::MacroEnabledTemplate.workbook_content_type()),
        "expected .xltm workbook content type, got:\n{ct}"
    );
}
