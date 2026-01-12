use formula_storage::Storage;
use serde_json::json;

#[test]
fn sheet_metadata_roundtrips_via_set_sheet_metadata() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let payload = json!({
        "formula_ui_formatting": {
            "schemaVersion": 1,
            "defaultFormat": null,
            "rowFormats": [],
            "colFormats": [],
            "cellFormats": [
                { "row": 0, "col": 0, "format": { "font": { "bold": true } } }
            ]
        }
    });

    storage
        .set_sheet_metadata(sheet.id, Some(payload.clone()))
        .expect("set metadata");

    let fetched = storage.get_sheet_meta(sheet.id).expect("get sheet meta");
    assert_eq!(fetched.metadata, Some(payload.clone()));

    // Clearing should round-trip as `None`.
    storage
        .set_sheet_metadata(sheet.id, None)
        .expect("clear metadata");
    let cleared = storage
        .get_sheet_meta(sheet.id)
        .expect("get sheet meta (cleared)");
    assert!(cleared.metadata.is_none());
}
