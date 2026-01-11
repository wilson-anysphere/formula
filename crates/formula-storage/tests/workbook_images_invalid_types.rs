use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn export_model_workbook_skips_invalid_workbook_images_rows() {
    let uri = "file:workbook_images_invalid_types?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    conn.execute(
        r#"
        INSERT INTO workbook_images (workbook_id, image_id, content_type, bytes)
        VALUES (?1, ?2, ?3, ?4)
        "#,
        rusqlite::params![
            workbook.id.to_string(),
            "valid",
            "image/png",
            vec![1u8, 2, 3]
        ],
    )
    .expect("insert valid workbook image");

    // Corrupt the `bytes` column by storing an INTEGER; rusqlite cannot deserialize this into a
    // `Vec<u8>` and export should ignore the row instead of failing the whole workbook export.
    conn.execute(
        r#"
        INSERT INTO workbook_images (workbook_id, image_id, content_type, bytes)
        VALUES (?1, ?2, ?3, ?4)
        "#,
        rusqlite::params![workbook.id.to_string(), "broken", "image/png", 123_i64],
    )
    .expect("insert invalid workbook image");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");

    let valid_id = formula_model::drawings::ImageId::new("valid".to_string());
    assert_eq!(exported.images.iter().count(), 1);
    let image = exported.images.get(&valid_id).expect("valid image loaded");
    assert_eq!(image.bytes, vec![1u8, 2, 3]);
    assert_eq!(image.content_type.as_deref(), Some("image/png"));
}
