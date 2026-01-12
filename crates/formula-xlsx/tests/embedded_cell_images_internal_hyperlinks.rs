use base64::Engine as _;
use formula_model::{CellRef, HyperlinkTarget};
use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Image, Workbook};

#[test]
fn embedded_cell_image_includes_internal_hyperlink_target() {
    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let image = Image::new_from_buffer(&png_bytes).expect("image from buffer");
    worksheet
        .embed_image(0, 0, &image)
        .expect("embed image into A1");

    let bytes = workbook.save_to_buffer().expect("save workbook");
    let bytes = rewrite_zip_part(&bytes, "xl/worksheets/sheet1.xml", |sheet_xml| {
        let xml = std::str::from_utf8(sheet_xml).expect("sheet xml utf-8");
        let patched = xml.replacen(
            "</worksheet>",
            "<hyperlinks><hyperlink ref=\"A1\" location=\"#Sheet1!B2\"/></hyperlinks></worksheet>",
            1,
        );
        assert!(
            patched.contains("location=\"#Sheet1!B2\""),
            "expected patched worksheet to contain internal hyperlink"
        );
        patched.into_bytes()
    });

    let pkg = XlsxPackage::from_bytes(&bytes).expect("load xlsx package");
    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");

    let key = (
        "xl/worksheets/sheet1.xml".to_string(),
        CellRef::from_a1("A1").unwrap(),
    );
    let image = images.get(&key).expect("expected image at Sheet1!A1");
    assert_eq!(image.image_bytes, png_bytes);
    assert_eq!(
        image.hyperlink_target,
        Some(HyperlinkTarget::Internal {
            sheet: "Sheet1".to_string(),
            cell: CellRef::from_a1("B2").unwrap(),
        })
    );
}

fn rewrite_zip_part(
    bytes: &[u8],
    part_name: &str,
    rewrite: impl FnOnce(&[u8]) -> Vec<u8>,
) -> Vec<u8> {
    use std::io::{Cursor, Read, Write};

    use zip::write::FileOptions;
    use zip::{ZipArchive, ZipWriter};

    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);

    let mut rewrite = Some(rewrite);
    let mut rewritten = false;
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        let name = file.name().to_string();
        let options = FileOptions::<()>::default().compression_method(file.compression());

        if file.is_dir() {
            zip.add_directory(name, options).expect("add dir");
            continue;
        }

        let mut data = Vec::new();
        file.read_to_end(&mut data).expect("read zip entry");
        if name == part_name {
            let f = rewrite.take().expect("rewrite function already used");
            data = f(&data);
            rewritten = true;
        }

        zip.start_file(name, options).expect("start file");
        zip.write_all(&data).expect("write zip entry");
    }

    assert!(rewritten, "expected to rewrite zip part {part_name}");
    zip.finish().expect("finish zip").into_inner()
}

