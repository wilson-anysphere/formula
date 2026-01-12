use base64::Engine as _;

#[test]
fn embedded_cell_images_strip_uri_fragments_in_relationship_targets() {
    // Use rust_xlsxwriter to generate a real embedded-image-in-cell workbook, then mutate the
    // richValueRel relationships to include a URI fragment in the image Target. The extractor must
    // strip the fragment when resolving OPC part names.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let image = rust_xlsxwriter::Image::new_from_buffer(&png_bytes).expect("image from buffer");
    worksheet
        .embed_image(0, 0, &image)
        .expect("embed image into A1");

    let bytes = workbook.save_to_buffer().expect("save workbook");
    let bytes = rewrite_zip_part(&bytes, "xl/richData/_rels/richValueRel.xml.rels", |rels| {
        let xml = std::str::from_utf8(rels).expect("rels xml utf-8");
        let doc = roxmltree::Document::parse(xml).expect("parse rels xml");
        let ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        let relationship = doc
            .descendants()
            .find(|n| {
                n.has_tag_name((ns, "Relationship"))
                    && n.attribute("Type")
                        .is_some_and(|t| t.ends_with("/relationships/image"))
            })
            .expect("expected an image relationship");
        let target = relationship
            .attribute("Target")
            .expect("expected image relationship Target attribute");
        assert!(
            !target.contains('#'),
            "expected image relationship Target to have no URI fragment, got: {target}"
        );
        // Append a fragment to the image target.
        let patched = xml.replacen(target, &format!("{target}#fragment"), 1);
        assert!(
            patched.contains("#fragment"),
            "expected patched rels xml to contain #fragment, got: {patched}"
        );
        patched.into_bytes()
    });

    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("read package");
    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");

    let key = (
        "xl/worksheets/sheet1.xml".to_string(),
        formula_model::CellRef::from_a1("A1").unwrap(),
    );
    let image = images.get(&key).expect("expected embedded image at A1");
    assert_eq!(image.image_part, "xl/media/image1.png");
    assert_eq!(image.image_bytes, png_bytes);
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
