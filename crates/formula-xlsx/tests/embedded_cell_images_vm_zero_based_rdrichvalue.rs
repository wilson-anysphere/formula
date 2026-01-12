use base64::Engine as _;
use formula_model::CellRef;
use formula_xlsx::XlsxPackage;
use roxmltree::Document;
use rust_xlsxwriter::{Image, Workbook};

#[test]
fn embedded_cell_images_supports_zero_based_vm_with_rdrichvalue_parts() {
    // Excel has been observed to emit 0-based worksheet `c/@vm` values even when `xl/metadata.xml`
    // uses the typical 1-based indexing for `<valueMetadata><bk>` records.
    //
    // `extract_embedded_cell_images` has a fast-path when the full RichData pipeline is present
    // (`xl/metadata.xml` + `xl/richData/rdrichvalue.xml` + `xl/richData/rdrichvaluestructure.xml`).
    // Ensure that path also tolerates zero-based worksheet `vm` values.

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
    let bytes = rewrite_zip_part(&bytes, "xl/worksheets/sheet1.xml", |sheet| {
        let xml = std::str::from_utf8(sheet).expect("sheet xml utf-8");
        assert!(
            xml.contains("vm=\""),
            "expected generated worksheet to contain vm attributes, got: {xml}"
        );
        let patched = decrement_vm_attributes(xml);
        assert!(
            patched.contains("vm=\"0\""),
            "expected patched worksheet to contain vm=\"0\", got: {patched}"
        );
        patched.into_bytes()
    });
    let bytes = rewrite_zip_part(&bytes, "xl/richData/_rels/richValueRel.xml.rels", |rels| {
        // Some producers emit `Target="xl/media/image1.png"` (missing the leading `/`), which would
        // otherwise resolve relative to `xl/richData/`.
        let xml = std::str::from_utf8(rels).expect("rels xml utf-8");
        let doc = Document::parse(xml).expect("parse rels xml");
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
        let patched = xml.replacen(target, "xl/media/image1.png", 1);
        assert!(
            patched.contains("xl/media/image1.png"),
            "expected patched rels xml to contain xl/media/image1.png, got: {patched}"
        );
        patched.into_bytes()
    });

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    assert!(pkg.part("xl/metadata.xml").is_some(), "expected xl/metadata.xml");
    assert!(
        pkg.part("xl/richData/rdrichvalue.xml").is_some(),
        "expected xl/richData/rdrichvalue.xml"
    );
    assert!(
        pkg.part("xl/richData/rdrichvaluestructure.xml").is_some(),
        "expected xl/richData/rdrichvaluestructure.xml"
    );

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");

    let key = (
        "xl/worksheets/sheet1.xml".to_string(),
        CellRef::from_a1("A1").unwrap(),
    );
    let img = images.get(&key).expect("expected embedded image at A1");
    assert_eq!(img.image_part, "xl/media/image1.png");
    assert_eq!(img.image_bytes, png_bytes);
}

fn decrement_vm_attributes(xml: &str) -> String {
    // Best-effort string rewrite: replace `vm="N"` with `vm="N-1"` for all digit-only values.
    //
    // This avoids depending on a full XML rewriter in tests while still producing realistic
    // "zero-based vm" scenarios.
    let mut out = String::with_capacity(xml.len());
    let mut cursor = 0usize;
    while let Some(idx) = xml[cursor..].find("vm=\"") {
        let start = cursor + idx;
        let value_start = start + "vm=\"".len();
        out.push_str(&xml[cursor..value_start]);

        let bytes = xml.as_bytes();
        let mut end = value_start;
        while end < bytes.len() && bytes[end].is_ascii_digit() {
            end += 1;
        }

        if end < bytes.len() && bytes[end] == b'"' {
            if let Ok(v) = xml[value_start..end].parse::<u32>() {
                out.push_str(&v.saturating_sub(1).to_string());
                cursor = end;
                continue;
            }
        }

        // Fallback: copy through without rewriting.
        out.push_str(&xml[value_start..end]);
        cursor = end;
    }
    out.push_str(&xml[cursor..]);
    out
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
