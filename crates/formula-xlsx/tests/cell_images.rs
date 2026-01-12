use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use base64::Engine;
use formula_model::drawings::ImageId;
use zip::write::FileOptions;
use zip::ZipWriter;

use formula_xlsx::cell_images::CellImages;
use formula_xlsx::XlsxPackage;

fn build_fixture_xlsx(image_target: &str) -> (Vec<u8>, Vec<u8>) {
    // 1x1 transparent PNG (same as `drawings_roundtrip.rs`).
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let parts: BTreeMap<String, Vec<u8>> = [
        (
            "[Content_Types].xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>
</Types>
"#
            .to_vec(),
        ),
        (
            "xl/cellimages.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:cellImages xmlns:cx="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
               xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:pic>
    <xdr:nvPicPr>
      <xdr:cNvPr id="1" name="Picture 1"/>
      <xdr:cNvPicPr/>
    </xdr:nvPicPr>
    <xdr:blipFill>
      <a:blip r:embed="rId1"/>
    </xdr:blipFill>
  </xdr:pic>
</cx:cellImages>
"#
            .to_vec(),
        ),
        (
            "xl/_rels/cellimages.xml.rels".to_string(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{image_target}"/>
</Relationships>
"#
            )
            .into_bytes(),
        ),
        ("xl/media/image1.png".to_string(), png_bytes.clone()),
    ]
    .into_iter()
    .collect();

    let cursor = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts {
        writer.start_file(name, options).unwrap();
        writer.write_all(&bytes).unwrap();
    }

    (writer.finish().unwrap().into_inner(), png_bytes)
}

#[test]
fn cell_images_import_loads_referenced_media() {
    let (bytes, png_bytes) = build_fixture_xlsx("media/image1.png");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("load fixture");

    let mut workbook = formula_model::Workbook::new();
    let parsed =
        CellImages::parse_from_parts(pkg.parts_map(), &mut workbook).expect("parse cell images");

    assert_eq!(parsed.parts.len(), 1);
    assert_eq!(parsed.parts[0].path, "xl/cellimages.xml");
    assert_eq!(parsed.parts[0].rels_path, "xl/_rels/cellimages.xml.rels");
    assert_eq!(parsed.parts[0].images.len(), 1);

    let image = &parsed.parts[0].images[0];
    assert_eq!(image.embed_rel_id, "rId1");
    assert_eq!(image.target_path, "xl/media/image1.png");
    assert_eq!(image.image_id, ImageId::new("image1.png"));
    assert!(image.pic_xml.as_deref().unwrap_or_default().contains("<xdr:pic>"));

    assert_eq!(
        workbook
            .images
            .get(&ImageId::new("image1.png"))
            .expect("image stored")
            .bytes,
        png_bytes
    );
}

#[test]
fn cell_images_import_tolerates_parent_media_targets() {
    // Some producers may emit `../media/*` targets for workbook-level parts; tolerate it as a
    // best-effort fallback when the canonical resolution misses.
    let (bytes, png_bytes) = build_fixture_xlsx("../media/image1.png");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("load fixture");

    let mut workbook = formula_model::Workbook::new();
    let parsed =
        CellImages::parse_from_parts(pkg.parts_map(), &mut workbook).expect("parse cell images");

    assert_eq!(parsed.parts.len(), 1);
    assert_eq!(parsed.parts[0].images.len(), 1);
    assert_eq!(
        workbook
            .images
            .get(&ImageId::new("image1.png"))
            .expect("image stored")
            .bytes,
        png_bytes
    );
}

#[test]
fn cell_images_import_discovers_nested_cellimages_parts() {
    // 1x1 transparent PNG (same as `drawings_roundtrip.rs`).
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let parts: BTreeMap<String, Vec<u8>> = [
        (
            "xl/subdir/cellimages.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:cellImages xmlns:cx="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
               xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:pic>
    <xdr:nvPicPr>
      <xdr:cNvPr id="1" name="Picture 1"/>
      <xdr:cNvPicPr/>
    </xdr:nvPicPr>
    <xdr:blipFill>
      <a:blip r:embed="rId1"/>
    </xdr:blipFill>
  </xdr:pic>
</cx:cellImages>
"#
            .to_vec(),
        ),
        (
            "xl/subdir/_rels/cellimages.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"#
            .to_vec(),
        ),
        ("xl/media/image1.png".to_string(), png_bytes.clone()),
    ]
    .into_iter()
    .collect();

    let cursor = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts {
        writer.start_file(name, options).unwrap();
        writer.write_all(&bytes).unwrap();
    }

    let bytes = writer.finish().unwrap().into_inner();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("load fixture");

    let mut workbook = formula_model::Workbook::new();
    let parsed =
        CellImages::parse_from_parts(pkg.parts_map(), &mut workbook).expect("parse cell images");

    assert_eq!(parsed.parts.len(), 1);
    assert_eq!(parsed.parts[0].path, "xl/subdir/cellimages.xml");
    assert_eq!(
        parsed.parts[0].rels_path,
        "xl/subdir/_rels/cellimages.xml.rels"
    );
    assert_eq!(parsed.parts[0].images.len(), 1);

    assert_eq!(
        workbook
            .images
            .get(&ImageId::new("image1.png"))
            .expect("image stored")
            .bytes,
        png_bytes
    );
}
