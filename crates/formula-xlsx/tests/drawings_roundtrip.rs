use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};

use base64::Engine;
use formula_model::drawings::{Anchor, AnchorPoint, CellOffset, EmuSize, ImageData, ImageId};
use formula_model::CellRef;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::XlsxPackage;

fn build_fixture_xlsx() -> Vec<u8> {
    // 1x1 transparent PNG.
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
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>
"#
            .to_vec(),
        ),
        (
            "_rels/.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#
            .to_vec(),
        ),
        (
            "xl/workbook.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#
            .to_vec(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"#
            .to_vec(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>
"#
            .to_vec(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
"#
            .to_vec(),
        ),
        (
            "xl/drawings/drawing1.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>3</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>5</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="Picture 1"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="952500" cy="952500"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="Shape 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>5</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>6</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:cxnSp>
      <xdr:nvCxnSpPr>
        <xdr:cNvPr id="3" name="UnknownConnector"/>
        <xdr:cNvCxnSpPr/>
      </xdr:nvCxnSpPr>
      <xdr:spPr>
        <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:cxnSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"#
            .to_vec(),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"#
            .to_vec(),
        ),
        ("xl/media/image1.png".to_string(), png_bytes),
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

    let cursor = writer.finish().unwrap();
    cursor.into_inner()
}

fn unzip_part(zip_bytes: &[u8], path: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).unwrap();
    let mut file = archive.by_name(path).unwrap();
    let mut out = Vec::new();
    file.read_to_end(&mut out).unwrap();
    out
}

#[test]
fn drawings_import_round_trip_preserves_media_and_unknown_elements() {
    let bytes = build_fixture_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("load fixture");
    let mut workbook = formula_model::Workbook::new();
    formula_xlsx::drawings::load_media_parts(&mut workbook, pkg.parts_map());

    let image_id = ImageId::new("image1.png");
    assert_eq!(
        workbook.images.get(&image_id).unwrap().bytes,
        unzip_part(&bytes, "xl/media/image1.png")
    );

    let part = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        pkg.parts_map(),
        &mut workbook,
    )
    .expect("parse drawing part");

    let drawings = &part.objects;
    assert!(drawings.iter().any(|d| matches!(
        d.kind,
        formula_model::drawings::DrawingObjectKind::Image { .. }
    )));
    assert!(drawings.iter().any(|d| matches!(
        d.kind,
        formula_model::drawings::DrawingObjectKind::Shape { .. }
    )));
    assert!(drawings.iter().any(|d| matches!(
        d.kind,
        formula_model::drawings::DrawingObjectKind::Unknown { .. }
    )));

    // The first object is the image from B2 (1,1) to D6 (5,3).
    let img = drawings
        .iter()
        .find(|d| {
            matches!(
                d.kind,
                formula_model::drawings::DrawingObjectKind::Image { .. }
            )
        })
        .unwrap();
    assert_eq!(img.size, Some(EmuSize::new(952_500, 952_500)));
    match img.anchor {
        Anchor::TwoCell { from, to } => {
            assert_eq!(from.cell, CellRef::new(1, 1));
            assert_eq!(to.cell, CellRef::new(5, 3));
        }
        other => panic!("unexpected image anchor: {other:?}"),
    }

    // Round-trip.
    let mut part = part;
    part.write_into_parts(pkg.parts_map_mut(), &workbook)
        .expect("write drawing part");
    let out = pkg.write_to_bytes().expect("save");
    let pkg2 = XlsxPackage::from_bytes(&out).expect("reload");

    let mut workbook2 = formula_model::Workbook::new();
    formula_xlsx::drawings::load_media_parts(&mut workbook2, pkg2.parts_map());

    assert_eq!(
        workbook2.images.get(&image_id).unwrap().bytes,
        unzip_part(&bytes, "xl/media/image1.png")
    );

    // Unknown anchor is preserved by emitting its original subtree unchanged.
    let drawing_xml = String::from_utf8(unzip_part(&out, "xl/drawings/drawing1.xml")).unwrap();
    assert!(drawing_xml.contains("UnknownConnector"));
    assert!(drawing_xml.contains("<xdr:cxnSp>"));
}

#[test]
fn insert_image_and_save_keeps_bytes() {
    let bytes = build_fixture_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("load fixture");
    let mut workbook = formula_model::Workbook::new();
    formula_xlsx::drawings::load_media_parts(&mut workbook, pkg.parts_map());

    let mut part = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        pkg.parts_map(),
        &mut workbook,
    )
    .expect("parse drawing part");

    let new_png = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAQAAADZc7J/AAAADElEQVR42mP8z8BQDwAF9QH5m2n1LwAAAABJRU5ErkJggg==")
        .expect("valid base64 png");

    let inserted_id = workbook.images.ensure_unique_name("image", "png");
    workbook.images.insert(
        inserted_id.clone(),
        ImageData {
            bytes: new_png.clone(),
            content_type: Some("image/png".to_string()),
        },
    );

    let anchor = Anchor::OneCell {
        from: AnchorPoint::new(CellRef::new(2, 2), CellOffset::new(0, 0)),
        ext: EmuSize::new(914_400, 914_400),
    };
    part.insert_image_object(&inserted_id, anchor);
    part.write_into_parts(pkg.parts_map_mut(), &workbook)
        .expect("write drawing part");

    let saved = pkg.write_to_bytes().expect("save with inserted image");
    let pkg2 = XlsxPackage::from_bytes(&saved).expect("reload saved");

    let mut workbook2 = formula_model::Workbook::new();
    formula_xlsx::drawings::load_media_parts(&mut workbook2, pkg2.parts_map());

    assert_eq!(
        workbook2.images.get(&inserted_id).unwrap().bytes,
        new_png
    );

    let part2 = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        pkg2.parts_map(),
        &mut workbook2,
    )
    .expect("parse drawing part");
    assert!(part2.objects.iter().any(|d| {
        matches!(
            &d.kind,
            formula_model::drawings::DrawingObjectKind::Image { image_id }
                if image_id == &inserted_id
        )
    }));
}
