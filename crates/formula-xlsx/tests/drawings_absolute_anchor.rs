use std::collections::BTreeMap;

use formula_model::drawings::{Anchor, CellOffset, DrawingObjectId, DrawingObjectKind, EmuSize, ImageId};
use formula_xlsx::drawings::DrawingPart;

#[test]
fn parse_absolute_anchor_drawing_part() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="123" y="456"/>
    <xdr:ext cx="789" cy="1011"/>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="1" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    // `DrawingPart::parse_from_parts` requires the `.rels` file, even if it is empty.
    let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

    let parts: BTreeMap<String, Vec<u8>> = [
        ("xl/drawings/drawing1.xml".to_string(), drawing_xml.as_bytes().to_vec()),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_string(),
            rels_xml.as_bytes().to_vec(),
        ),
    ]
    .into_iter()
    .collect();

    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        &parts,
        &mut workbook,
    )
    .expect("parse drawing part with absoluteAnchor");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(123, 456),
            ext: EmuSize::new(789, 1011),
        }
    );
    assert_eq!(drawing.objects[0].size, Some(EmuSize::new(789, 1011)));
}

#[test]
fn parse_absolute_anchor_picture_drawing_part() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="10" y="20"/>
    <xdr:ext cx="30" cy="40"/>
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
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="../media/image1.png"/>
</Relationships>"#;

    let parts: BTreeMap<String, Vec<u8>> = [
        ("xl/drawings/drawing1.xml".to_string(), drawing_xml.as_bytes().to_vec()),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_string(),
            rels_xml.as_bytes().to_vec(),
        ),
        // The XLSX reader does not validate image bytes; any payload is sufficient.
        ("xl/media/image1.png".to_string(), b"fake png".to_vec()),
    ]
    .into_iter()
    .collect();

    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        &parts,
        &mut workbook,
    )
    .expect("parse absoluteAnchor pic drawing part");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(10, 20),
            ext: EmuSize::new(30, 40),
        }
    );

    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::Image { image_id } if image_id == &ImageId::new("image1.png")
    ));
    assert!(workbook.images.get(&ImageId::new("image1.png")).is_some());
}

#[test]
fn parse_absolute_anchor_shape_trims_cnvpr_id_in_parts() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:absoluteAnchor>
    <xdr:pos x="10" y="20"/>
    <xdr:ext cx="30" cy="40"/>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id=" 7 " name="Shape 7"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let parts: BTreeMap<String, Vec<u8>> =
        [("xl/drawings/drawing1.xml".to_string(), drawing_xml.as_bytes().to_vec())]
            .into_iter()
            .collect();

    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        &parts,
        &mut workbook,
    )
    .expect("parse drawing part");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(10, 20),
            ext: EmuSize::new(30, 40),
        }
    );
    assert_eq!(drawing.objects[0].id, DrawingObjectId(7));
    assert_eq!(drawing.objects[0].size, Some(EmuSize::new(30, 40)));
    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::Shape { raw_xml } if raw_xml.contains("Shape 7")
    ));
    assert!(workbook.images.is_empty());
}

#[test]
fn parse_absolute_anchor_picture_drawing_part_from_archive() {
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="10" y="20"/>
    <xdr:ext cx="30" cy="40"/>
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
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="../media/image1.png"/>
</Relationships>"#;

    let options =
        FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    zip.start_file("xl/drawings/drawing1.xml", options)
        .expect("start drawing");
    zip.write_all(drawing_xml.as_bytes())
        .expect("write drawing xml");
    zip.start_file("xl/drawings/_rels/drawing1.xml.rels", options)
        .expect("start rels");
    zip.write_all(rels_xml.as_bytes()).expect("write rels xml");
    zip.start_file("xl/media/image1.png", options)
        .expect("start image");
    zip.write_all(b"fake png bytes").expect("write image");
    let bytes = zip.finish().expect("finish zip").into_inner();

    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_archive(
        0,
        "xl/drawings/drawing1.xml",
        &mut archive,
        &mut workbook,
    )
    .expect("parse drawing part from archive");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(10, 20),
            ext: EmuSize::new(30, 40),
        }
    );
    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::Image { image_id } if image_id == &ImageId::new("image1.png")
    ));
    assert!(workbook.images.get(&ImageId::new("image1.png")).is_some());
}

#[test]
fn parse_absolute_anchor_shape_trims_cnvpr_id_in_archive() {
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:absoluteAnchor>
    <xdr:pos x="10" y="20"/>
    <xdr:ext cx="30" cy="40"/>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id=" 7 " name="Shape 7"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let options =
        FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    zip.start_file("xl/drawings/drawing1.xml", options)
        .expect("start drawing");
    zip.write_all(drawing_xml.as_bytes())
        .expect("write drawing xml");
    let bytes = zip.finish().expect("finish zip").into_inner();

    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_archive(
        0,
        "xl/drawings/drawing1.xml",
        &mut archive,
        &mut workbook,
    )
    .expect("parse drawing part from archive");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(10, 20),
            ext: EmuSize::new(30, 40),
        }
    );
    assert_eq!(drawing.objects[0].id, DrawingObjectId(7));
    assert_eq!(drawing.objects[0].size, Some(EmuSize::new(30, 40)));
    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::Shape { raw_xml } if raw_xml.contains("Shape 7")
    ));
    assert!(workbook.images.is_empty());
}

#[test]
fn parse_absolute_anchor_chart_drawing_part_from_archive() {
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="123" y="456"/>
    <xdr:ext cx="789" cy="1011"/>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="1" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let options =
        FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    zip.start_file("xl/drawings/drawing1.xml", options)
        .expect("start drawing");
    zip.write_all(drawing_xml.as_bytes())
        .expect("write drawing xml");
    let bytes = zip.finish().expect("finish zip").into_inner();

    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_archive(
        0,
        "xl/drawings/drawing1.xml",
        &mut archive,
        &mut workbook,
    )
    .expect("parse drawing part from archive");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(123, 456),
            ext: EmuSize::new(789, 1011),
        }
    );
    assert_eq!(drawing.objects[0].size, Some(EmuSize::new(789, 1011)));
    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::ChartPlaceholder { rel_id, .. } if rel_id == "rId1"
    ));
}

#[test]
fn parse_absolute_anchor_picture_missing_relationship_falls_back_to_unknown_in_archive() {
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="10" y="20"/>
    <xdr:ext cx="30" cy="40"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="5" name="Picture 5"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rIdMissing"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let options =
        FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    zip.start_file("xl/drawings/drawing1.xml", options)
        .expect("start drawing");
    zip.write_all(drawing_xml.as_bytes())
        .expect("write drawing xml");
    let bytes = zip.finish().expect("finish zip").into_inner();

    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_archive(
        0,
        "xl/drawings/drawing1.xml",
        &mut archive,
        &mut workbook,
    )
    .expect("parse drawing part from archive");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(10, 20),
            ext: EmuSize::new(30, 40),
        }
    );
    assert_eq!(drawing.objects[0].id, DrawingObjectId(5));
    assert_eq!(drawing.objects[0].size, Some(EmuSize::new(30, 40)));
    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::Unknown { raw_xml } if raw_xml.contains("rIdMissing")
    ));
    assert!(workbook.images.is_empty());
}

#[test]
fn parse_absolute_anchor_malformed_shape_and_frame_preserve_size_in_archive() {
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="0" y="0"/>
    <xdr:ext cx="30" cy="40"/>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="bad" name="Shape 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
  <xdr:absoluteAnchor>
    <xdr:pos x="1" y="2"/>
    <xdr:ext cx="50" cy="60"/>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="bad2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let options =
        FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    zip.start_file("xl/drawings/drawing1.xml", options)
        .expect("start drawing");
    zip.write_all(drawing_xml.as_bytes())
        .expect("write drawing xml");
    let bytes = zip.finish().expect("finish zip").into_inner();

    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_archive(
        0,
        "xl/drawings/drawing1.xml",
        &mut archive,
        &mut workbook,
    )
    .expect("parse drawing part from archive");

    assert_eq!(drawing.objects.len(), 2);

    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(0, 0),
            ext: EmuSize::new(30, 40),
        }
    );
    assert_eq!(drawing.objects[0].id, DrawingObjectId(1));
    assert_eq!(drawing.objects[0].size, Some(EmuSize::new(30, 40)));
    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::Unknown { raw_xml } if raw_xml.contains("id=\"bad\"")
    ));

    assert_eq!(
        drawing.objects[1].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(1, 2),
            ext: EmuSize::new(50, 60),
        }
    );
    assert_eq!(drawing.objects[1].id, DrawingObjectId(2));
    assert_eq!(drawing.objects[1].size, Some(EmuSize::new(50, 60)));
    assert!(matches!(
        &drawing.objects[1].kind,
        DrawingObjectKind::Unknown { raw_xml } if raw_xml.contains("id=\"bad2\"")
    ));

    assert!(workbook.images.is_empty());
}

#[test]
fn parse_absolute_anchor_unknown_object_preserves_id_in_parts() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="10" y="20"/>
    <xdr:ext cx="30" cy="40"/>
    <xdr:cxnSp>
      <!-- non-canonical a:cNvPr before the canonical xdr:cNvPr -->
      <a:cNvPr id="999" name="Wrong"/>
      <xdr:nvCxnSpPr>
        <xdr:cNvPr id="2" name="Connector 1"/>
        <xdr:cNvCxnSpPr/>
      </xdr:nvCxnSpPr>
      <xdr:spPr/>
    </xdr:cxnSp>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let parts: BTreeMap<String, Vec<u8>> =
        [("xl/drawings/drawing1.xml".to_string(), drawing_xml.as_bytes().to_vec())]
            .into_iter()
            .collect();

    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        &parts,
        &mut workbook,
    )
    .expect("parse drawing part");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(10, 20),
            ext: EmuSize::new(30, 40),
        }
    );
    assert_eq!(drawing.objects[0].id, DrawingObjectId(2));
    assert_eq!(drawing.objects[0].size, Some(EmuSize::new(30, 40)));
    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::Unknown { raw_xml } if raw_xml.contains("cxnSp")
    ));
    assert!(workbook.images.is_empty());
}

#[test]
fn parse_absolute_anchor_unknown_object_preserves_size_in_archive() {
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="10" y="20"/>
    <xdr:ext cx="30" cy="40"/>
    <xdr:cxnSp>
      <!-- non-canonical a:cNvPr before the canonical xdr:cNvPr -->
      <a:cNvPr id="999" name="Wrong"/>
      <xdr:nvCxnSpPr>
        <xdr:cNvPr id="2" name="Connector 1"/>
        <xdr:cNvCxnSpPr/>
      </xdr:nvCxnSpPr>
      <xdr:spPr/>
    </xdr:cxnSp>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let options =
        FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    zip.start_file("xl/drawings/drawing1.xml", options)
        .expect("start drawing");
    zip.write_all(drawing_xml.as_bytes())
        .expect("write drawing xml");
    let bytes = zip.finish().expect("finish zip").into_inner();

    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_archive(
        0,
        "xl/drawings/drawing1.xml",
        &mut archive,
        &mut workbook,
    )
    .expect("parse drawing part from archive");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(10, 20),
            ext: EmuSize::new(30, 40),
        }
    );
    assert_eq!(drawing.objects[0].id, DrawingObjectId(2));
    assert_eq!(drawing.objects[0].size, Some(EmuSize::new(30, 40)));
    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::Unknown { raw_xml } if raw_xml.contains("cxnSp")
    ));
    assert!(workbook.images.is_empty());
}

#[test]
fn parse_absolute_anchor_malformed_pic_preserves_id_in_archive() {
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="10" y="20"/>
    <xdr:ext cx="30" cy="40"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="7" name="Picture 7"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    let options =
        FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    zip.start_file("xl/drawings/drawing1.xml", options)
        .expect("start drawing");
    zip.write_all(drawing_xml.as_bytes())
        .expect("write drawing xml");
    let bytes = zip.finish().expect("finish zip").into_inner();

    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_archive(
        0,
        "xl/drawings/drawing1.xml",
        &mut archive,
        &mut workbook,
    )
    .expect("parse drawing part from archive");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(drawing.objects[0].id, DrawingObjectId(7));
    assert_eq!(drawing.objects[0].size, Some(EmuSize::new(30, 40)));
    assert!(matches!(
        &drawing.objects[0].kind,
        DrawingObjectKind::Unknown { raw_xml } if raw_xml.contains("Picture 7")
    ));
    assert!(workbook.images.is_empty());
}
