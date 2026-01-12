use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;

#[test]
fn macro_strip_preserves_comment_vml_when_removing_ole_shape() {
    assert_mixed_vml_stripping(build_mixed_vml_package_with_o_relid());
}

#[test]
fn macro_strip_preserves_comment_vml_when_removing_ole_shape_with_r_id() {
    assert_mixed_vml_stripping(build_mixed_vml_package_with_r_id());
}

fn assert_mixed_vml_stripping(bytes: Vec<u8>) {
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("parse pkg");

    assert!(pkg.part("xl/embeddings/oleObject1.bin").is_some());
    assert!(pkg.part("xl/drawings/vmlDrawing1.vml").is_some());

    pkg.remove_vba_project().expect("strip macros");

    assert!(
        pkg.part("xl/embeddings/oleObject1.bin").is_none(),
        "expected embedded OLE binary to be deleted"
    );

    let vml_rels = std::str::from_utf8(
        pkg.part("xl/drawings/_rels/vmlDrawing1.vml.rels")
            .expect("vml rels present"),
    )
    .expect("vml rels utf8");
    assert!(
        !vml_rels.contains("rIdOle"),
        "expected OLE relationship to be removed from vmlDrawing rels"
    );

    let vml = std::str::from_utf8(
        pkg.part("xl/drawings/vmlDrawing1.vml")
            .expect("vml part present"),
    )
    .expect("vml utf8");
    assert!(
        !vml.contains("rIdOle"),
        "expected VML to no longer reference removed OLE relationship id"
    );
    assert!(
        !vml.contains("OLEObject"),
        "expected OLE subtree to be stripped from VML"
    );

    assert!(
        vml.contains("ObjectType=\"Note\""),
        "expected legacy comment note shape to remain in VML"
    );
    assert!(vml.contains("<x:Row>0</x:Row>"), "expected note row marker");
    assert!(
        vml.contains("<x:Column>0</x:Column>"),
        "expected note column marker"
    );

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())
        .expect("sheet1.xml utf8");
    assert!(
        sheet_xml.contains("legacyDrawing"),
        "expected sheet to keep legacyDrawing element"
    );
    assert!(
        sheet_xml.contains("rIdVml"),
        "expected sheet to keep legacyDrawing relationship id"
    );

    let sheet_rels = std::str::from_utf8(pkg.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap())
        .expect("sheet1 rels utf8");
    assert!(
        sheet_rels.contains("Id=\"rIdVml\""),
        "expected sheet relationship to vml drawing to remain"
    );
}

fn build_mixed_vml_package_with_o_relid() -> Vec<u8> {
    build_mixed_vml_package(r#"o:relid="rIdOle""#, false)
}

fn build_mixed_vml_package_with_r_id() -> Vec<u8> {
    build_mixed_vml_package(r#"r:id="rIdOle""#, true)
}

fn build_mixed_vml_package(ole_rel_attr: &str, include_r_ns: bool) -> Vec<u8> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <legacyDrawing r:id="rIdVml"/>
</worksheet>"#;

    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>"#;

    let r_ns = if include_r_ns {
        r#" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
    } else {
        ""
    };

    let vml_template = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel"__R_NS__>
  <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1"/>
  </o:shapelayout>
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
    <v:stroke joinstyle="miter"/>
    <v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  <v:shape id="_x0000_s1025" type="#_x0000_t202"
           style="position:absolute; margin-left:80pt;margin-top:5pt;width:104pt;height:64pt; z-index:1; visibility:hidden"
           fillcolor="#ffffe1" o:insetmode="auto">
    <v:fill color2="#ffffe1"/>
    <v:shadow on="t" color="black" obscured="t"/>
    <v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto">
      <div style="text-align:left"/>
    </v:textbox>
    <x:ClientData ObjectType="Note">
      <x:MoveWithCells/>
      <x:SizeWithCells/>
      <x:Anchor>1, 15, 0, 2, 3, 15, 5, 2</x:Anchor>
      <x:AutoFill>False</x:AutoFill>
      <x:Row>0</x:Row><x:Column>0</x:Column>
    </x:ClientData>
  </v:shape>
  <v:shape id="_x0000_s2048" type="#_x0000_t202"
           style="position:absolute;margin-left:0;margin-top:0;width:50pt;height:20pt;z-index:2">
    <o:OLEObject Type="Embed" ProgID="Forms.CommandButton.1" ShapeID="_x0000_s2048" DrawAspect="Content" ObjectID="_1234" __OLE_REL_ATTR__/>
    <x:ClientData ObjectType="Pict"/>
  </v:shape>
 </xml>"##;

    let vml = vml_template
        .replace("__R_NS__", r_ns)
        .replace("__OLE_REL_ATTR__", ole_rel_attr);

    let vml_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
        .unwrap();
    zip.write_all(sheet_rels.as_bytes()).unwrap();

    zip.start_file("xl/drawings/vmlDrawing1.vml", options)
        .unwrap();
    zip.write_all(vml.as_bytes()).unwrap();

    zip.start_file("xl/drawings/_rels/vmlDrawing1.vml.rels", options)
        .unwrap();
    zip.write_all(vml_rels.as_bytes()).unwrap();

    zip.start_file("xl/embeddings/oleObject1.bin", options)
        .unwrap();
    zip.write_all(b"OLE DATA").unwrap();

    // Include a stub vbaProject.bin so `remove_vba_project()` is exercised in a
    // macro-enabled-like package.
    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(b"VBA DATA").unwrap();

    zip.finish().unwrap().into_inner()
}
