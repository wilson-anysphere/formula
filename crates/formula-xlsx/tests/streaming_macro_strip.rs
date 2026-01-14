use std::collections::BTreeSet;
use std::io::{Cursor, Write};

use formula_xlsx::{strip_vba_project_streaming, XlsxPackage};
use zip::write::FileOptions;
use zip::ZipWriter;

fn load_basic_fixture() -> Vec<u8> {
    std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/macros/basic.xlsm"
    ))
    .expect("fixture exists")
}

fn build_synthetic_macro_package() -> Vec<u8> {
    build_synthetic_macro_package_impl(false)
}

fn build_synthetic_macro_package_with_leading_slash_entries() -> Vec<u8> {
    build_synthetic_macro_package_impl(true)
}

fn build_synthetic_macro_package_with_backslash_and_case_entries() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>
  <Override PartName="/xl/controls/control1.xml" ContentType="application/vnd.ms-excel.control+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
  <Override PartName="/xl/vbaData.xml" ContentType="application/vnd.ms-office.vbaData+xml"/>
  <Override PartName="/customUI/customUI.xml" ContentType="application/xml"/>
  <Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2020/relationships/cellImage" Target="cellimages.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml#_x0000_s1025"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../controls/control1.xml"/>
</Relationships>"#;

    let vba_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let custom_ui_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"></customUI>"#;

    let custom_ui_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="image1.png"/>
</Relationships>"#;

    let activex_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"></ax:ocx>"#;

    let activex_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>"#;

    let ctrl_props_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ctrlProp xmlns="http://schemas.microsoft.com/office/2006/activeX"></ctrlProp>"#;

    let cellimages_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
</cellImages>"#;

    // Use `../media/*` to ensure macro stripping resolves relationship targets best-effort:
    // some producers emit workbook-level targets relative to the `.rels` directory.
    let cellimages_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let control_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<control xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Button1"/>"#;

    let control_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let noncanonical_name = |name: &str| name.replace('/', "\\").to_ascii_uppercase();

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file(noncanonical_name("[Content_Types].xml"), options)
        .unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("_rels/.rels"), options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("xl/workbook.xml"), options)
        .unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("xl/_rels/workbook.xml.rels"), options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("xl/worksheets/sheet1.xml"), options)
        .unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file(
        noncanonical_name("xl/worksheets/_rels/sheet1.xml.rels"),
        options,
    )
    .unwrap();
    zip.write_all(sheet_rels.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("xl/cellimages.xml"), options)
        .unwrap();
    zip.write_all(cellimages_xml).unwrap();

    zip.start_file(noncanonical_name("xl/_rels/cellimages.xml.rels"), options)
        .unwrap();
    zip.write_all(cellimages_rels).unwrap();

    zip.start_file(noncanonical_name("xl/media/image1.png"), options)
        .unwrap();
    zip.write_all(b"not-a-real-png").unwrap();

    zip.start_file(noncanonical_name("xl/controls/control1.xml"), options)
        .unwrap();
    zip.write_all(control_xml.as_bytes()).unwrap();

    zip.start_file(
        noncanonical_name("xl/controls/_rels/control1.xml.rels"),
        options,
    )
    .unwrap();
    zip.write_all(control_rels).unwrap();

    zip.start_file(noncanonical_name("customUI/customUI.xml"), options)
        .unwrap();
    zip.write_all(custom_ui_xml.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("customUI/customUI14.xml"), options)
        .unwrap();
    zip.write_all(custom_ui_xml.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("customUI/_rels/customUI.xml.rels"), options)
        .unwrap();
    zip.write_all(custom_ui_rels.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("customUI/image1.png"), options)
        .unwrap();
    zip.write_all(b"not-a-real-png").unwrap();

    zip.start_file(noncanonical_name("xl/vbaProject.bin"), options)
        .unwrap();
    zip.write_all(b"fake-vba-project").unwrap();

    zip.start_file(noncanonical_name("xl/_rels/vbaProject.bin.rels"), options)
        .unwrap();
    zip.write_all(vba_rels.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("xl/vbaProjectSignature.bin"), options)
        .unwrap();
    zip.write_all(b"fake-signature").unwrap();

    zip.start_file(noncanonical_name("xl/vbaData.xml"), options)
        .unwrap();
    zip.write_all(b"<vbaData/>").unwrap();

    zip.start_file(noncanonical_name("xl/activeX/activeX1.xml"), options)
        .unwrap();
    zip.write_all(activex_xml.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("xl/activeX/_rels/activeX1.xml.rels"), options)
        .unwrap();
    zip.write_all(activex_rels.as_bytes()).unwrap();

    zip.start_file(noncanonical_name("xl/activeX/activeX1.bin"), options)
        .unwrap();
    zip.write_all(b"activex-binary").unwrap();

    zip.start_file(noncanonical_name("xl/embeddings/oleObject1.bin"), options)
        .unwrap();
    zip.write_all(b"ole-embedding").unwrap();

    zip.start_file(noncanonical_name("xl/ctrlProps/ctrlProp1.xml"), options)
        .unwrap();
    zip.write_all(ctrl_props_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_synthetic_macro_package_impl(leading_slash_entries: bool) -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>
  <Override PartName="/xl/controls/control1.xml" ContentType="application/vnd.ms-excel.control+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
  <Override PartName="/xl/vbaData.xml" ContentType="application/vnd.ms-office.vbaData+xml"/>
  <Override PartName="/customUI/customUI.xml" ContentType="application/xml"/>
  <Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2020/relationships/cellImage" Target="cellimages.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml#_x0000_s1025"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../controls/control1.xml"/>
</Relationships>"#;

    let vba_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let custom_ui_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"></customUI>"#;

    let custom_ui_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="image1.png"/>
</Relationships>"#;

    let activex_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"></ax:ocx>"#;

    let activex_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>"#;

    let ctrl_props_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ctrlProp xmlns="http://schemas.microsoft.com/office/2006/activeX"></ctrlProp>"#;

    let cellimages_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
</cellImages>"#;

    // Use `../media/*` to ensure macro stripping resolves relationship targets best-effort:
    // some producers emit workbook-level targets relative to the `.rels` directory.
    let cellimages_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let control_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<control xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Button1"/>"#;

    let control_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    let part_name = |name: &str| {
        if leading_slash_entries {
            format!("/{name}")
        } else {
            name.to_string()
        }
    };

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file(part_name("xl/workbook.xml"), options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file(part_name("xl/_rels/workbook.xml.rels"), options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file(part_name("xl/worksheets/sheet1.xml"), options)
        .unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file(part_name("xl/worksheets/_rels/sheet1.xml.rels"), options)
        .unwrap();
    zip.write_all(sheet_rels.as_bytes()).unwrap();

    zip.start_file(part_name("xl/cellimages.xml"), options).unwrap();
    zip.write_all(cellimages_xml).unwrap();

    zip.start_file(part_name("xl/_rels/cellimages.xml.rels"), options)
        .unwrap();
    zip.write_all(cellimages_rels).unwrap();

    zip.start_file(part_name("xl/media/image1.png"), options).unwrap();
    zip.write_all(b"not-a-real-png").unwrap();

    zip.start_file(part_name("xl/controls/control1.xml"), options)
        .unwrap();
    zip.write_all(control_xml.as_bytes()).unwrap();

    zip.start_file(part_name("xl/controls/_rels/control1.xml.rels"), options)
        .unwrap();
    zip.write_all(control_rels).unwrap();

    zip.start_file(part_name("customUI/customUI.xml"), options)
        .unwrap();
    zip.write_all(custom_ui_xml.as_bytes()).unwrap();

    zip.start_file(part_name("customUI/customUI14.xml"), options)
        .unwrap();
    zip.write_all(custom_ui_xml.as_bytes()).unwrap();

    zip.start_file(part_name("customUI/_rels/customUI.xml.rels"), options)
        .unwrap();
    zip.write_all(custom_ui_rels.as_bytes()).unwrap();

    zip.start_file(part_name("customUI/image1.png"), options)
        .unwrap();
    zip.write_all(b"not-a-real-png").unwrap();

    zip.start_file(part_name("xl/vbaProject.bin"), options)
        .unwrap();
    zip.write_all(b"fake-vba-project").unwrap();

    zip.start_file(part_name("xl/_rels/vbaProject.bin.rels"), options)
        .unwrap();
    zip.write_all(vba_rels.as_bytes()).unwrap();

    zip.start_file(part_name("xl/vbaProjectSignature.bin"), options)
        .unwrap();
    zip.write_all(b"fake-signature").unwrap();

    zip.start_file(part_name("xl/vbaData.xml"), options).unwrap();
    zip.write_all(b"<vbaData/>").unwrap();

    zip.start_file(part_name("xl/activeX/activeX1.xml"), options)
        .unwrap();
    zip.write_all(activex_xml.as_bytes()).unwrap();

    zip.start_file(part_name("xl/activeX/_rels/activeX1.xml.rels"), options)
        .unwrap();
    zip.write_all(activex_rels.as_bytes()).unwrap();

    zip.start_file(part_name("xl/activeX/activeX1.bin"), options)
        .unwrap();
    zip.write_all(b"activex-binary").unwrap();

    zip.start_file(part_name("xl/embeddings/oleObject1.bin"), options)
        .unwrap();
    zip.write_all(b"ole-embedding").unwrap();

    zip.start_file(part_name("xl/ctrlProps/ctrlProp1.xml"), options)
        .unwrap();
    zip.write_all(ctrl_props_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn assert_streaming_matches_in_memory(original: &[u8]) {
    // In-memory macro stripping (reference semantics).
    let mut expected_pkg = XlsxPackage::from_bytes(original).expect("parse pkg");
    expected_pkg
        .remove_vba_project()
        .expect("strip macros in-memory");
    let expected_bytes = expected_pkg.write_to_bytes().expect("write stripped pkg");
    let expected_pkg = XlsxPackage::from_bytes(&expected_bytes).expect("parse stripped pkg");

    // Streaming macro stripping.
    let mut cursor = Cursor::new(Vec::new());
    strip_vba_project_streaming(Cursor::new(original.to_vec()), &mut cursor)
        .expect("strip macros streamingly");
    let actual_bytes = cursor.into_inner();
    let actual_pkg = XlsxPackage::from_bytes(&actual_bytes).expect("parse streaming stripped pkg");

    // Sanity: macro payload must be absent and content types must be de-macro'd.
    assert!(
        actual_pkg.vba_project_bin().is_none(),
        "expected streaming output to remove vbaProject.bin"
    );
    let ct = std::str::from_utf8(actual_pkg.part("[Content_Types].xml").unwrap()).unwrap();
    assert!(
        !ct.contains("macroEnabled.main+xml"),
        "expected streaming output to drop macroEnabled content type"
    );

    // Exact part equivalence with the in-memory reference implementation.
    let expected_names: BTreeSet<String> = expected_pkg.part_names().map(str::to_owned).collect();
    let actual_names: BTreeSet<String> = actual_pkg.part_names().map(str::to_owned).collect();
    assert_eq!(expected_names, actual_names, "part name sets differ");
    for (name, bytes) in expected_pkg.parts() {
        assert_eq!(
            Some(bytes),
            actual_pkg.part(name),
            "part differs after streaming strip: {name}"
        );
    }
}

#[test]
fn strip_vba_project_streaming_matches_in_memory_on_basic_fixture() {
    let fixture = load_basic_fixture();
    assert_streaming_matches_in_memory(&fixture);
}

#[test]
fn strip_vba_project_streaming_matches_in_memory_on_synthetic_macro_package() {
    let bytes = build_synthetic_macro_package();
    assert_streaming_matches_in_memory(&bytes);
}

#[test]
fn strip_vba_project_streaming_matches_in_memory_on_leading_slash_entries() {
    let bytes = build_synthetic_macro_package_with_leading_slash_entries();
    assert_streaming_matches_in_memory(&bytes);
}

#[test]
fn strip_vba_project_streaming_matches_in_memory_on_backslash_and_case_entries() {
    let bytes = build_synthetic_macro_package_with_backslash_and_case_entries();
    assert_streaming_matches_in_memory(&bytes);
}
