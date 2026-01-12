use std::io::{Cursor, Write};

use formula_model::Workbook;
use formula_xlsx::{write_workbook_to_writer_with_kind, WorkbookKind, XlsxDocument, XlsxPackage};

fn sample_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Sheet1").unwrap();
    workbook
}

fn workbook_main_content_type_from_ct_xml(ct_xml: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(ct_xml).ok()?;
    for node in doc.descendants().filter(|n| n.is_element()) {
        if node.tag_name().name() != "Override" {
            continue;
        }
        if node.attribute("PartName") != Some("/xl/workbook.xml") {
            continue;
        }
        return node.attribute("ContentType").map(|s| s.to_string());
    }
    None
}

fn assert_workbook_main_content_type(bytes: &[u8], expected: &str) {
    let pkg = XlsxPackage::from_bytes(bytes).expect("read package");
    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types part"))
        .expect("utf8 content types");
    let found =
        workbook_main_content_type_from_ct_xml(ct).expect("workbook override ContentType present");
    assert_eq!(found, expected);
}

#[test]
fn writes_correct_workbook_main_content_type_for_each_kind() {
    let cases = [
        WorkbookKind::Workbook,
        WorkbookKind::MacroEnabledWorkbook,
        WorkbookKind::Template,
        WorkbookKind::MacroEnabledTemplate,
        WorkbookKind::MacroEnabledAddIn,
    ];

    for kind in cases {
        let expected = kind.workbook_content_type();

        // Simple exporter (`writer.rs`)
        let workbook = sample_workbook();
        let mut cursor = Cursor::new(Vec::new());
        write_workbook_to_writer_with_kind(&workbook, &mut cursor, kind)
            .expect("write workbook (simple)");
        assert_workbook_main_content_type(&cursor.into_inner(), expected);

        // Higher-fidelity writer (`write/*`)
        let doc = XlsxDocument::new_with_kind(sample_workbook(), kind);
        let bytes = doc.save_to_vec().expect("write workbook (doc)");
        assert_workbook_main_content_type(&bytes, expected);
    }
}

fn build_minimal_macro_enabled_template_fixture() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.template.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"
    Target="vbaProject.bin"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut zip::ZipWriter<Cursor<Vec<u8>>>,
        options: zip::write::FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(
        &mut zip,
        options,
        "[Content_Types].xml",
        content_types.as_bytes(),
    );
    add_file(&mut zip, options, "_rels/.rels", root_rels.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/workbook.xml",
        workbook_xml.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        "xl/_rels/workbook.xml.rels",
        workbook_rels.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        "xl/worksheets/sheet1.xml",
        worksheet_xml.as_bytes(),
    );

    add_file(&mut zip, options, "xl/vbaProject.bin", b"dummy-vba");
    add_file(
        &mut zip,
        options,
        "xl/_rels/vbaProject.bin.rels",
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
    );

    zip.finish().unwrap().into_inner()
}

#[test]
fn macro_stripping_xltm_to_xltx_sets_template_main_content_type() {
    let fixture = build_minimal_macro_enabled_template_fixture();
    let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read fixture");

    pkg.remove_vba_project_with_kind(WorkbookKind::Template)
        .expect("strip macros");

    let written = pkg.write_to_bytes().expect("write stripped package");
    assert_workbook_main_content_type(&written, WorkbookKind::Template.workbook_content_type());
}
