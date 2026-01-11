use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use zip::write::FileOptions;
use zip::ZipWriter;

use formula_xlsx::XlsxPackage;

fn build_fixture_xlsx() -> Vec<u8> {
    let parts: BTreeMap<String, Vec<u8>> = [
        (
            "[Content_Types].xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
</Types>
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#
            .to_vec(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>
"#
            .to_vec(),
        ),
        ("xl/embeddings/oleObject1.bin".to_string(), vec![1, 2, 3, 4]),
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

#[test]
fn preserve_drawing_parts_follows_relationships_for_embeddings() {
    let fixture_bytes = build_fixture_xlsx();
    let pkg = XlsxPackage::from_bytes(&fixture_bytes).expect("load fixture package");

    let preserved = pkg.preserve_drawing_parts().expect("preserve drawing parts");
    assert!(
        preserved.parts.contains_key("xl/embeddings/oleObject1.bin"),
        "expected embeddings part to be preserved via relationship traversal"
    );
}

