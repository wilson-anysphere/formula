use std::io::{Cursor, Write};

use formula_model::drawings::ImageId;

fn build_minimal_cellimages_xlsx() -> Vec<u8> {
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
</Relationships>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    // Minimal `xl/cellimages.xml` that references an image relationship ID.
    let cellimages_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:etc="http://schemas.microsoft.com/office/spreadsheetml/2020/11/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <etc:cellImage r:id="rId1"/>
</etc:cellImages>"#;

    let cellimages_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/cellimages.xml", options).unwrap();
    zip.write_all(cellimages_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/cellimages.xml.rels", options)
        .unwrap();
    zip.write_all(cellimages_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(b"png-bytes").unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn load_from_bytes_populates_workbook_images_from_cellimages() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = build_minimal_cellimages_xlsx();
    let doc = formula_xlsx::load_from_bytes(&bytes)?;

    assert!(
        doc.workbook
            .images
            .get(&ImageId::new("image1.png"))
            .is_some(),
        "expected Workbook.images to contain image1.png"
    );

    Ok(())
}

