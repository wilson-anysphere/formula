use std::io::Write as _;

use formula_model::CellRef;
use formula_xlsx::XlsxPackage;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn build_zip(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip.start_file(*name, options).expect("start zip file");
        zip.write_all(bytes).expect("write zip bytes");
    }

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn rich_value_part_indices_use_numeric_suffix_ordering() {
    // Minimal workbook with one sheet (Sheet1) and a single vm-mapped cell (A1).
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    // Cell A1 uses vm="1", which is mapped in metadata.xml to rich value global index 2.
    let sheet1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"><v>#VALUE!</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // metadata.xml maps vm=1 -> rich value index 2.
    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <xlrd:rvb i="2"/>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>"#;

    // The rich value store is split across multiple parts. Each part contains a single <rv>.
    //
    // The "global index" is defined by concatenating <rv> elements across parts in numeric suffix
    // order:
    //   richValue.xml   -> idx 0
    //   richValue2.xml  -> idx 1
    //   richValue10.xml -> idx 2
    //
    // Each <rv> points at a relationship index via a <v> child.
    let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv><v>0</v></rv>
</rvData>"#;
    let rich_value2_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv><v>1</v></rv>
</rvData>"#;
    let rich_value10_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv><v>2</v></rv>
</rvData>"#;

    // Relationship index -> rId list.
    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/2022/10/richvaluerel"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
  <rel r:id="rId3"/>
</richValueRel>"#;

    // rId -> target part. Targets are resolved relative to xl/richData/richValueRel.xml.
    let rich_value_rel_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://example.com/image" Target="../media/image2.png"/>
  <Relationship Id="rId3" Type="http://example.com/image" Target="../media/image3.png"/>
</Relationships>"#;

    let image1 = b"image-1-bytes";
    let image2 = b"image-2-bytes";
    let image3 = b"image-3-bytes";

    let xlsx_bytes = build_zip(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet1_xml),
        ("xl/metadata.xml", metadata_xml),
        ("xl/richData/richValue.xml", rich_value_xml),
        ("xl/richData/richValue2.xml", rich_value2_xml),
        ("xl/richData/richValue10.xml", rich_value10_xml),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        (
            "xl/richData/_rels/richValueRel.xml.rels",
            rich_value_rel_rels_xml,
        ),
        ("xl/media/image1.png", image1),
        ("xl/media/image2.png", image2),
        ("xl/media/image3.png", image3),
    ]);

    let pkg = XlsxPackage::from_bytes(&xlsx_bytes).expect("parse xlsx bytes");
    let images = pkg
        .extract_rich_cell_images_by_cell()
        .expect("extract rich cell images");

    let key = ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap());
    assert_eq!(
        images.get(&key).map(|v| v.as_slice()),
        Some(image3.as_slice()),
        "rich value global index should be based on numeric-suffix ordering"
    );
}
