use std::io::Write as _;

use formula_model::CellRef;
use formula_xlsx::rich_data::extract_rich_cell_images;
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
fn extract_rich_value_images_accepts_plural_richvalues_part_name() {
    // `extract_rich_value_images` uses:
    //   xl/metadata.xml `<rvb i="..."/>` -> rich value global index
    //   xl/richData/richValue*.xml / richValues*.xml -> `<rv>` entries with embedded relationship IDs
    //   xl/richData/_rels/<part>.rels -> image targets
    let metadata = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rvb i="0"/>
</metadata>"#;

    let rich_values = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rv>
    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdImg"/>
  </rv>
</richValue>"#;

    let rich_values_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let image_bytes = b"fake-png-bytes";

    let xlsx_bytes = build_zip(&[
        ("xl/metadata.xml", metadata),
        ("xl/richData/richValues.xml", rich_values),
        ("xl/richData/_rels/richValues.xml.rels", rich_values_rels),
        ("xl/media/image1.png", image_bytes),
    ]);

    let pkg = XlsxPackage::from_bytes(&xlsx_bytes).expect("parse xlsx bytes");
    let extracted = pkg
        .extract_rich_value_images()
        .expect("extract rich value images");

    assert_eq!(
        extracted.images.get(&0).map(Vec::as_slice),
        Some(image_bytes.as_slice()),
        "expected plural xl/richData/richValues.xml to be discovered as a rich value store part"
    );
    assert!(
        extracted.warnings.is_empty(),
        "did not expect warnings, got: {:?}",
        extracted.warnings
    );
}

#[test]
fn extract_rich_cell_images_accepts_plural_richvalues_part_name() {
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

    // Cell A1 uses vm="0", which is mapped in metadata.xml to rich value global index 0.
    let sheet1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>#VALUE!</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // metadata.xml maps vm=0 -> rich value index 0.
    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{DUMMY}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>"#;

    // Minimal plural rich value table with one record pointing at relationship index 0.
    let rich_values_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv><v>0</v></rv>
</rvData>"#;

    // Relationship index -> rId list.
    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/2022/10/richvaluerel"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

    // rId -> target part. Targets are resolved relative to xl/richData/richValueRel.xml.
    let rich_value_rel_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/image" Target="../media/image1.png"/>
</Relationships>"#;

    let image_bytes = b"image-1-bytes";

    let xlsx_bytes = build_zip(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet1_xml),
        ("xl/metadata.xml", metadata_xml),
        ("xl/richData/richValues.xml", rich_values_xml),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        (
            "xl/richData/_rels/richValueRel.xml.rels",
            rich_value_rel_rels_xml,
        ),
        ("xl/media/image1.png", image_bytes),
    ]);

    let pkg = XlsxPackage::from_bytes(&xlsx_bytes).expect("parse xlsx bytes");
    let images = extract_rich_cell_images(&pkg).expect("extract rich cell images");

    let key = ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap());
    assert_eq!(
        images.get(&key).map(|v| v.as_slice()),
        Some(image_bytes.as_slice()),
        "expected plural xl/richData/richValues.xml to be discovered as a rich value store part"
    );
}

