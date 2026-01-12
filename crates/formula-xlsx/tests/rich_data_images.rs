use std::io::{Cursor, Write};

use base64::Engine;
use formula_model::CellRef;
use formula_xlsx::rich_data::extract_rich_cell_images;
use formula_xlsx::XlsxPackage;
use zip::write::FileOptions;
use zip::ZipWriter;

fn png_1x1() -> Vec<u8> {
    // 1x1 transparent PNG.
    base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png")
}

fn build_rich_data_package(metadata_xml: &str) -> Vec<u8> {
    let png = png_1x1();
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

    // Cell A1 references value-metadata index 0 (`vm="0"`).
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Relationship index 0 is the first `r:id` reference in this part.
    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rv:richValueRel xmlns:rv="http://schemas.microsoft.com/office/spreadsheetml/2017/richvalue"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
   <rv:rel r:id="rId1"/>
</rv:richValueRel>"#;

    // Minimal richValue payload with a single `<rv>` record that references relationship index 0.
    // Excel's real schema is richer; this is the smallest thing our extractor understands.
    let rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rv:richValue xmlns:rv="http://schemas.microsoft.com/office/spreadsheetml/2017/richvalue">
  <rv:rv>
    <rv:v t="rel">0</rv:v>
  </rv:rv>
</rv:richValue>"#;

    let rich_value_rel_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet_xml.as_bytes()),
        ("xl/metadata.xml", metadata_xml.as_bytes()),
        ("xl/richData/richValue.xml", rich_value_xml.as_bytes()),
        (
            "xl/richData/richValueRel.xml",
            rich_value_rel_xml.as_bytes(),
        ),
        (
            "xl/richData/_rels/richValueRel.xml.rels",
            rich_value_rel_rels.as_bytes(),
        ),
        ("xl/media/image1.png", png.as_slice()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn extracts_in_cell_images_with_direct_extlst_rvb_list() {
    // This is the simpler binding form where `valueMetadata/bk/rc@v` indexes an `extLst` list of
    // `rvb @i` bindings directly (Task 233 behavior).
    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <valueMetadata>
    <bk><rc v="0"/></bk>
  </valueMetadata>
  <extLst>
    <ext uri="{DUMMY}">
      <xlrd:rvb i="0"/>
    </ext>
  </extLst>
</metadata>"#;

    let bytes = build_rich_data_package(metadata_xml);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = extract_rich_cell_images(&pkg).expect("extract richData images");

    let key = ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap());
    assert_eq!(images.get(&key).cloned(), Some(png_1x1()));
}

 #[test]
 fn extracts_in_cell_images_with_future_metadata_indirection() {
     // Excel commonly emits `metadataTypes` + `futureMetadata`, and `valueMetadata/rc` references
     // the XLRICHVALUE type via `t=` and chooses a `futureMetadata/bk` entry via `v=`.
     let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
 <metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
   <metadataTypes>
     <metadataType name="XLRICHVALUE"/>
   </metadataTypes>
   <!-- Unrelated rvb to ensure resolution uses metadataTypes/futureMetadata instead of scanning
        all rvb entries in document order. -->
   <extLst>
     <ext uri="{DUMMY}">
       <xlrd:rvb i="999"/>
     </ext>
   </extLst>
   <futureMetadata name="XLRICHVALUE">
     <bk>
       <extLst>
         <ext uri="{DUMMY}">
           <xlrd:rvb i="0"/>
         </ext>
       </extLst>
     </bk>
   </futureMetadata>
   <valueMetadata>
     <bk>
       <rc t="0" v="0"/>
     </bk>
   </valueMetadata>
 </metadata>"#;

    let bytes = build_rich_data_package(metadata_xml);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = extract_rich_cell_images(&pkg).expect("extract richData images");

    let key = ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap());
    assert_eq!(images.get(&key).cloned(), Some(png_1x1()));
}
