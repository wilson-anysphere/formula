#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};
use std::process::Command;

use formula_model::{Alignment, Range};
use formula_xlsx::write_minimal_xlsx;
use tempfile::tempdir;

fn build_synthetic_rich_data_xlsx_impl(include_rich_value_part: bool) -> Vec<u8> {
    let rich_value_override = if include_rich_value_part {
        r#"  <Override PartName="/xl/richData/richValue.xml" ContentType="application/vnd.ms-excel.richvalue+xml"/>
"#
    } else {
        ""
    };

    let content_types = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
{rich_value_override}  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
</Types>"#,
        rich_value_override = rich_value_override
    );

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // vm=1 (1-based) -> futureMetadata bk 0 -> rv index 0.
    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </valueMetadata>
</metadata>"#;

    let rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
    </rv>
  </values>
</rvData>"#;

    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>"#;

    let rich_value_rel_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    // Any bytes are fine for the CLI, but keep it a valid PNG for sanity.
    let image1_png: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00,
        0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet1_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    if include_rich_value_part {
        zip.start_file("xl/richData/richValue.xml", options)
            .unwrap();
        zip.write_all(rich_value_xml.as_bytes()).unwrap();
    }

    zip.start_file("xl/richData/richValueRel.xml", options)
        .unwrap();
    zip.write_all(rich_value_rel_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rel_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(image1_png).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_synthetic_rich_data_xlsx() -> Vec<u8> {
    build_synthetic_rich_data_xlsx_impl(true)
}

fn build_synthetic_rich_data_xlsx_without_rich_value_part() -> Vec<u8> {
    build_synthetic_rich_data_xlsx_impl(false)
}

#[test]
fn dump_rich_data_cli_prints_resolved_mapping() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_synthetic_rich_data_xlsx();
    let dir = tempdir()?;
    let path = dir.path().join("fixture.xlsx");
    std::fs::write(&path, bytes)?;

    let bin = env!("CARGO_BIN_EXE_dump_rich_data");
    let output = Command::new(bin).arg(&path).output()?;
    assert!(
        output.status.success(),
        "dump_rich_data failed: status={:?} stderr={}",
        output.status.code(),
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8(output.stdout)?;
    assert!(
        stdout.contains("Sheet1!A1 vm=1 -> rv=0 -> xl/media/image1.png"),
        "unexpected stdout:\n{stdout}"
    );

    Ok(())
}

#[test]
fn dump_rich_data_cli_prints_no_richdata_message() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = write_minimal_xlsx(&[] as &[Range], &[] as &[Alignment])?;
    let dir = tempdir()?;
    let path = dir.path().join("fixture.xlsx");
    std::fs::write(&path, bytes)?;

    let bin = env!("CARGO_BIN_EXE_dump_rich_data");
    let output = Command::new(bin).arg(&path).output()?;
    assert!(
        output.status.success(),
        "dump_rich_data failed: status={:?} stderr={}",
        output.status.code(),
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8(output.stdout)?;
    assert_eq!(stdout.trim_end(), "no richData found", "unexpected stdout:\n{stdout}");

    Ok(())
}

#[test]
fn dump_rich_data_cli_resolves_without_rich_value_parts() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = build_synthetic_rich_data_xlsx_without_rich_value_part();
    let dir = tempdir()?;
    let path = dir.path().join("fixture.xlsx");
    std::fs::write(&path, bytes)?;

    let bin = env!("CARGO_BIN_EXE_dump_rich_data");
    let output = Command::new(bin).arg(&path).output()?;
    assert!(
        output.status.success(),
        "dump_rich_data failed: status={:?} stderr={}",
        output.status.code(),
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8(output.stdout)?;
    assert!(
        stdout.contains("Sheet1!A1 vm=1 -> rv=0 -> xl/media/image1.png"),
        "unexpected stdout:\n{stdout}"
    );

    Ok(())
}
