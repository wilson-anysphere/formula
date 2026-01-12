#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};
use std::process::Command;

use formula_model::{Alignment, Range};
use formula_xlsx::write_minimal_xlsx;
use tempfile::tempdir;

fn build_synthetic_rich_data_xlsx_impl(
    include_rich_value_part: bool,
    include_metadata: bool,
    include_rd_rich_value_parts: bool,
) -> Vec<u8> {
    let rich_value_override = if include_rich_value_part {
        r#"  <Override PartName="/xl/richData/richValue.xml" ContentType="application/vnd.ms-excel.richvalue+xml"/>
"#
    } else {
        ""
    };
    let metadata_override = if include_metadata {
        r#"  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
"#
    } else {
        ""
    };
    let rd_rich_value_overrides = if include_rd_rich_value_parts {
        // Match the public constants in `crates/formula-xlsx/src/rich_data/mod.rs`.
        r#"  <Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
  <Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
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
{metadata_override}{rich_value_override}{rd_rich_value_overrides}  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
</Types>"#,
        rich_value_override = rich_value_override,
        metadata_override = metadata_override,
        rd_rich_value_overrides = rd_rich_value_overrides,
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

    let metadata_rel = if include_metadata {
        r#"  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
"#
    } else {
        ""
    };
    let rd_rich_value_rels = if include_rd_rich_value_parts {
        r#"  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/>
"#
    } else {
        ""
    };

    let workbook_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
{metadata_rel}{rd_rich_value_rels}</Relationships>"#,
        metadata_rel = metadata_rel,
        rd_rich_value_rels = rd_rich_value_rels
    );

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

    let rd_rich_value_structure_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
    <k n="Text" t="s"/>
  </s>
</rvStructures>"#;

    let rd_rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0">
    <v>0</v>
    <v>6</v>
    <v>Example Alt Text</v>
  </rv>
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
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
        0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
        0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78,
        0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
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

    if include_metadata {
        zip.start_file("xl/metadata.xml", options).unwrap();
        zip.write_all(metadata_xml.as_bytes()).unwrap();
    }

    if include_rich_value_part {
        zip.start_file("xl/richData/richValue.xml", options)
            .unwrap();
        zip.write_all(rich_value_xml.as_bytes()).unwrap();
    }

    if include_rd_rich_value_parts {
        zip.start_file("xl/richData/rdrichvaluestructure.xml", options)
            .unwrap();
        zip.write_all(rd_rich_value_structure_xml.as_bytes()).unwrap();

        zip.start_file("xl/richData/rdrichvalue.xml", options)
            .unwrap();
        zip.write_all(rd_rich_value_xml.as_bytes()).unwrap();
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
    build_synthetic_rich_data_xlsx_impl(true, true, false)
}

fn build_synthetic_rich_data_xlsx_without_rich_value_part() -> Vec<u8> {
    build_synthetic_rich_data_xlsx_impl(false, true, false)
}

fn build_synthetic_rich_data_xlsx_without_metadata() -> Vec<u8> {
    build_synthetic_rich_data_xlsx_impl(false, false, false)
}

fn build_synthetic_rich_data_xlsx_with_rd_rich_value_parts() -> Vec<u8> {
    build_synthetic_rich_data_xlsx_impl(false, true, true)
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

#[test]
fn dump_rich_data_cli_resolves_without_metadata_part() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_synthetic_rich_data_xlsx_without_metadata();
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
        stdout.contains("Sheet1!A1 vm=1 -> rv=- -> xl/media/image1.png rel=0"),
        "unexpected stdout:\n{stdout}"
    );

    Ok(())
}

#[test]
fn dump_rich_data_cli_extracts_cell_images_and_writes_manifest(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_synthetic_rich_data_xlsx();
    let dir = tempdir()?;
    let path = dir.path().join("fixture.xlsx");
    std::fs::write(&path, bytes)?;

    let out_dir = dir.path().join("out");

    let bin = env!("CARGO_BIN_EXE_dump_rich_data");
    let output = Command::new(bin)
        .arg(&path)
        .arg("--extract-cell-images-out")
        .arg(&out_dir)
        .output()?;
    assert!(
        output.status.success(),
        "dump_rich_data failed: status={:?} stderr={}",
        output.status.code(),
        String::from_utf8_lossy(&output.stderr)
    );

    let image_path = out_dir.join("Sheet1_A1.png");
    assert!(
        image_path.exists(),
        "expected extracted image at {}",
        image_path.display()
    );

    let manifest_path = out_dir.join("manifest.tsv");
    let manifest = std::fs::read_to_string(&manifest_path)?;
    let mut lines = manifest.lines();

    let header = lines.next().unwrap_or_default();
    assert_eq!(
        header,
        "sheet\tcell\tbytes\tfile\timage_part\tcalc_origin\talt_text\thyperlink"
    );

    let row = lines.next().unwrap_or_default();
    let cols: Vec<&str> = row.split('\t').collect();
    assert_eq!(cols.len(), 8, "unexpected manifest row: {row}");

    assert_eq!(cols[0], "Sheet1");
    assert_eq!(cols[1], "A1");
    assert!(cols[2].parse::<usize>().is_ok(), "bytes column not numeric: {row}");
    assert_eq!(cols[3], "Sheet1_A1.png");
    assert_eq!(cols[4], "xl/media/image1.png");
    assert_eq!(cols[5], "0");
    assert_eq!(cols[6], "-");
    assert_eq!(cols[7], "-");

    Ok(())
}

#[test]
fn dump_rich_data_cli_extracts_cell_images_without_metadata(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_synthetic_rich_data_xlsx_without_metadata();
    let dir = tempdir()?;
    let path = dir.path().join("fixture.xlsx");
    std::fs::write(&path, bytes)?;

    let out_dir = dir.path().join("out");

    let bin = env!("CARGO_BIN_EXE_dump_rich_data");
    let output = Command::new(bin)
        .arg(&path)
        .arg("--extract-cell-images-out")
        .arg(&out_dir)
        .output()?;
    assert!(
        output.status.success(),
        "dump_rich_data failed: status={:?} stderr={}",
        output.status.code(),
        String::from_utf8_lossy(&output.stderr)
    );

    let image_path = out_dir.join("Sheet1_A1.png");
    assert!(
        image_path.exists(),
        "expected extracted image at {}",
        image_path.display()
    );

    let manifest_path = out_dir.join("manifest.tsv");
    let manifest = std::fs::read_to_string(&manifest_path)?;
    let mut lines = manifest.lines();

    let header = lines.next().unwrap_or_default();
    assert_eq!(
        header,
        "sheet\tcell\tbytes\tfile\timage_part\tcalc_origin\talt_text\thyperlink"
    );

    let row = lines.next().unwrap_or_default();
    let cols: Vec<&str> = row.split('\t').collect();
    assert_eq!(cols.len(), 8, "unexpected manifest row: {row}");

    assert_eq!(cols[0], "Sheet1");
    assert_eq!(cols[1], "A1");
    assert!(cols[2].parse::<usize>().is_ok(), "bytes column not numeric: {row}");
    assert_eq!(cols[3], "Sheet1_A1.png");
    assert_eq!(cols[4], "xl/media/image1.png");
    // Metadata is missing, so the extractor falls back to relationship-slot indexing without
    // rdRichValue CalcOrigin/alt text data.
    assert_eq!(cols[5], "0");
    assert_eq!(cols[6], "-");
    assert_eq!(cols[7], "-");

    Ok(())
}

#[test]
fn dump_rich_data_cli_extracts_cell_images_with_alt_text_and_calc_origin(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_synthetic_rich_data_xlsx_with_rd_rich_value_parts();
    let dir = tempdir()?;
    let path = dir.path().join("fixture.xlsx");
    std::fs::write(&path, bytes)?;

    let out_dir = dir.path().join("out");

    let bin = env!("CARGO_BIN_EXE_dump_rich_data");
    let output = Command::new(bin)
        .arg(&path)
        .arg("--extract-cell-images-out")
        .arg(&out_dir)
        .output()?;
    assert!(
        output.status.success(),
        "dump_rich_data failed: status={:?} stderr={}",
        output.status.code(),
        String::from_utf8_lossy(&output.stderr)
    );

    let manifest_path = out_dir.join("manifest.tsv");
    let manifest = std::fs::read_to_string(&manifest_path)?;
    let mut lines = manifest.lines();

    let header = lines.next().unwrap_or_default();
    assert_eq!(
        header,
        "sheet\tcell\tbytes\tfile\timage_part\tcalc_origin\talt_text\thyperlink"
    );

    let row = lines.next().unwrap_or_default();
    let cols: Vec<&str> = row.split('\t').collect();
    assert_eq!(cols.len(), 8, "unexpected manifest row: {row}");

    assert_eq!(cols[0], "Sheet1");
    assert_eq!(cols[1], "A1");
    assert!(cols[2].parse::<usize>().is_ok(), "bytes column not numeric: {row}");
    assert_eq!(cols[3], "Sheet1_A1.png");
    assert_eq!(cols[4], "xl/media/image1.png");
    assert_eq!(cols[5], "6");
    assert_eq!(cols[6], "Example Alt Text");
    assert_eq!(cols[7], "-");

    Ok(())
}
