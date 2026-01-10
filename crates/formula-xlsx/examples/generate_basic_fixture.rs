use std::io::{Cursor, Write};

fn main() {
    let fixture_dir = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/macros");
    std::fs::create_dir_all(&fixture_dir).expect("create fixture dir");
    let fixture_path = fixture_dir.join("basic.xlsm");

    let vba_project_bin = build_vba_project_bin();
    let xlsm_bytes = build_xlsm(&vba_project_bin);

    std::fs::write(&fixture_path, &xlsm_bytes).expect("write fixture");
    eprintln!("Wrote {}", fixture_path.display());
}

fn build_xlsm(vba_project_bin: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("ct file");
    zip.write_all(content_types_xml().as_bytes())
        .expect("write ct");

    zip.start_file("_rels/.rels", options).expect("rels");
    zip.write_all(package_rels_xml().as_bytes())
        .expect("write rels");

    zip.start_file("xl/workbook.xml", options).expect("wb");
    zip.write_all(workbook_xml().as_bytes()).expect("wb");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("wb rels");
    zip.write_all(workbook_rels_xml().as_bytes())
        .expect("wb rels");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("sheet");
    zip.write_all(sheet1_xml().as_bytes()).expect("sheet");

    zip.start_file("xl/vbaProject.bin", options)
        .expect("vbaProject.bin");
    zip.write_all(vba_project_bin)
        .expect("write vbaProject.bin");

    zip.finish().expect("finish").into_inner()
}

fn content_types_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</Types>
"#
    .to_owned()
}

fn package_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#
    .to_owned()
}

fn workbook_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#
    .to_owned()
}

fn workbook_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>
"#
    .to_owned()
}

fn sheet1_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#
    .to_owned()
}

fn build_vba_project_bin() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Root-level PROJECT stream (text).
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");

    // `VBA/dir` is a compressed container holding binary records.
    let dir_decompressed = build_dir_stream();
    let dir_compressed = make_compressed_container_literals(&dir_decompressed, true);
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_compressed).expect("write dir");
    }

    // Module stream: some header bytes then compressed source container.
    let module_code = build_module_code();
    let module_compressed = make_compressed_container_literals(module_code.as_bytes(), true);
    let header_len = 10usize;
    let mut module_stream = vec![0u8; header_len];
    module_stream.extend_from_slice(&module_compressed);
    {
        let mut s = ole.create_stream("VBA/Module1").expect("Module1 stream");
        s.write_all(&module_stream).expect("write module");
    }

    ole.into_inner().into_inner()
}

fn build_dir_stream() -> Vec<u8> {
    let mut out = Vec::new();
    // PROJECTNAME (0x0004)
    push_record(&mut out, 0x0004, b"VBAProject");
    // PROJECTCONSTANTS (0x000C)
    push_record(&mut out, 0x000C, b"");

    // MODULENAME (0x0019)
    push_record(&mut out, 0x0019, b"Module1");

    // MODULESTREAMNAME (0x001A) + reserved u16 at end.
    let mut stream_name = Vec::new();
    stream_name.extend_from_slice(b"Module1");
    stream_name.extend_from_slice(&0u16.to_le_bytes());
    push_record(&mut out, 0x001A, &stream_name);

    // MODULETYPE (0x0021): 0x0000 = standard module (in our parser)
    push_record(&mut out, 0x0021, &0u16.to_le_bytes());

    // MODULETEXTOFFSET (0x0031): points to start of compressed container in Module1 stream.
    let text_offset = 10u32;
    push_record(&mut out, 0x0031, &text_offset.to_le_bytes());
    out
}

fn build_module_code() -> String {
    [
        r#"Attribute VB_Name = "Module1""#,
        "Option Explicit",
        "",
        "Sub Hello()",
        r#"    MsgBox "Hello from VBA""#,
        "End Sub",
        "",
        "Sub WriteCells()",
        r#"    Range("A1").Value = "Written""#,
        r#"    Range("B2").Value = 42"#,
        "End Sub",
        "",
    ]
    .join("\r\n")
}

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

/// Build an MS-OVBA compressed container where the chunk is encoded as a "compressed chunk"
/// containing only literal tokens (no copy tokens). This is convenient for fixtures because
/// it exercises the compressed path without needing a full compressor.
fn make_compressed_container_literals(data: &[u8], compressed_chunk: bool) -> Vec<u8> {
    let chunk_data = if compressed_chunk {
        // flags + literals. Each flag byte controls 8 tokens (LSB-first).
        let mut out = Vec::new();
        for group in data.chunks(8) {
            out.push(0x00); // all literals
            out.extend_from_slice(group);
        }
        out
    } else {
        data.to_vec()
    };

    let size_field = (chunk_data.len() - 1) as u16;
    let header = if compressed_chunk {
        0xB000u16 | size_field
    } else {
        0x3000u16 | size_field
    };

    let mut out = Vec::new();
    out.push(0x01);
    out.extend_from_slice(&header.to_le_bytes());
    out.extend_from_slice(&chunk_data);
    out
}
