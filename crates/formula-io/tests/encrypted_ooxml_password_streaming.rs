#![cfg(feature = "encrypted-workbooks")]

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::{Cursor, Write as _};
use std::sync::atomic::{AtomicUsize, Ordering};

use formula_io::{open_workbook_model_with_options, OpenOptions};
use formula_model::CellValue;
use formula_office_crypto::{encrypt_package_to_ole, EncryptOptions};

struct TrackingAlloc;

static MAX_ALLOC: AtomicUsize = AtomicUsize::new(0);

#[global_allocator]
static GLOBAL: TrackingAlloc = TrackingAlloc;

unsafe impl GlobalAlloc for TrackingAlloc {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        track(layout.size());
        System.alloc(layout)
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        track(layout.size());
        System.alloc_zeroed(layout)
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        track(new_size);
        System.realloc(ptr, layout, new_size)
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        System.dealloc(ptr, layout)
    }
}

fn track(size: usize) {
    let mut current = MAX_ALLOC.load(Ordering::Relaxed);
    while size > current {
        match MAX_ALLOC.compare_exchange(current, size, Ordering::Relaxed, Ordering::Relaxed) {
            Ok(_) => break,
            Err(prev) => current = prev,
        }
    }
}

fn reset_max_alloc() {
    MAX_ALLOC.store(0, Ordering::Relaxed);
}

fn max_alloc() -> usize {
    MAX_ALLOC.load(Ordering::Relaxed)
}

fn build_plain_xlsx_with_filler(filler_size: usize) -> Vec<u8> {
    use zip::write::FileOptions;

    fn content_types_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"#
    }

    fn rels_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#
    }

    fn workbook_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#
    }

    fn workbook_rels_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"#
    }

    fn styles_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"#
    }

    fn worksheet_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="n"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>
"#
    }

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types_xml().as_bytes()).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(rels_xml().as_bytes()).unwrap();

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml().as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels_xml().as_bytes()).unwrap();

        zip.start_file("xl/styles.xml", options).unwrap();
        zip.write_all(styles_xml().as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml().as_bytes()).unwrap();

        zip.start_file("xl/unused/big.bin", options).unwrap();
        let mut remaining = filler_size;
        let chunk = vec![0xAAu8; 64 * 1024];
        while remaining > 0 {
            let n = remaining.min(chunk.len());
            zip.write_all(&chunk[..n]).unwrap();
            remaining -= n;
        }

        zip.finish().unwrap();
    }
    cursor.into_inner()
}

#[test]
fn opens_encrypted_xlsx_without_full_in_memory_decrypt() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsx");

    // Large unreferenced part so the decrypted ZIP file is large, but workbook parsing should not
    // need to read it.
    let plain_xlsx = build_plain_xlsx_with_filler(24 * 1024 * 1024); // 24MiB
    let mut enc_opts = EncryptOptions::default();
    enc_opts.spin_count = 1_000;
    let ole_bytes = encrypt_package_to_ole(&plain_xlsx, "password", enc_opts).expect("encrypt");
    std::fs::write(&path, &ole_bytes).expect("write encrypted workbook");

    // Drop buffers from fixture generation before measuring allocations during open.
    drop(plain_xlsx);
    drop(ole_bytes);
    reset_max_alloc();

    let workbook = open_workbook_model_with_options(
        &path,
        OpenOptions {
            password: Some("password".to_string()),
            // Keep the decrypt cache small so the allocation budget remains tight.
            encrypted_package_cache_max_bytes: Some(8 * 1024 * 1024),
        },
    )
    .expect("open encrypted workbook");

    assert_eq!(workbook.sheets.len(), 1);
    let cell = workbook.sheets[0].cell_a1("A1").unwrap().expect("A1");
    assert_eq!(cell.value, CellValue::Number(42.0));

    // If the decryption path materialized the full decrypted package, we'd expect a single large
    // allocation roughly equal to the ZIP size (~24MiB). The streaming path should not allocate
    // multi-megabyte buffers.
    assert!(
        max_alloc() < 4 * 1024 * 1024,
        "expected streaming decrypt to avoid multi-megabyte allocations; max alloc was {} bytes",
        max_alloc()
    );
}
