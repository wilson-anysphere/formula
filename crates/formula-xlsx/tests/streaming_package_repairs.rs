use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use formula_xlsx::{
    openxml::parse_relationships, patch_xlsx_streaming_workbook_cell_patches_with_part_overrides,
    PartOverride, WorkbookCellPatches, XlsxPackage,
};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

fn build_zip(files: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in files {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn streaming_write_repairs_macro_content_types_and_workbook_rels() -> Result<(), Box<dyn std::error::Error>>
{
    // Minimal macro-enabled workbook with broken `[Content_Types].xml` and workbook relationships:
    // - workbook main content type is macro-free
    // - `vbaProject.bin` override is missing
    // - workbook.xml.rels is missing the vbaProject relationship
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets/>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let bytes = build_zip(&[
        ("[Content_Types].xml", content_types),
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/vbaProject.bin", b"fake-vba-project"),
    ]);

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut out,
        &WorkbookCellPatches::default(),
        &HashMap::<String, PartOverride>::new(),
    )?;
    let out_bytes = out.into_inner();

    let pkg = XlsxPackage::from_bytes(&out_bytes)?;

    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())?;
    assert!(
        ct.contains("application/vnd.ms-office.vbaProject"),
        "expected macro repair to insert vbaProject.bin override in [Content_Types].xml, got:\n{ct}"
    );
    assert!(
        ct.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        "expected macro repair to upgrade workbook main content type in [Content_Types].xml, got:\n{ct}"
    );

    let rels_bytes = pkg.part("xl/_rels/workbook.xml.rels").unwrap();
    let rels = parse_relationships(rels_bytes)?;
    let vba_rel = rels
        .iter()
        .find(|rel| {
            rel.type_uri == "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
        })
        .expect("expected workbook.xml.rels to contain a vbaProject relationship");
    assert_eq!(vba_rel.target, "vbaProject.bin");

    Ok(())
}

#[test]
fn streaming_write_inserts_png_default_content_type_when_media_is_present(
) -> Result<(), Box<dyn std::error::Error>> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

    let bytes = build_zip(&[
        ("[Content_Types].xml", content_types),
        ("xl/media/image1.png", b"not-a-real-png"),
    ]);

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut out,
        &WorkbookCellPatches::default(),
        &HashMap::<String, PartOverride>::new(),
    )?;
    let out_bytes = out.into_inner();

    let pkg = XlsxPackage::from_bytes(&out_bytes)?;
    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())?;
    assert!(
        ct.contains(r#"Extension="png""#) && ct.contains(r#"ContentType="image/png""#),
        "expected streaming writer to insert png Default entry, got:\n{ct}"
    );

    Ok(())
}

#[test]
fn streaming_write_repairs_vba_signature_rels_and_content_types(
) -> Result<(), Box<dyn std::error::Error>> {
    // Like `streaming_write_repairs_macro_content_types_and_workbook_rels`, but also includes
    // `vbaProjectSignature.bin` + `vbaData.xml` and omits `xl/_rels/vbaProject.bin.rels` so the
    // streaming writer must synthesize it.
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets/>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let bytes = build_zip(&[
        ("[Content_Types].xml", content_types),
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/vbaProject.bin", b"fake-vba-project"),
        ("xl/vbaProjectSignature.bin", b"fake-vba-signature"),
        ("xl/vbaData.xml", br#"<vbaData/>"#),
    ]);

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut out,
        &WorkbookCellPatches::default(),
        &HashMap::<String, PartOverride>::new(),
    )?;
    let out_bytes = out.into_inner();

    let pkg = XlsxPackage::from_bytes(&out_bytes)?;

    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())?;
    assert!(
        ct.contains("application/vnd.ms-office.vbaProjectSignature"),
        "expected macro repair to insert vbaProjectSignature.bin override in [Content_Types].xml, got:\n{ct}"
    );
    assert!(
        ct.contains("application/vnd.ms-office.vbaData+xml"),
        "expected macro repair to insert vbaData.xml override in [Content_Types].xml, got:\n{ct}"
    );

    let rels_bytes = pkg
        .part("xl/_rels/vbaProject.bin.rels")
        .expect("expected vbaProject.bin.rels to be synthesized when signature exists");
    let rels = parse_relationships(rels_bytes)?;
    let sig_rel = rels
        .iter()
        .find(|rel| {
            rel.type_uri
                == "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"
        })
        .expect("expected vbaProject.bin.rels to contain a vbaProjectSignature relationship");
    assert_eq!(sig_rel.target, "vbaProjectSignature.bin");

    Ok(())
}

#[test]
fn streaming_write_repairs_macro_with_backslash_rels_entry_without_appending_duplicate(
) -> Result<(), Box<dyn std::error::Error>> {
    // Some producers incorrectly use `\` separators in the underlying ZIP entry names.
    // Ensure the streaming repair path patches the existing entry (instead of appending a second
    // canonical `xl/_rels/workbook.xml.rels` entry).
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets/>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let bytes = build_zip(&[
        ("[Content_Types].xml", content_types),
        ("xl/workbook.xml", workbook_xml),
        // Non-canonical ZIP entry name:
        ("xl\\_rels\\workbook.xml.rels", workbook_rels),
        ("xl/vbaProject.bin", b"fake-vba-project"),
    ]);

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut out,
        &WorkbookCellPatches::default(),
        &HashMap::<String, PartOverride>::new(),
    )?;
    let out_bytes = out.into_inner();

    // Verify the output ZIP does not contain a duplicate canonical relationships part.
    let mut zip = ZipArchive::new(Cursor::new(&out_bytes))?;
    let mut names = Vec::new();
    for i in 0..zip.len() {
        let file = zip.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        names.push(file.name().to_string());
    }
    assert!(
        names.iter().any(|n| n == "xl\\_rels\\workbook.xml.rels"),
        "expected output to preserve the original backslash entry name, got: {names:?}"
    );
    assert!(
        !names.iter().any(|n| n == "xl/_rels/workbook.xml.rels"),
        "expected output to avoid appending a canonical duplicate rels part, got: {names:?}"
    );

    // And the repaired relationships should contain the VBA relationship.
    let mut rels_bytes = Vec::new();
    zip.by_name("xl\\_rels\\workbook.xml.rels")?
        .read_to_end(&mut rels_bytes)?;
    let rels = parse_relationships(&rels_bytes)?;
    assert!(
        rels.iter().any(|rel| {
            rel.type_uri == "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
                && rel.target == "vbaProject.bin"
        }),
        "expected workbook rels to contain vbaProject relationship, got: {rels:?}"
    );

    Ok(())
}

#[test]
fn streaming_macro_repair_does_not_append_equivalent_user_override_key(
) -> Result<(), Box<dyn std::error::Error>> {
    // Regression test: if a caller supplies a PartOverride keyed by a non-canonical name (e.g.
    // backslashes), and the streaming repair layer also produces a canonical repair override for
    // the same effective part, ensure we don't append the caller override as a duplicate ZIP entry.
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets/>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let bytes = build_zip(&[
        ("[Content_Types].xml", content_types),
        ("xl/workbook.xml", workbook_xml),
        // Non-canonical ZIP entry name:
        ("xl\\_rels\\workbook.xml.rels", workbook_rels),
        ("xl/vbaProject.bin", b"fake-vba-project"),
    ]);

    // Caller override uses the *non-canonical* key. The repair layer will produce a canonical key.
    let mut overrides = HashMap::new();
    overrides.insert(
        "xl\\_rels\\workbook.xml.rels".to_string(),
        PartOverride::Replace(workbook_rels.to_vec()),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut out,
        &WorkbookCellPatches::default(),
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    // Count occurrences of the backslash entry name to ensure it wasn't appended twice.
    let mut zip = ZipArchive::new(Cursor::new(&out_bytes))?;
    let mut rel_count = 0usize;
    let mut rel_bytes = Vec::new();
    for i in 0..zip.len() {
        let mut file = zip.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        if file.name() == "xl\\_rels\\workbook.xml.rels" {
            rel_count += 1;
            // Read the (single) rels payload for later assertions.
            file.read_to_end(&mut rel_bytes)?;
        }
    }
    assert_eq!(
        rel_count, 1,
        "expected exactly one workbook rels entry, got {rel_count}"
    );

    let rels = parse_relationships(&rel_bytes)?;
    assert!(
        rels.iter().any(|rel| {
            rel.type_uri == "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
                && rel.target == "vbaProject.bin"
        }),
        "expected repaired workbook rels to contain vbaProject relationship, got: {rels:?}"
    );

    Ok(())
}

#[test]
fn streaming_part_override_matches_backslash_entry_name_without_appending_duplicate(
) -> Result<(), Box<dyn std::error::Error>> {
    let original_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let replacement_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>"#;

    let bytes = build_zip(&[("xl\\_rels\\workbook.xml.rels", original_rels)]);

    let mut overrides = HashMap::new();
    // Provide a canonical override key, while the source ZIP uses backslashes.
    overrides.insert(
        "xl/_rels/workbook.xml.rels".to_string(),
        PartOverride::Replace(replacement_rels.to_vec()),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut out,
        &WorkbookCellPatches::default(),
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    let mut zip = ZipArchive::new(Cursor::new(&out_bytes))?;
    let mut names = Vec::new();
    for i in 0..zip.len() {
        let file = zip.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        names.push(file.name().to_string());
    }
    assert!(
        names.iter().any(|n| n == "xl\\_rels\\workbook.xml.rels"),
        "expected output to preserve the original backslash entry name, got: {names:?}"
    );
    assert!(
        !names.iter().any(|n| n == "xl/_rels/workbook.xml.rels"),
        "expected output to avoid appending a canonical duplicate rels part, got: {names:?}"
    );

    let mut rels_bytes = Vec::new();
    zip.by_name("xl\\_rels\\workbook.xml.rels")?
        .read_to_end(&mut rels_bytes)?;
    let rels = parse_relationships(&rels_bytes)?;
    assert!(
        rels.iter().any(|rel| {
            rel.type_uri == "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
                && rel.target == "vbaProject.bin"
        }),
        "expected overridden workbook rels to be written into the existing backslash entry, got: {rels:?}"
    );

    Ok(())
}

#[test]
fn streaming_repair_does_not_add_image_defaults_for_removed_percent_encoded_media(
) -> Result<(), Box<dyn std::error::Error>> {
    // Ensure streaming repair logic considers the *effective* part set when deciding whether to
    // insert image `<Default>` entries into `[Content_Types].xml`.
    //
    // In particular, if the source package contains a percent-encoded image part name and the
    // caller removes it via an unescaped override key, we should not insert a png default.
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

    let bytes = build_zip(&[
        ("[Content_Types].xml", content_types),
        ("xl/media/image%201.png", b"not-a-real-png"),
    ]);

    let mut overrides = HashMap::new();
    // Remove using an unescaped (space-containing) part name; the ZIP entry is percent-encoded.
    overrides.insert("xl/media/image 1.png".to_string(), PartOverride::Remove);

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut out,
        &WorkbookCellPatches::default(),
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    // Verify the image part was removed.
    let mut zip = ZipArchive::new(Cursor::new(&out_bytes))?;
    let mut names = Vec::new();
    for i in 0..zip.len() {
        let file = zip.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        names.push(file.name().to_string());
    }
    assert!(
        !names.iter().any(|n| n.contains("image%201.png")),
        "expected removed image part to be absent, got: {names:?}"
    );

    // And streaming repair should not inject a png `<Default>` when no pngs remain.
    let pkg = XlsxPackage::from_bytes(&out_bytes)?;
    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())?;
    assert!(
        !ct.contains(r#"Extension="png""#),
        "expected no png Default entry after removing all png parts, got:\n{ct}"
    );

    Ok(())
}
