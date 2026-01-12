use std::io::{Cursor, Read, Write};

use formula_xlsx::XlsxPackage;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[test]
fn macro_strip_drops_elements_referencing_removed_r_embed_and_r_link_relationships() -> Result<(), Box<dyn std::error::Error>>
{
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <macros>
    <pic r:embed="rId10"><child/></pic>
    <pic r:link="rId11"/>
  </macros>
  <keep r:embed="rId12"/>
</worksheet>"#;

    let worksheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="../vbaProject.bin"/>
  <Relationship Id="rId11" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="../vbaProject.bin"/>
  <Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"#;

    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#;

    let pkg_bytes = build_zip(&[
        ("xl/vbaProject.bin", b"VBA!"),
        ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
        ("xl/worksheets/_rels/sheet1.xml.rels", worksheet_rels.as_bytes()),
        ("xl/drawings/drawing1.xml", drawing_xml.as_bytes()),
    ])?;

    let mut pkg = XlsxPackage::from_bytes(&pkg_bytes)?;
    pkg.remove_vba_project()?;
    let saved = pkg.write_to_bytes()?;

    let mut archive = ZipArchive::new(Cursor::new(&saved))?;

    // Worksheet XML should lose any element referencing removed relationship IDs via r:embed/r:link.
    let mut worksheet_out = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut worksheet_out)?;
    assert!(
        !worksheet_out.contains(r#"r:embed="rId10""#),
        "expected r:embed references to removed relationships to be deleted"
    );
    assert!(
        !worksheet_out.contains(r#"r:link="rId11""#),
        "expected r:link references to removed relationships to be deleted"
    );
    assert!(
        !worksheet_out.contains("<pic"),
        "expected entire element subtrees referencing removed relationships to be dropped"
    );
    assert!(
        worksheet_out.contains(r#"r:embed="rId12""#),
        "expected unrelated relationship references to remain"
    );

    // Worksheet rels should no longer contain the removed relationship IDs.
    let mut rels_out = String::new();
    archive
        .by_name("xl/worksheets/_rels/sheet1.xml.rels")?
        .read_to_string(&mut rels_out)?;
    assert!(!rels_out.contains(r#"Id="rId10""#));
    assert!(!rels_out.contains(r#"Id="rId11""#));
    assert!(rels_out.contains(r#"Id="rId12""#));

    assert_no_missing_internal_relationship_targets(&saved)?;

    Ok(())
}

#[test]
fn macro_strip_drops_vml_elements_referencing_removed_o_relid_relationships_but_keeps_note_shapes(
) -> Result<(), Box<dyn std::error::Error>> {
    let vml = r#"<?xml version="1.0" encoding="UTF-8"?>
<v:xml xmlns:v="urn:schemas-microsoft-com:vml"
       xmlns:o="urn:schemas-microsoft-com:office:office"
       xmlns:x="urn:custom">
  <v:shape id="shape_office" o:relid="rId1"><v:textpath/></v:shape>
  <v:shape id="shape_other" x:relid="rId2"><v:textpath/></v:shape>
  <v:shape id="shape_note">
    <o:ClientData ObjectType="Note">
      <o:MoveWithCells/>
    </o:ClientData>
  </v:shape>
</v:xml>"#;

    let vml_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="../vbaProject.bin"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="../vbaProject.bin"/>
</Relationships>"#;

    let pkg_bytes = build_zip(&[
        ("xl/vbaProject.bin", b"VBA!"),
        ("xl/drawings/vmlDrawing1.vml", vml.as_bytes()),
        ("xl/drawings/_rels/vmlDrawing1.vml.rels", vml_rels.as_bytes()),
    ])?;

    let mut pkg = XlsxPackage::from_bytes(&pkg_bytes)?;
    pkg.remove_vba_project()?;
    let saved = pkg.write_to_bytes()?;

    let mut archive = ZipArchive::new(Cursor::new(&saved))?;

    let mut vml_out = String::new();
    archive
        .by_name("xl/drawings/vmlDrawing1.vml")?
        .read_to_string(&mut vml_out)?;

    assert!(
        !vml_out.contains("shape_office"),
        "expected VML shapes referencing removed o:relid relationships to be dropped"
    );
    assert!(
        !vml_out.contains("shape_other"),
        "expected VML shapes referencing removed *:relid relationships to be dropped"
    );
    assert!(
        vml_out.contains(r#"ObjectType="Note""#),
        "comment note VML shapes must remain intact when they do not reference removed IDs"
    );
    assert!(
        vml_out.contains("shape_note"),
        "comment note VML shapes must remain intact when they do not reference removed IDs"
    );

    let mut rels_out = String::new();
    archive
        .by_name("xl/drawings/_rels/vmlDrawing1.vml.rels")?
        .read_to_string(&mut rels_out)?;
    assert!(!rels_out.contains(r#"Id="rId1""#));
    assert!(!rels_out.contains(r#"Id="rId2""#));

    assert_no_missing_internal_relationship_targets(&saved)?;

    Ok(())
}

fn build_zip(parts: &[(&str, &[u8])]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip.start_file(*name, options)?;
        zip.write_all(bytes)?;
    }

    Ok(zip.finish()?.into_inner())
}

fn assert_no_missing_internal_relationship_targets(
    zip_bytes: &[u8],
) -> Result<(), Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes))?;
    let mut entries: Vec<String> = Vec::new();
    for i in 0..archive.len() {
        let file = archive.by_index(i)?;
        if file.is_file() {
            entries.push(file.name().to_string());
        }
    }

    // Re-open to read contents (ZipArchive borrows mutably per entry).
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes))?;
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if !file.is_file() || !file.name().ends_with(".rels") {
            continue;
        }

        let rels_name = file.name().to_string();
        let source_part = source_part_from_rels_name(&rels_name);
        let mut xml = String::new();
        file.read_to_string(&mut xml)?;

        let mut reader = quick_xml::Reader::from_str(&xml);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e)
                    if e.name().as_ref() == b"Relationship" =>
                {
                    let mut target = None;
                    let mut target_mode = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        match attr.key.as_ref() {
                            b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                            b"TargetMode" => target_mode = Some(attr.unescape_value()?.into_owned()),
                            _ => {}
                        }
                    }

                    if target_mode
                        .as_deref()
                        .is_some_and(|m| m.eq_ignore_ascii_case("External"))
                    {
                        continue;
                    }

                    let Some(target) = target else { continue };
                    let resolved = resolve_target_for_test(&source_part, &target);
                    assert!(
                        entries.contains(&resolved),
                        "{rels_name} has dangling relationship target {target:?} (resolved to {resolved})"
                    );
                }
                quick_xml::events::Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }

    Ok(())
}

fn source_part_from_rels_name(rels_name: &str) -> String {
    if let Some(rest) = rels_name.strip_prefix("_rels/") {
        let without_suffix = rest.strip_suffix(".rels").unwrap_or(rest);
        return without_suffix.to_string();
    }

    if let Some((dir, file)) = rels_name.rsplit_once("/_rels/") {
        let file = file.strip_suffix(".rels").unwrap_or(file);
        return if dir.is_empty() {
            file.to_string()
        } else {
            format!("{dir}/{file}")
        };
    }

    rels_name.to_string()
}

fn resolve_target_for_test(source_part: &str, target: &str) -> String {
    let (target, is_absolute) = match target.strip_prefix('/') {
        Some(target) => (target, true),
        None => (target, false),
    };
    let base_dir = if is_absolute {
        ""
    } else {
        source_part
            .rsplit_once('/')
            .map(|(dir, _)| dir)
            .unwrap_or("")
    };

    let mut components: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').collect()
    };

    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                components.pop();
            }
            _ => components.push(segment),
        }
    }

    components.join("/")
}

