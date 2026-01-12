use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;
use quick_xml::events::Event;

const RELATIONSHIPS_NS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn build_macrosheet_fixture() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/macrosheets/sheet2.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>
  <Override PartName="/xl/dialogsheets/sheet3.xml" ContentType="application/vnd.ms-excel.dialogsheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="MacroSheet" sheetId="2" r:id="rId2"/>
    <sheet name="DialogSheet" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"
    Target="macrosheets/sheet2.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"
    Target="dialogsheets/sheet3.xml"/>
  <Relationship Id="rId4"
    Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"
    Target="vbaProject.bin"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let macro_sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<macroSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let dialog_sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let empty_rels = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut zip::ZipWriter<Cursor<Vec<u8>>>,
        options: zip::write::FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "[Content_Types].xml", content_types.as_bytes());
    add_file(&mut zip, options, "_rels/.rels", root_rels.as_bytes());
    add_file(&mut zip, options, "xl/workbook.xml", workbook_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/_rels/workbook.xml.rels",
        workbook_rels.as_bytes(),
    );
    add_file(&mut zip, options, "xl/worksheets/sheet1.xml", worksheet_xml.as_bytes());
    add_file(&mut zip, options, "xl/macrosheets/sheet2.xml", macro_sheet_xml.as_bytes());
    add_file(&mut zip, options, "xl/dialogsheets/sheet3.xml", dialog_sheet_xml.as_bytes());

    // Include nested relationship parts so macro stripping needs to delete them as well.
    add_file(
        &mut zip,
        options,
        "xl/macrosheets/_rels/sheet2.xml.rels",
        empty_rels,
    );
    add_file(
        &mut zip,
        options,
        "xl/dialogsheets/_rels/sheet3.xml.rels",
        empty_rels,
    );
    add_file(&mut zip, options, "xl/vbaProject.bin", b"dummy-vba");
    add_file(&mut zip, options, "xl/_rels/vbaProject.bin.rels", empty_rels);

    zip.finish().unwrap().into_inner()
}

#[test]
fn macro_stripping_removes_macrosheets_and_dialogsheets() {
    let fixture = build_macrosheet_fixture();
    let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read fixture");

    pkg.remove_vba_project().expect("strip macros");

    let written = pkg.write_to_bytes().expect("write stripped package");
    let pkg2 = XlsxPackage::from_bytes(&written).expect("read stripped package");

    assert!(pkg2.part("xl/vbaProject.bin").is_none());
    assert!(pkg2.part("xl/macrosheets/sheet2.xml").is_none());
    assert!(pkg2.part("xl/dialogsheets/sheet3.xml").is_none());
    assert!(pkg2.part("xl/macrosheets/_rels/sheet2.xml.rels").is_none());
    assert!(pkg2.part("xl/dialogsheets/_rels/sheet3.xml.rels").is_none());

    let workbook_xml = std::str::from_utf8(pkg2.part("xl/workbook.xml").unwrap())
        .expect("workbook xml utf-8");
    assert!(
        !workbook_xml.contains(r#"name="MacroSheet""#),
        "expected workbook.xml to drop macro sheet entry (got {workbook_xml:?})"
    );
    assert!(
        !workbook_xml.contains(r#"name="DialogSheet""#),
        "expected workbook.xml to drop dialog sheet entry (got {workbook_xml:?})"
    );
    assert!(
        !workbook_xml.contains(r#"r:id="rId2""#) && !workbook_xml.contains(r#"r:id="rId3""#),
        "expected workbook.xml to drop dangling r:ids (got {workbook_xml:?})"
    );

    let workbook_rels = std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap())
        .expect("workbook rels utf-8");
    assert!(
        !workbook_rels.contains("macrosheets/"),
        "expected workbook rels to stop referencing macrosheets (got {workbook_rels:?})"
    );
    assert!(
        !workbook_rels.contains("dialogsheets/"),
        "expected workbook rels to stop referencing dialogsheets (got {workbook_rels:?})"
    );

    let content_types = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap())
        .expect("content types utf-8");
    assert!(
        !content_types.contains("macroEnabled.main+xml"),
        "expected workbook content type to be downgraded to .xlsx (got {content_types:?})"
    );
    assert!(!content_types.contains("/xl/vbaProject.bin"));
    assert!(!content_types.contains("/xl/macrosheets/sheet2.xml"));
    assert!(!content_types.contains("/xl/dialogsheets/sheet3.xml"));

    // Ensure we didn't leave any dangling relationship targets or relationship id references.
    validate_opc_relationships(pkg2.parts_map()).expect("stripped package relationships are consistent");
}

fn validate_opc_relationships(parts: &BTreeMap<String, Vec<u8>>) -> Result<(), String> {
    for rels_part in parts.keys().filter(|name| name.ends_with(".rels")) {
        let Some(source_part) = source_part_from_rels_part(rels_part) else {
            continue;
        };

        if !source_part.is_empty() && !parts.contains_key(&source_part) {
            return Err(format!(
                "orphan relationship part {rels_part} (missing source {source_part})"
            ));
        }

        let xml = parts
            .get(rels_part)
            .ok_or_else(|| format!("missing rels part {rels_part}"))?;
        let ids = parse_relationship_ids(xml)?;
        let targets = parse_internal_relationship_targets(xml, &source_part)?;
        for target in targets {
            if !parts.contains_key(&target) {
                return Err(format!(
                    "relationship target {target} referenced from {rels_part} is missing"
                ));
            }
        }

        if !source_part.is_empty()
            && (source_part.ends_with(".xml") || source_part.ends_with(".vml"))
            && parts.contains_key(&source_part)
        {
            let source_xml = parts
                .get(&source_part)
                .ok_or_else(|| format!("missing source part {source_part}"))?;
            let references = parse_relationship_id_references(source_xml)?;
            for id in references {
                if !ids.contains(&id) {
                    return Err(format!(
                        "dangling relationship id {id} referenced from {source_part} (missing from {rels_part})"
                    ));
                }
            }
        }
    }

    Ok(())
}

fn parse_relationship_ids(xml: &[u8]) -> Result<BTreeSet<String>, String> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = BTreeSet::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e) | Event::Empty(ref e))
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| e.to_string())?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Id") {
                        out.insert(
                            attr.unescape_value()
                                .map_err(|e| e.to_string())?
                                .into_owned(),
                        );
                    }
                }
            }
            Ok(_) => {}
            Err(e) => return Err(e.to_string()),
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_internal_relationship_targets(xml: &[u8], source_part: &str) -> Result<Vec<String>, String> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e) | Event::Empty(ref e))
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                let mut target = None;
                let mut target_mode = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| e.to_string())?;
                    match local_name(attr.key.as_ref()) {
                        b"Target" => {
                            target = Some(
                                attr.unescape_value()
                                    .map_err(|e| e.to_string())?
                                    .into_owned(),
                            )
                        }
                        b"TargetMode" => {
                            target_mode = Some(
                                attr.unescape_value()
                                    .map_err(|e| e.to_string())?
                                    .into_owned(),
                            )
                        }
                        _ => {}
                    }
                }

                if target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }

                let Some(target) = target else {
                    continue;
                };
                let target = strip_fragment(&target);
                out.push(resolve_target(source_part, target));
            }
            Ok(_) => {}
            Err(e) => return Err(e.to_string()),
        }
        buf.clear();
    }

    Ok(out)
}

fn strip_fragment(target: &str) -> &str {
    target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target)
}

fn resolve_target(source_part: &str, target: &str) -> String {
    let (target, is_absolute) = match target.strip_prefix('/') {
        Some(target) => (target, true),
        None => (target, false),
    };

    let base_dir = if is_absolute || source_part.is_empty() {
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

fn parse_relationship_id_references(xml: &[u8]) -> Result<BTreeSet<String>, String> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = BTreeSet::new();
    let mut namespace_context = NamespaceContext::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => {
                let changes = namespace_context.apply_namespace_decls(e)?;
                collect_relationship_id_attrs(e, &namespace_context, &mut out)?;
                namespace_context.push(changes);
            }
            Ok(Event::Empty(ref e)) => {
                let changes = namespace_context.apply_namespace_decls(e)?;
                collect_relationship_id_attrs(e, &namespace_context, &mut out)?;
                namespace_context.rollback(changes);
            }
            Ok(Event::End(_)) => namespace_context.pop(),
            Ok(_) => {}
            Err(e) => return Err(e.to_string()),
        }
        buf.clear();
    }

    Ok(out)
}

fn collect_relationship_id_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    namespace_context: &NamespaceContext,
    out: &mut BTreeSet<String>,
) -> Result<(), String> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|e| e.to_string())?;
        let key = attr.key.as_ref();

        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }

        let (prefix, local) = split_prefixed_name(key);
        let namespace_uri = prefix.and_then(|p| namespace_context.namespace_for_prefix(p));
        if !is_relationship_id_attribute(namespace_uri, local) {
            continue;
        }
        out.insert(
            attr.unescape_value()
                .map_err(|e| e.to_string())?
                .into_owned(),
        );
    }

    Ok(())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|&b| b == b':').next().unwrap_or(name)
}

fn split_prefixed_name(name: &[u8]) -> (Option<&[u8]>, &[u8]) {
    match name.iter().position(|b| *b == b':') {
        Some(idx) => (Some(&name[..idx]), &name[idx + 1..]),
        None => (None, name),
    }
}

fn is_relationship_id_attribute(namespace_uri: Option<&[u8]>, local_name: &[u8]) -> bool {
    if local_name.eq_ignore_ascii_case(b"relid") {
        return true;
    }

    match namespace_uri {
        Some(ns) if ns == RELATIONSHIPS_NS => {
            local_name.eq_ignore_ascii_case(b"id")
                || local_name.eq_ignore_ascii_case(b"embed")
                || local_name.eq_ignore_ascii_case(b"link")
        }
        _ => false,
    }
}

fn source_part_from_rels_part(rels_part: &str) -> Option<String> {
    if rels_part == "_rels/.rels" {
        return Some(String::new());
    }

    if let Some(rels_file) = rels_part.strip_prefix("_rels/") {
        let rels_file = rels_file.strip_suffix(".rels")?;
        return Some(rels_file.to_string());
    }

    let (dir, rels_file) = rels_part.rsplit_once("/_rels/")?;
    let rels_file = rels_file.strip_suffix(".rels")?;

    if dir.is_empty() {
        return Some(rels_file.to_string());
    }

    Some(format!("{dir}/{rels_file}"))
}

#[derive(Debug, Default)]
struct NamespaceContext {
    /// prefix -> namespace URI
    prefixes: BTreeMap<Vec<u8>, Vec<u8>>,
    /// Stack of prefix changes for each started element that was written.
    stack: Vec<Vec<(Vec<u8>, Option<Vec<u8>>)>>,
}

impl NamespaceContext {
    fn apply_namespace_decls(
        &mut self,
        e: &quick_xml::events::BytesStart<'_>,
    ) -> Result<Vec<(Vec<u8>, Option<Vec<u8>>)>, String> {
        let mut changes = Vec::new();

        for attr in e.attributes().with_checks(false) {
            let attr = attr.map_err(|e| e.to_string())?;
            let key = attr.key.as_ref();

            // Default namespace (`xmlns="..."`) affects element names, but not attributes.
            if key == b"xmlns" {
                continue;
            }

            let Some(prefix) = key.strip_prefix(b"xmlns:") else {
                continue;
            };

            let uri = attr
                .unescape_value()
                .map_err(|e| e.to_string())?
                .into_owned()
                .into_bytes();
            let old = self.prefixes.insert(prefix.to_vec(), uri);
            changes.push((prefix.to_vec(), old));
        }

        Ok(changes)
    }

    fn rollback(&mut self, changes: Vec<(Vec<u8>, Option<Vec<u8>>)>) {
        for (prefix, old) in changes.into_iter().rev() {
            match old {
                Some(uri) => {
                    self.prefixes.insert(prefix, uri);
                }
                None => {
                    self.prefixes.remove(&prefix);
                }
            }
        }
    }

    fn push(&mut self, changes: Vec<(Vec<u8>, Option<Vec<u8>>)>) {
        self.stack.push(changes);
    }

    fn pop(&mut self) {
        if let Some(changes) = self.stack.pop() {
            self.rollback(changes);
        }
    }

    fn namespace_for_prefix(&self, prefix: &[u8]) -> Option<&[u8]> {
        self.prefixes.get(prefix).map(Vec::as_slice)
    }
}

