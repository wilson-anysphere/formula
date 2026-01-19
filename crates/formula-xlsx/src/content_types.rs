use std::collections::{HashMap, HashSet};
use std::path::Path;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};
use roxmltree::Document;

use crate::workbook::ChartExtractionError;
use crate::XlsxPackage;

impl XlsxPackage {
    pub(crate) fn merge_content_types<'a, I>(
        &mut self,
        source_content_types_xml: &[u8],
        inserted_parts: I,
    ) -> Result<(), ChartExtractionError>
    where
        I: IntoIterator<Item = &'a String>,
    {
        let (source_defaults, source_overrides) = parse_content_types(source_content_types_xml)?;

        let content_types_part = "[Content_Types].xml";
        let xml_bytes = self
            .part(content_types_part)
            .ok_or_else(|| ChartExtractionError::MissingPart(content_types_part.to_string()))?;
        let xml_str = std::str::from_utf8(xml_bytes)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(content_types_part.to_string(), e))?
            .to_string();

        let mut needed_defaults: HashSet<String> = HashSet::new();
        let mut needed_overrides: HashMap<String, String> = HashMap::new();

        for part in inserted_parts {
            let part = part.strip_prefix('/').unwrap_or(part);
            if part.ends_with(".rels") {
                continue;
            }
            let part_name = format!("/{part}");
            if let Some(content_type) = source_overrides.get(part_name.as_str()) {
                needed_overrides.insert(part_name, content_type.clone());
                continue;
            }

            if let Some(ext) = Path::new(part).extension().and_then(|ext| ext.to_str()) {
                if source_defaults.contains_key(ext) {
                    needed_defaults.insert(ext.to_string());
                }
            }
        }

        if needed_defaults.is_empty() && needed_overrides.is_empty() {
            return Ok(());
        }

        let updated = patch_content_types_xml(
            xml_str.as_bytes(),
            content_types_part,
            &source_defaults,
            &mut needed_defaults,
            &mut needed_overrides,
        )?;
        if let Some(updated) = updated {
            self.set_part(content_types_part, updated);
        }

        Ok(())
    }
}

fn patch_content_types_xml(
    xml: &[u8],
    part_name: &str,
    source_defaults: &HashMap<String, String>,
    needed_defaults: &mut HashSet<String>,
    needed_overrides: &mut HashMap<String, String>,
) -> Result<Option<Vec<u8>>, ChartExtractionError> {
    fn local_name(name: &[u8]) -> &[u8] {
        crate::openxml::local_name(name)
    }

    fn prefixed_tag(container_name: &[u8], local: &str) -> String {
        match container_name.iter().position(|&b| b == b':') {
            Some(idx) => {
                let prefix = std::str::from_utf8(&container_name[..idx]).unwrap_or_default();
                format!("{prefix}:{local}")
            }
            None => local.to_string(),
        }
    }

    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    let Some(cap) = xml.len().checked_add(256) else {
        return Err(ChartExtractionError::AllocationFailure("patch_content_types_xml output"));
    };
    if out.try_reserve(cap).is_err() {
        return Err(ChartExtractionError::AllocationFailure("patch_content_types_xml output"));
    }
    let mut writer = XmlWriter::new(out);
    let mut buf = Vec::new();

    let mut default_tag_name: Option<String> = None;
    let mut override_tag_name: Option<String> = None;
    let mut wrote_inserts = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
        match event {
            Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Default") => {
                if default_tag_name.is_none() {
                    default_tag_name = Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Extension") {
                        let ext = attr
                            .unescape_value()
                            .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?
                            .into_owned();
                        needed_defaults.remove(ext.trim());
                    }
                }
                writer
                    .write_event(Event::Start(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Default") => {
                if default_tag_name.is_none() {
                    default_tag_name = Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Extension") {
                        let ext = attr
                            .unescape_value()
                            .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?
                            .into_owned();
                        needed_defaults.remove(ext.trim());
                    }
                }
                writer
                    .write_event(Event::Empty(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") => {
                if override_tag_name.is_none() {
                    override_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"PartName") {
                        let part = attr
                            .unescape_value()
                            .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?
                            .into_owned();
                        needed_overrides.remove(part.trim());
                    }
                }
                writer
                    .write_event(Event::Start(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") => {
                if override_tag_name.is_none() {
                    override_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"PartName") {
                        let part = attr
                            .unescape_value()
                            .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?
                            .into_owned();
                        needed_overrides.remove(part.trim());
                    }
                }
                writer
                    .write_event(Event::Empty(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::End(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Types") => {
                if !needed_defaults.is_empty() || !needed_overrides.is_empty() {
                    wrote_inserts = true;
                    let default_tag = default_tag_name
                        .clone()
                        .unwrap_or_else(|| prefixed_tag(e.name().as_ref(), "Default"));
                    let override_tag = override_tag_name
                        .clone()
                        .unwrap_or_else(|| prefixed_tag(e.name().as_ref(), "Override"));

                    for ext in needed_defaults.iter() {
                        if let Some(content_type) = source_defaults.get(ext) {
                            let mut el = BytesStart::new(default_tag.as_str());
                            el.push_attribute(("Extension", ext.as_str()));
                            el.push_attribute(("ContentType", content_type.as_str()));
                            writer
                                .write_event(Event::Empty(el))
                                .map_err(|e| {
                                    ChartExtractionError::XmlStructure(format!("{part_name}: {e}"))
                                })?;
                        }
                    }
                    for (part, content_type) in needed_overrides.iter() {
                        let mut el = BytesStart::new(override_tag.as_str());
                        el.push_attribute(("PartName", part.as_str()));
                        el.push_attribute(("ContentType", content_type.as_str()));
                        writer
                            .write_event(Event::Empty(el))
                            .map_err(|e| {
                                ChartExtractionError::XmlStructure(format!("{part_name}: {e}"))
                            })?;
                    }
                }

                writer
                    .write_event(Event::End(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Types") => {
                // Degenerate case: self-closing `<Types/>`. Expand it and insert.
                wrote_inserts = true;

                let types_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                let default_tag = default_tag_name
                    .clone()
                    .unwrap_or_else(|| prefixed_tag(types_name.as_bytes(), "Default"));
                let override_tag = override_tag_name
                    .clone()
                    .unwrap_or_else(|| prefixed_tag(types_name.as_bytes(), "Override"));

                writer
                    .write_event(Event::Start(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;

                for ext in needed_defaults.iter() {
                    if let Some(content_type) = source_defaults.get(ext) {
                        let mut el = BytesStart::new(default_tag.as_str());
                        el.push_attribute(("Extension", ext.as_str()));
                        el.push_attribute(("ContentType", content_type.as_str()));
                        writer
                            .write_event(Event::Empty(el))
                            .map_err(|e| {
                                ChartExtractionError::XmlStructure(format!("{part_name}: {e}"))
                            })?;
                    }
                }
                for (part, content_type) in needed_overrides.iter() {
                    let mut el = BytesStart::new(override_tag.as_str());
                    el.push_attribute(("PartName", part.as_str()));
                    el.push_attribute(("ContentType", content_type.as_str()));
                    writer
                        .write_event(Event::Empty(el))
                        .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
                }

                writer
                    .write_event(Event::End(BytesEnd::new(types_name.as_str())))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::Eof => break,
            other => {
                writer
                    .write_event(other)
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
        }

        buf.clear();
    }

    if wrote_inserts {
        Ok(Some(writer.into_inner()))
    } else {
        Ok(None)
    }
}

fn parse_content_types(
    xml_bytes: &[u8],
) -> Result<(HashMap<String, String>, HashMap<String, String>), ChartExtractionError> {
    let part_name = "[Content_Types].xml";
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let doc =
        Document::parse(xml).map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;

    let mut defaults = HashMap::new();
    let mut overrides = HashMap::new();

    for node in doc.descendants().filter(|n| n.is_element()) {
        match node.tag_name().name() {
            "Default" => {
                let Some(ext) = node.attribute("Extension") else {
                    continue;
                };
                let Some(content_type) = node.attribute("ContentType") else {
                    continue;
                };
                defaults.insert(ext.to_string(), content_type.to_string());
            }
            "Override" => {
                let Some(part_name) = node.attribute("PartName") else {
                    continue;
                };
                let Some(content_type) = node.attribute("ContentType") else {
                    continue;
                };
                overrides.insert(part_name.to_string(), content_type.to_string());
            }
            _ => {}
        }
    }

    Ok((defaults, overrides))
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use super::*;
    use roxmltree::Document;

    fn package_with_content_types(ct_xml: &str) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("[Content_Types].xml", options)
            .expect("start [Content_Types].xml");
        zip.write_all(ct_xml.as_bytes())
            .expect("write [Content_Types].xml");

        let bytes = zip.finish().expect("finish zip").into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read package")
    }

    #[test]
    fn merges_defaults_for_non_media_parts() {
        let destination_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

        let source_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
</Types>"#;

        let mut pkg = package_with_content_types(destination_xml);
        let inserted = vec!["xl/embeddings/oleObject1.bin".to_string()];
        pkg.merge_content_types(source_xml.as_bytes(), inserted.iter())
            .expect("merge content types");

        let updated =
            std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types part"))
                .expect("utf8 content types");
        assert!(updated.contains(
            r#"<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>"#
        ));
    }

    #[test]
    fn preserves_overrides_for_inserted_parts() {
        let destination_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

        let source_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>"#;

        let mut pkg = package_with_content_types(destination_xml);
        let inserted = vec!["xl/charts/chart1.xml".to_string()];
        pkg.merge_content_types(source_xml.as_bytes(), inserted.iter())
            .expect("merge content types");

        let updated =
            std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types part"))
                .expect("utf8 content types");
        assert!(updated.contains(
            r#"<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#
        ));
    }

    #[test]
    fn merge_content_types_tolerates_weird_types_closing_tag() {
        let destination_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types >"#;

        let source_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
</Types>"#;

        let mut pkg = package_with_content_types(destination_xml);
        let inserted = vec!["xl/embeddings/oleObject1.bin".to_string()];
        pkg.merge_content_types(source_xml.as_bytes(), inserted.iter())
            .expect("merge content types");

        let updated =
            std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types part"))
                .expect("utf8 content types");
        let doc = Document::parse(updated).expect("parse updated content types");
        assert!(
            doc.descendants().any(|n| {
                n.is_element()
                    && n.tag_name().name() == "Default"
                    && n.attribute("Extension") == Some("bin")
                    && n.attribute("ContentType")
                        == Some("application/vnd.openxmlformats-officedocument.oleObject")
            }),
            "expected bin Default to be inserted, got:\n{updated}"
        );
    }

    #[test]
    fn merge_content_types_preserves_prefix_when_root_is_prefixed() {
        let destination_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
</ct:Types>"#;

        let source_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>"#;

        let mut pkg = package_with_content_types(destination_xml);
        let inserted = vec![
            "xl/embeddings/oleObject1.bin".to_string(),
            "xl/charts/chart1.xml".to_string(),
        ];
        pkg.merge_content_types(source_xml.as_bytes(), inserted.iter())
            .expect("merge content types");

        let updated =
            std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types part"))
                .expect("utf8 content types");

        Document::parse(updated).expect("output XML should be well-formed");
        assert!(
            updated.contains(
                r#"<ct:Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>"#
            ),
            "expected prefixed <ct:Default> to be inserted, got:\n{updated}"
        );
        assert!(
            updated.contains(
                r#"<ct:Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#
            ),
            "expected prefixed <ct:Override> to be inserted, got:\n{updated}"
        );
        assert!(
            !updated.contains("<Default") && !updated.contains("<Override"),
            "should not introduce unprefixed <Default>/<Override> elements, got:\n{updated}"
        );
    }
}
