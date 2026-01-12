use std::collections::{HashMap, HashSet};
use std::path::Path;

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
        let mut xml = std::str::from_utf8(xml_bytes)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(content_types_part.to_string(), e))?
            .to_string();

        let insert_idx = xml
            .rfind("</Types>")
            .ok_or_else(|| ChartExtractionError::XmlStructure("missing </Types>".to_string()))?;

        let mut needed_defaults: HashSet<&str> = HashSet::new();
        let mut needed_overrides: Vec<(String, String)> = Vec::new();

        for part in inserted_parts {
            let part = part.strip_prefix('/').unwrap_or(part);
            if part.ends_with(".rels") {
                continue;
            }
            let part_name = format!("/{part}");
            if let Some(content_type) = source_overrides.get(part_name.as_str()) {
                needed_overrides.push((part_name, content_type.clone()));
                continue;
            }

            if let Some(ext) = Path::new(part).extension().and_then(|ext| ext.to_str()) {
                if source_defaults.contains_key(ext) {
                    needed_defaults.insert(ext);
                }
            }
        }

        for ext in needed_defaults {
            if xml.contains(&format!("Extension=\"{ext}\"")) {
                continue;
            }
            if let Some(content_type) = source_defaults.get(ext) {
                xml.insert_str(
                    insert_idx,
                    &format!("  <Default Extension=\"{ext}\" ContentType=\"{content_type}\"/>\n"),
                );
            }
        }

        for (part_name, content_type) in needed_overrides {
            if xml.contains(&format!("PartName=\"{part_name}\"")) {
                continue;
            }
            xml.insert_str(
                insert_idx,
                &format!(
                    "  <Override PartName=\"{part_name}\" ContentType=\"{content_type}\"/>\n"
                ),
            );
        }

        self.set_part(content_types_part, xml.into_bytes());
        Ok(())
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
}
