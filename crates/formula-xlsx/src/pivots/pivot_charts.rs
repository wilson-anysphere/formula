use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::resolve_relationship_target;
use crate::package::{XlsxError, XlsxPackage};

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PivotChartPart {
    pub part_name: String,
    pub pivot_source_name: Option<String>,
    pub pivot_source_part: Option<String>,
}

impl XlsxPackage {
    /// Locate chart parts that declare a `<c:pivotSource>` element and resolve their pivot targets.
    ///
    /// Excel stores pivot charts as normal chart parts (`xl/charts/chartN.xml`) and uses
    /// `<c:pivotSource name="..." r:id="..."/>` to bind the chart to the pivot table/cache.
    ///
    /// This parser extracts the binding information while leaving the original XML untouched for
    /// round-trip fidelity.
    pub fn pivot_chart_parts(&self) -> Result<Vec<PivotChartPart>, XlsxError> {
        parse_pivot_chart_parts(self)
    }
}

fn parse_pivot_chart_parts(package: &XlsxPackage) -> Result<Vec<PivotChartPart>, XlsxError> {
    let chart_parts = package
        .part_names()
        .filter(|name| name.starts_with("xl/charts/") && name.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let mut parts = Vec::with_capacity(chart_parts.len());
    for part_name in chart_parts {
        let xml = package
            .part(&part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.clone()))?;

        let parsed = parse_chart_xml(xml)?;
        let pivot_source_part = match parsed.pivot_source_rid.as_deref() {
            Some(rid) => match resolve_relationship_target(package, &part_name, rid) {
                Ok(target) => target,
                Err(_) => None,
            },
            None => None,
        };

        parts.push(PivotChartPart {
            part_name,
            pivot_source_name: parsed.pivot_source_name,
            pivot_source_part,
        });
    }

    Ok(parts)
}

#[derive(Debug)]
struct ParsedChartXml {
    pivot_source_name: Option<String>,
    pivot_source_rid: Option<String>,
}

fn parse_chart_xml(xml: &[u8]) -> Result<ParsedChartXml, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut pivot_source_name = None;
    let mut pivot_source_rid = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                if start.local_name().as_ref() == b"pivotSource" {
                    for attr in start.attributes().with_checks(false) {
                        let attr = attr?;
                        match attr.key.local_name().as_ref() {
                            b"name" => pivot_source_name = Some(attr.unescape_value()?.to_string()),
                            b"id" => pivot_source_rid = Some(attr.unescape_value()?.to_string()),
                            _ => {}
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(ParsedChartXml {
        pivot_source_name,
        pivot_source_rid,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn pivot_chart_parts_is_best_effort_when_chart_rels_is_malformed() {
        let chart_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:pivotSource name="PivotTable1" r:id="rId1"/>
</c:chartSpace>"#;

        // Intentionally malformed relationships XML so `openxml::parse_relationships` fails.
        let malformed_rels = br#"<Relationships"#;
        assert!(
            crate::openxml::parse_relationships(malformed_rels).is_err(),
            "expected malformed rels to fail parsing"
        );

        let pkg = build_package(&[
            ("xl/charts/chart1.xml", chart_xml),
            ("xl/charts/_rels/chart1.xml.rels", malformed_rels),
        ]);

        let parts = pkg.pivot_chart_parts().expect("pivot chart parts");
        assert_eq!(parts.len(), 1);
        assert_eq!(parts[0].part_name, "xl/charts/chart1.xml");
        assert_eq!(parts[0].pivot_source_name.as_deref(), Some("PivotTable1"));
        assert_eq!(parts[0].pivot_source_part, None);
    }
}
