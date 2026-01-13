use std::collections::{BTreeMap, BTreeSet};
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::{parse_relationships, resolve_relationship_target, resolve_target};
use crate::package::{XlsxError, XlsxPackage};
use crate::XlsxDocument;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PivotChartPart {
    pub part_name: String,
    pub pivot_source_name: Option<String>,
    pub pivot_source_part: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PivotChartWithPlacement {
    pub chart: PivotChartPart,
    /// Drawing parts (`xl/drawings/drawing*.xml`) that reference this chart.
    pub placed_on_drawings: Vec<String>,
    /// Worksheet or chartsheet parts (`xl/worksheets/*.xml` / `xl/chartsheets/*.xml`) that
    /// reference one of the drawing parts in `placed_on_drawings`.
    pub placed_on_sheets: Vec<String>,
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

    /// Locate pivot chart parts and also resolve where each chart is placed (drawings + sheets).
    pub fn pivot_chart_parts_with_placement(
        &self,
    ) -> Result<Vec<PivotChartWithPlacement>, XlsxError> {
        parse_pivot_chart_parts_with_placement(self)
    }
}

fn parse_pivot_chart_parts(package: &XlsxPackage) -> Result<Vec<PivotChartPart>, XlsxError> {
    parse_pivot_chart_parts_with(
        package.part_names(),
        |name| package.part(name),
        |base, rid| resolve_relationship_target(package, base, rid),
    )
}

impl XlsxDocument {
    /// Locate chart parts that declare a `<c:pivotSource>` element and resolve their pivot targets.
    pub fn pivot_chart_parts(&self) -> Result<Vec<PivotChartPart>, XlsxError> {
        parse_pivot_chart_parts_with(
            self.parts().keys(),
            |name| {
                let name = name.strip_prefix('/').unwrap_or(name);
                self.parts().get(name).map(|bytes| bytes.as_slice())
            },
            |base, rid| {
                crate::openxml::resolve_relationship_target_from_parts(
                    |name| {
                        let name = name.strip_prefix('/').unwrap_or(name);
                        self.parts().get(name).map(|bytes| bytes.as_slice())
                    },
                    base,
                    rid,
                )
            },
        )
    }
}

fn parse_pivot_chart_parts_with<'a, PN, Part, Resolve>(
    part_names: PN,
    part: Part,
    resolve_relationship_target: Resolve,
) -> Result<Vec<PivotChartPart>, XlsxError>
where
    PN: IntoIterator,
    PN::Item: AsRef<str>,
    Part: Fn(&str) -> Option<&'a [u8]>,
    Resolve: Fn(&str, &str) -> Result<Option<String>, XlsxError>,
{
    let mut chart_parts = Vec::new();
    for name in part_names {
        let name = name.as_ref();
        let name = name.strip_prefix('/').unwrap_or(name);
        if name.starts_with("xl/charts/") && name.ends_with(".xml") {
            chart_parts.push(name.to_string());
        }
    }

    let mut parts = Vec::with_capacity(chart_parts.len());
    for part_name in chart_parts {
        let xml = part(&part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.clone()))?;

        let parsed = match parse_chart_xml(xml) {
            Ok(parsed) => Some(parsed),
            Err(_) => None,
        };

        let pivot_source_part = match parsed
            .as_ref()
            .and_then(|parsed| parsed.pivot_source_rid.as_deref())
        {
            Some(rid) => resolve_relationship_target(&part_name, rid).ok().flatten(),
            None => None,
        };

        parts.push(PivotChartPart {
            part_name,
            pivot_source_name: parsed.and_then(|parsed| parsed.pivot_source_name),
            pivot_source_part,
        });
    }

    Ok(parts)
}

fn parse_pivot_chart_parts_with_placement(
    package: &XlsxPackage,
) -> Result<Vec<PivotChartWithPlacement>, XlsxError> {
    let parts = parse_pivot_chart_parts(package)?;

    let chart_to_drawings = build_chart_to_drawings_map(package)?;
    let drawing_to_sheets = build_drawing_to_sheets_map(package)?;

    let mut out = Vec::with_capacity(parts.len());
    for chart in parts {
        let placed_on_drawings = chart_to_drawings
            .get(&chart.part_name)
            .map(|set| set.iter().cloned().collect::<Vec<_>>())
            .unwrap_or_default();

        let mut placed_on_sheets: BTreeSet<String> = BTreeSet::new();
        for drawing_part in &placed_on_drawings {
            if let Some(sheets) = drawing_to_sheets.get(drawing_part) {
                placed_on_sheets.extend(sheets.iter().cloned());
            }
        }

        out.push(PivotChartWithPlacement {
            chart,
            placed_on_drawings,
            placed_on_sheets: placed_on_sheets.into_iter().collect(),
        });
    }

    Ok(out)
}

fn build_chart_to_drawings_map(
    package: &XlsxPackage,
) -> Result<BTreeMap<String, BTreeSet<String>>, XlsxError> {
    let drawing_rels = package
        .part_names()
        .filter(|name| name.starts_with("xl/drawings/_rels/") && name.ends_with(".rels"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let mut chart_to_drawings: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    for rels_name in drawing_rels {
        let rels_bytes = package
            .part(&rels_name)
            .ok_or_else(|| XlsxError::MissingPart(rels_name.clone()))?;
        let relationships = parse_relationships(rels_bytes)?;
        let drawing_part = part_name_from_rels(&rels_name, "xl/drawings/");

        for rel in relationships {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }

            let target = resolve_target(&drawing_part, &rel.target);
            if is_chart_part(&target) {
                chart_to_drawings
                    .entry(target)
                    .or_default()
                    .insert(drawing_part.clone());
            }
        }
    }

    Ok(chart_to_drawings)
}

fn build_drawing_to_sheets_map(
    package: &XlsxPackage,
) -> Result<BTreeMap<String, BTreeSet<String>>, XlsxError> {
    const REL_TYPE_DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

    let worksheet_rels = package
        .part_names()
        .filter(|name| name.starts_with("xl/worksheets/_rels/") && name.ends_with(".rels"))
        .map(str::to_string)
        .collect::<Vec<_>>();
    let chartsheet_rels = package
        .part_names()
        .filter(|name| name.starts_with("xl/chartsheets/_rels/") && name.ends_with(".rels"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let mut drawing_to_sheets: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();

    for rels_name in worksheet_rels {
        let sheet_part = part_name_from_rels(&rels_name, "xl/worksheets/");
        collect_sheet_drawings(
            package,
            &rels_name,
            &sheet_part,
            REL_TYPE_DRAWING,
            &mut drawing_to_sheets,
        )?;
    }

    for rels_name in chartsheet_rels {
        let sheet_part = part_name_from_rels(&rels_name, "xl/chartsheets/");
        collect_sheet_drawings(
            package,
            &rels_name,
            &sheet_part,
            REL_TYPE_DRAWING,
            &mut drawing_to_sheets,
        )?;
    }

    Ok(drawing_to_sheets)
}

fn collect_sheet_drawings(
    package: &XlsxPackage,
    rels_name: &str,
    sheet_part: &str,
    drawing_rel_type: &str,
    drawing_to_sheets: &mut BTreeMap<String, BTreeSet<String>>,
) -> Result<(), XlsxError> {
    let rels_bytes = package
        .part(rels_name)
        .ok_or_else(|| XlsxError::MissingPart(rels_name.to_string()))?;
    let relationships = parse_relationships(rels_bytes)?;

    for rel in relationships {
        if rel.type_uri != drawing_rel_type {
            continue;
        }
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }

        let target = resolve_target(sheet_part, &rel.target);
        if is_drawing_part(&target) {
            drawing_to_sheets
                .entry(target)
                .or_default()
                .insert(sheet_part.to_string());
        }
    }

    Ok(())
}

fn part_name_from_rels(rels_name: &str, base_dir: &str) -> String {
    // Example: `xl/drawings/_rels/drawing1.xml.rels` -> `xl/drawings/drawing1.xml`
    // Example: `xl/worksheets/_rels/sheet1.xml.rels` -> `xl/worksheets/sheet1.xml`
    let prefix = format!("{base_dir}_rels/");
    let trimmed = rels_name.strip_prefix(&prefix).unwrap_or(rels_name);
    let trimmed = trimmed.strip_suffix(".rels").unwrap_or(trimmed);
    format!("{base_dir}{trimmed}")
}

fn is_chart_part(part_name: &str) -> bool {
    part_name.starts_with("xl/charts/chart") && part_name.ends_with(".xml")
}

fn is_drawing_part(part_name: &str) -> bool {
    part_name.starts_with("xl/drawings/drawing") && part_name.ends_with(".xml")
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

    #[test]
    fn pivot_chart_parts_is_best_effort_when_chart_xml_is_malformed() {
        // Intentionally malformed chart XML so `parse_chart_xml` fails.
        let malformed_chart_xml = br#"<c:chartSpace"#;
        assert!(
            parse_chart_xml(malformed_chart_xml).is_err(),
            "expected malformed chart xml to fail parsing"
        );

        let pkg = build_package(&[("xl/charts/chart1.xml", malformed_chart_xml)]);

        let parts = pkg.pivot_chart_parts().expect("pivot chart parts");
        assert_eq!(parts.len(), 1);
        assert_eq!(parts[0].part_name, "xl/charts/chart1.xml");
        assert_eq!(parts[0].pivot_source_name, None);
        assert_eq!(parts[0].pivot_source_part, None);
    }
}
