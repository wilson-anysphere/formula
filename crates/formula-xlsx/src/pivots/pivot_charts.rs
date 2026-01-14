use std::collections::{BTreeMap, BTreeSet};
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::{parse_relationships, resolve_relationship_target, resolve_target};
use crate::package::{XlsxError, XlsxPackage};
use crate::sheet_metadata::parse_workbook_sheets;
use crate::XlsxDocument;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PivotChartPart {
    pub part_name: String,
    pub pivot_source_name: Option<String>,
    pub pivot_source_part: Option<String>,
    pub placed_on_drawings: Vec<String>,
    /// Sheet parts (worksheets or chartsheets) that host this chart (e.g. `xl/worksheets/sheet1.xml`,
    /// `xl/chartsheets/sheet1.xml`).
    pub placed_on_sheets: Vec<String>,
    /// Workbook sheet names for [`Self::placed_on_sheets`] when resolvable from `xl/workbook.xml`.
    pub placed_on_sheet_names: Vec<String>,
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
        // `pivot_chart_parts()` now resolves placement metadata directly on `PivotChartPart`, so this
        // API is just a convenience wrapper that avoids duplicating relationship traversal.
        let charts = self.pivot_chart_parts()?;
        let mut out = Vec::with_capacity(charts.len());
        for chart in charts {
            out.push(PivotChartWithPlacement {
                placed_on_drawings: chart.placed_on_drawings.clone(),
                placed_on_sheets: chart.placed_on_sheets.clone(),
                chart,
            });
        }
        Ok(out)
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
    let part_names: Vec<String> = part_names
        .into_iter()
        .map(|name| name.as_ref().strip_prefix('/').unwrap_or(name.as_ref()).to_string())
        .collect();

    let mut chart_parts = Vec::new();
    let mut drawing_rels = Vec::new();
    let mut worksheet_rels = Vec::new();
    let mut chartsheet_rels = Vec::new();

    for name in &part_names {
        if name.starts_with("xl/charts/") && name.ends_with(".xml") {
            chart_parts.push(name.clone());
        } else if name.starts_with("xl/drawings/_rels/") && name.ends_with(".rels") {
            drawing_rels.push(name.clone());
        } else if name.starts_with("xl/worksheets/_rels/") && name.ends_with(".rels") {
            worksheet_rels.push(name.clone());
        } else if name.starts_with("xl/chartsheets/_rels/") && name.ends_with(".rels") {
            chartsheet_rels.push(name.clone());
        }
    }

    // Ensure deterministic output and avoid duplicates when producers emit multiple equivalent
    // part names.
    chart_parts.sort();
    chart_parts.dedup();
    drawing_rels.sort();
    drawing_rels.dedup();
    worksheet_rels.sort();
    worksheet_rels.dedup();
    chartsheet_rels.sort();
    chartsheet_rels.dedup();

    let mut chart_to_drawings: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    for rels_name in drawing_rels {
        let Some(rels_bytes) = part(&rels_name) else {
            continue;
        };
        // Best-effort: malformed `.rels` parts are ignored.
        let relationships = match parse_relationships(rels_bytes) {
            Ok(relationships) => relationships,
            Err(_) => continue,
        };
        let drawing_part = drawing_part_name_from_rels(&rels_name);
        for rel in relationships {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }
            if !is_chart_relationship_type(&rel.type_uri) {
                continue;
            }
            let target = resolve_target(&drawing_part, &rel.target);
            if target.starts_with("xl/charts/") && target.ends_with(".xml") {
                chart_to_drawings
                    .entry(target)
                    .or_default()
                    .insert(drawing_part.clone());
            }
        }
    }

    let mut drawing_to_sheets: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    let mut chart_to_chartsheets: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    for rels_name in worksheet_rels.into_iter().chain(chartsheet_rels) {
        let Some(rels_bytes) = part(&rels_name) else {
            continue;
        };
        // Best-effort: malformed `.rels` parts are ignored.
        let relationships = match parse_relationships(rels_bytes) {
            Ok(relationships) => relationships,
            Err(_) => continue,
        };

        let is_worksheet = rels_name.starts_with("xl/worksheets/_rels/");
        let sheet_part = if is_worksheet {
            worksheet_part_name_from_rels(&rels_name)
        } else {
            chartsheet_part_name_from_rels(&rels_name)
        };

        for rel in relationships {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }
            let type_uri = rel.type_uri.trim();
            if is_drawing_relationship_type(type_uri) {
                let target = resolve_target(&sheet_part, &rel.target);
                if target.starts_with("xl/drawings/") {
                    drawing_to_sheets
                        .entry(target)
                        .or_default()
                        .insert(sheet_part.clone());
                }
            } else if !is_worksheet && is_chart_relationship_type(type_uri) {
                // Chartsheets can link directly to chart parts without an intermediate drawing.
                let target = resolve_target(&sheet_part, &rel.target);
                if target.starts_with("xl/charts/") && target.ends_with(".xml") {
                    chart_to_chartsheets
                        .entry(target)
                        .or_default()
                        .insert(sheet_part.clone());
                }
            }
        }
    }

    let sheet_name_by_part = sheet_name_by_part_with(&part, &resolve_relationship_target);

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

        let placed_on_drawings = chart_to_drawings
            .get(&part_name)
            .map(|drawings| drawings.iter().cloned().collect::<Vec<_>>())
            .unwrap_or_default();

        let mut placed_on_sheets: BTreeSet<String> = BTreeSet::new();
        for drawing in &placed_on_drawings {
            if let Some(sheets) = drawing_to_sheets.get(drawing) {
                placed_on_sheets.extend(sheets.iter().cloned());
            }
        }
        if let Some(sheets) = chart_to_chartsheets.get(&part_name) {
            placed_on_sheets.extend(sheets.iter().cloned());
        }
        let placed_on_sheets = placed_on_sheets.into_iter().collect::<Vec<_>>();

        let mut placed_on_sheet_names: BTreeSet<String> = BTreeSet::new();
        for sheet_part in &placed_on_sheets {
            if let Some(name) = sheet_name_by_part.get(sheet_part) {
                placed_on_sheet_names.insert(name.clone());
            }
        }
        let placed_on_sheet_names = placed_on_sheet_names.into_iter().collect::<Vec<_>>();

        parts.push(PivotChartPart {
            part_name,
            pivot_source_name: parsed.and_then(|parsed| parsed.pivot_source_name),
            pivot_source_part,
            placed_on_drawings,
            placed_on_sheets,
            placed_on_sheet_names,
        });
    }

    Ok(parts)
}

fn drawing_part_name_from_rels(rels_name: &str) -> String {
    // Example: xl/drawings/_rels/drawing1.xml.rels -> xl/drawings/drawing1.xml
    let trimmed = rels_name
        .strip_prefix("xl/drawings/_rels/")
        .unwrap_or(rels_name);
    let trimmed = trimmed.strip_suffix(".rels").unwrap_or(trimmed);
    format!("xl/drawings/{trimmed}")
}

fn worksheet_part_name_from_rels(rels_name: &str) -> String {
    // Example: xl/worksheets/_rels/sheet1.xml.rels -> xl/worksheets/sheet1.xml
    let trimmed = rels_name
        .strip_prefix("xl/worksheets/_rels/")
        .unwrap_or(rels_name);
    let trimmed = trimmed.strip_suffix(".rels").unwrap_or(trimmed);
    format!("xl/worksheets/{trimmed}")
}

fn chartsheet_part_name_from_rels(rels_name: &str) -> String {
    // Example: xl/chartsheets/_rels/sheet1.xml.rels -> xl/chartsheets/sheet1.xml
    let trimmed = rels_name
        .strip_prefix("xl/chartsheets/_rels/")
        .unwrap_or(rels_name);
    let trimmed = trimmed.strip_suffix(".rels").unwrap_or(trimmed);
    format!("xl/chartsheets/{trimmed}")
}

fn is_drawing_relationship_type(type_uri: &str) -> bool {
    // Most producers use the canonical OfficeDocument relationship URI, but some third-party
    // tools may vary the prefix. Since we only need to locate drawing parts, match by suffix.
    type_uri.trim_end().ends_with("/drawing")
}

fn is_chart_relationship_type(type_uri: &str) -> bool {
    // Most producers use the canonical OfficeDocument relationship URI, but some third-party
    // tools may vary the prefix. Since we only need to locate chart parts, match by suffix.
    type_uri.trim_end().ends_with("/chart")
}

fn sheet_name_by_part_with<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    resolve_relationship_target: &impl Fn(&str, &str) -> Result<Option<String>, XlsxError>,
) -> BTreeMap<String, String> {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = match part(workbook_part) {
        Some(bytes) => bytes,
        None => return BTreeMap::new(),
    };
    let workbook_xml = match String::from_utf8(workbook_xml.to_vec()) {
        Ok(xml) => xml,
        Err(_) => return BTreeMap::new(),
    };
    let sheets = match parse_workbook_sheets(&workbook_xml) {
        Ok(sheets) => sheets,
        Err(_) => return BTreeMap::new(),
    };

    let mut out = BTreeMap::new();
    for sheet in sheets {
        let resolved = resolve_relationship_target(workbook_part, &sheet.rel_id)
            .ok()
            .flatten()
            .or_else(|| {
                let guess_ws = format!("xl/worksheets/sheet{}.xml", sheet.sheet_id);
                if part(&guess_ws).is_some() {
                    return Some(guess_ws);
                }
                let guess_cs = format!("xl/chartsheets/sheet{}.xml", sheet.sheet_id);
                part(&guess_cs).map(|_| guess_cs)
            });
        if let Some(part) = resolved {
            out.insert(part, sheet.name);
        }
    }

    out
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
        assert!(parts[0].placed_on_drawings.is_empty());
        assert!(parts[0].placed_on_sheets.is_empty());
        assert!(parts[0].placed_on_sheet_names.is_empty());
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
